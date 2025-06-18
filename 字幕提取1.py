import os
import sys
import subprocess
import webbrowser
import time
import threading
from tkinter import Tk, Label, Button, messagebox, filedialog, Frame, Scale
from tkinter import HORIZONTAL, IntVar, StringVar, Entry, Checkbutton, OptionMenu
from tkinter import Canvas, Scrollbar, VERTICAL, RIGHT, LEFT, Y, BOTH, NW

# Default paths
DEFAULT_VIDEO_PATH = "C:/Users/admin/Desktop/字幕提取/001_Installation Video_With Sub.mp4"
DEFAULT_OUTPUT_DIR = "C:/Users/admin/Desktop/字幕提取"

def check_ffmpeg():
    """Check if FFmpeg is installed"""
    try:
        subprocess.run(["ffmpeg", "-version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)
        return True
    except FileNotFoundError:
        return False

def check_dependencies():
    """Check if required dependencies are installed"""
    missing_packages = []
    
    try:
        import cv2
    except ImportError:
        missing_packages.append("opencv-python")
    
    try:
        import numpy
    except ImportError:
        missing_packages.append("numpy")
    
    try:
        import PIL
    except ImportError:
        missing_packages.append("pillow")
        
    try:
        import paddleocr
        import paddle
    except ImportError:
        missing_packages.append("paddlepaddle")
        missing_packages.append("paddleocr")
    
    return missing_packages

def install_dependencies(packages):
    """Install missing dependencies"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install"] + packages)
        return True
    except subprocess.CalledProcessError:
        return False

def download_ffmpeg():
    """Open FFmpeg download page"""
    url = "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip"
    webbrowser.open(url)
    messagebox.showinfo("Download FFmpeg", 
                       "FFmpeg download page opened. After downloading:\n"
                       "1. Extract the files\n"
                       "2. Copy ffmpeg.exe and ffprobe.exe from the bin directory to the current script directory\n"
                       "3. Restart this script")

def extract_embedded_subtitles(video_path, output_path):
    """Extract embedded subtitles"""
    try:
        # Check subtitle streams
        command = ["ffprobe", "-loglevel", "error", "-select_streams", "s", 
                  "-show_entries", "stream=index:stream_tags=language", 
                  "-of", "csv=p=0", video_path]
        
        result = subprocess.run(command, capture_output=True, text=True)
        subtitle_streams = result.stdout.strip().split('\n')
        
        if not subtitle_streams or subtitle_streams[0] == '':
            return False, "No embedded subtitles in the video"
        
        # Use the first subtitle stream
        stream_idx = subtitle_streams[0].split(',')[0]
        
        # Extract subtitles
        command = ["ffmpeg", "-i", video_path, "-map", f"0:s:{stream_idx}",
                  "-c:s", "srt", output_path, "-y"]
        
        subprocess.run(command, check=True)
        return True, f"Subtitles successfully extracted to: {output_path}"
        
    except subprocess.CalledProcessError as e:
        return False, f"Error extracting subtitles: {e}"
    except Exception as e:
        return False, f"Unknown error occurred: {e}"

def format_time(seconds):
    """Convert seconds to SRT format time code (HH:MM:SS,mmm)"""
    hours = int(seconds // 3600)
    seconds %= 3600
    minutes = int(seconds // 60)
    seconds %= 60
    milliseconds = int((seconds - int(seconds)) * 1000)
    seconds = int(seconds)
    
    return f"{hours:02d}:{minutes:02d}:{seconds:02d},{milliseconds:03d}"

def preprocess_image_for_ocr(img, options):
    """Preprocess image for better OCR results"""
    import cv2
    import numpy as np
    from PIL import Image, ImageEnhance, ImageFilter
    
    # Convert to PIL Image for some processing
    pil_img = Image.fromarray(img)
    
    # Apply enhancements if enabled
    if options.get('enhance_image', True):
        # Sharpen the image
        enhancer = ImageEnhance.Sharpness(pil_img)
        pil_img = enhancer.enhance(2.0)  # Increase sharpness
        
        # Increase contrast
        enhancer = ImageEnhance.Contrast(pil_img)
        pil_img = enhancer.enhance(1.5)  # Increase contrast
        
        # Adjust brightness if needed
        enhancer = ImageEnhance.Brightness(pil_img)
        pil_img = enhancer.enhance(1.1)  # Slight brightness increase
    
    # Convert back to OpenCV format
    img = np.array(pil_img)
    
    # Convert to grayscale if not already
    if len(img.shape) > 2:
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    else:
        gray = img.copy()
    
    # Apply special processing for subtitles
    if options.get('subtitle_mode', True):
        # Use adaptive thresholding to better handle subtitles with shadows or outlines
        if options.get('adaptive_threshold', True):
            binary = cv2.adaptiveThreshold(
                gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                cv2.THRESH_BINARY, 11, 2
            )
        else:
            # Simple thresholding for high contrast subtitles
            _, binary = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY)
        
        # If denoising is enabled
        if options.get('denoise', True):
            binary = cv2.fastNlMeansDenoising(binary, None, 10, 7, 21)
    else:
        # Standard processing for non-subtitle text
        if options.get('adaptive_threshold', True):
            binary = cv2.adaptiveThreshold(
                gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                cv2.THRESH_BINARY, 11, 2
            )
        else:
            _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        if options.get('denoise', True):
            binary = cv2.fastNlMeansDenoising(binary, None, 10, 7, 21)
    
    return img  # Return the preprocessed image, PaddleOCR works better with original images

def cleanup_text(text):
    """Clean up OCR text"""
    if not text:
        return ""
        
    # Remove empty lines and trim whitespace
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    # Filter out very short lines (likely noise)
    lines = [line for line in lines if len(line) > 1]
    
    if not lines:
        return ""
        
    return '\n'.join(lines)

def extract_hardcoded_subtitles(video_path, output_path, options, progress_callback=None):
    """Extract hardcoded subtitles using PaddleOCR"""
    try:
        import cv2
        import numpy as np
        from paddleocr import PaddleOCR
        
        # Initialize PaddleOCR with appropriate language
        lang = options.get('language', 'ch')
        use_angle = options.get('use_angle_cls', True)
        
        # Create OCR engine
        ocr = PaddleOCR(use_angle_cls=use_angle, lang=lang, 
                       det_db_thresh=0.3,          # Lower threshold to detect more text blocks
                       det_db_box_thresh=0.5,      # Confidence threshold for text boxes
                       rec_model_dir=options.get('model_dir', None))  # Optional custom model directory
        
        # Open the video
        cap = cv2.VideoCapture(video_path)
        if not cap.isOpened():
            return False, "Could not open video file"
        
        fps = cap.get(cv2.CAP_PROP_FPS)
        total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
        
        # Get video resolution
        frame_width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        frame_height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        
        # Subtitle area
        subtitle_top = options.get('subtitle_y1', int(frame_height * 0.8))
        subtitle_bottom = options.get('subtitle_y2', frame_height)
        
        # Subtitle extraction state
        last_text = ""
        subtitle_blocks = []
        current_block = {"start": 0, "text": ""}
        frame_count = 0
        sampling_interval = options.get('sampling_interval', 5)  # Sample every 5 frames
        
        # Text stability tracking
        text_buffer = []
        text_stability_threshold = options.get('text_stability', 3)  # How many consistent readings to accept text
        
        with open(output_path, "w", encoding="utf-8") as f:
            subtitles_count = 0
            
            # Set progress reporting
            progress_interval = total_frames // 100  # 1% progress interval
            if progress_interval == 0:
                progress_interval = 1
            
            while True:
                ret, frame = cap.read()
                if not ret:
                    break
                
                # Only process frames at specified interval
                if frame_count % sampling_interval == 0:
                    current_time = frame_count / fps
                    
                    # Extract subtitle region
                    subtitle_region = frame[subtitle_top:subtitle_bottom, :]
                    
                    # Enhanced preprocessing for better OCR
                    processed_region = preprocess_image_for_ocr(subtitle_region, options)
                    
                    # Perform OCR
                    ocr_result = ocr.ocr(processed_region, cls=use_angle)
                    
                    # Extract text from OCR result
                    extracted_text = ""
                    if ocr_result:
                        # PaddleOCR result structure may vary based on version
                        try:
                            # For newer versions
                            if isinstance(ocr_result, list) and len(ocr_result) > 0 and isinstance(ocr_result[0], list):
                                for line in ocr_result[0]:
                                    if len(line) >= 2:  # Has text and confidence
                                        text = line[1][0]  # Text content
                                        confidence = line[1][1]  # Confidence score
                                        
                                        # Only use text with reasonable confidence
                                        min_confidence = options.get('min_confidence', 0.7)
                                        if confidence >= min_confidence:
                                            extracted_text += text + "\n"
                            # For older versions or different return format
                            else:
                                for line in ocr_result:
                                    if isinstance(line, list) and len(line) > 0:
                                        for word_info in line:
                                            if len(word_info) >= 2:
                                                text = word_info[1][0]
                                                confidence = word_info[1][1]
                                                min_confidence = options.get('min_confidence', 0.7)
                                                if confidence >= min_confidence:
                                                    extracted_text += text + "\n"
                        except Exception as e:
                            print(f"Error processing OCR result: {e}")
                            # Fallback for unexpected result format
                            extracted_text = str(ocr_result)
                    
                    # Clean up the text
                    text = cleanup_text(extracted_text)
                    
                    # Implement text stability check
                    if text:
                        # Add to buffer
                        text_buffer.append(text)
                        
                        # Keep only the most recent samples
                        if len(text_buffer) > text_stability_threshold:
                            text_buffer.pop(0)
                        
                        # Check if the text is stable
                        if len(text_buffer) == text_stability_threshold:
                            most_common_text = max(set(text_buffer), key=text_buffer.count)
                            # Only accept if it appears in majority of frames
                            if text_buffer.count(most_common_text) >= text_stability_threshold // 2 + 1:
                                # Text is stable, use it if different from last accepted text
                                stable_text = most_common_text
                                if stable_text != last_text:
                                    # If already have a subtitle block, end current block
                                    if current_block["text"]:
                                        current_block["end"] = current_time
                                        subtitle_blocks.append(current_block)
                                        
                                        # Write in SRT format
                                        subtitles_count += 1
                                        f.write(f"{subtitles_count}\n")
                                        f.write(f"{format_time(current_block['start'])} --> {format_time(current_block['end'])}\n")
                                        f.write(f"{current_block['text']}\n\n")
                                    
                                    # Start new subtitle block
                                    current_block = {"start": current_time, "text": stable_text}
                                    last_text = stable_text
                    elif len(text_buffer) > 0:
                        # Clear buffer if we've had consecutive empty frames
                        text_buffer.pop(0)
                        if not text_buffer:  # If buffer is now empty, end current subtitle
                            if current_block["text"]:
                                current_block["end"] = current_time
                                subtitle_blocks.append(current_block)
                                
                                # Write in SRT format
                                subtitles_count += 1
                                f.write(f"{subtitles_count}\n")
                                f.write(f"{format_time(current_block['start'])} --> {format_time(current_block['end'])}\n")
                                f.write(f"{current_block['text']}\n\n")
                                
                                # Reset current block
                                current_block = {"start": 0, "text": ""}
                                last_text = ""
                
                # Update progress
                if frame_count % progress_interval == 0 and progress_callback:
                    percent = min(100, int((frame_count * 100) / total_frames))
                    progress_callback(percent)
                
                frame_count += 1
            
            # Process the last subtitle block
            if current_block["text"]:
                current_block["end"] = frame_count / fps
                subtitle_blocks.append(current_block)
                
                subtitles_count += 1
                f.write(f"{subtitles_count}\n")
                f.write(f"{format_time(current_block['start'])} --> {format_time(current_block['end'])}\n")
                f.write(f"{current_block['text']}\n\n")
        
        cap.release()
        return True, f"PaddleOCR extracted {subtitles_count} subtitles and saved to: {output_path}"
    
    except Exception as e:
        import traceback
        return False, f"Error extracting subtitles with PaddleOCR: {str(e)}\n{traceback.format_exc()}"

def extract_subtitles(video_path, output_dir, ocr_options=None, progress_callback=None):
    """Main function to extract subtitles, first try embedded, then OCR if fails"""
    if not os.path.exists(video_path):
        return False, f"Video file does not exist: {video_path}"
    
    # Check if FFmpeg is available
    if not check_ffmpeg():
        return False, "FFmpeg not found, cannot continue"
    
    # Prepare output path
    base_name = os.path.splitext(os.path.basename(video_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.srt")
    
    # First try to extract embedded subtitles
    success, message = extract_embedded_subtitles(video_path, output_path)
    if success:
        return True, message
    
    # If embedded subtitle extraction fails, and OCR is enabled, try OCR
    if ocr_options and ocr_options.get('enable_ocr', True):
        try:
            from paddleocr import PaddleOCR
            return extract_hardcoded_subtitles(video_path, output_path, ocr_options, progress_callback)
        except ImportError:
            return False, "PaddleOCR not installed, cannot perform OCR recognition"
    
    return False, message

def preview_subtitle_area(video_path, y1, y2):
    """Preview subtitle area"""
    try:
        import cv2
        import numpy as np
        
        cap = cv2.VideoCapture(video_path)
        if not cap.isOpened():
            messagebox.showerror("Error", "Could not open video file")
            return
        
        # Jump to middle of video
        total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
        cap.set(cv2.CAP_PROP_POS_FRAMES, total_frames // 2)
        
        ret, frame = cap.read()
        if not ret:
            messagebox.showerror("Error", "Could not read video frame")
            cap.release()
            return
        
        # Draw subtitle area
        preview_frame = frame.copy()
        cv2.rectangle(preview_frame, (0, y1), (frame.shape[1], y2), (0, 255, 0), 2)
        
        # Save preview image to temp file
        temp_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "subtitle_preview.jpg")
        cv2.imwrite(temp_file, preview_frame)
        
        # Show preview
        if sys.platform == 'win32':
            os.startfile(temp_file)
        else:
            import subprocess
            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.call([opener, temp_file])
        
        cap.release()
    
    except Exception as e:
        messagebox.showerror("Error", f"Error previewing subtitle area: {e}")

class ScrollableFrame(Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = Canvas(self)
        scrollbar = Scrollbar(self, orient=VERTICAL, command=canvas.yview)
        self.scrollable_frame = Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor=NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

class SubtitleExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PaddleOCR Subtitle Extractor")
        self.root.geometry("600x700")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Create scrollable frame
        self.scroll_container = ScrollableFrame(root)
        self.scroll_container.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # Main frame is now the scrollable frame
        self.frame = self.scroll_container.scrollable_frame
        
        # Progress variables
        self.progress_var = StringVar(value="Ready")
        self.progress_percent = IntVar(value=0)
        
        # Check dependencies
        missing_packages = check_dependencies()
        if missing_packages:
            if messagebox.askyesno("Missing Dependencies", 
                                 f"The following Python packages need to be installed: {', '.join(missing_packages)}\nInstall now?"):
                if install_dependencies(missing_packages):
                    messagebox.showinfo("Installation Successful", "Dependencies installed successfully, please restart the program")
                else:
                    messagebox.showerror("Installation Failed", "Failed to install dependencies, please install manually")
            self.root.destroy()
            return
        
        # Check FFmpeg
        ffmpeg_available = check_ffmpeg()
        
        if not ffmpeg_available:
            Label(self.frame, text="FFmpeg not detected, please download and install first", fg="red", font=("Arial", 12)).pack(pady=10)
            Button(self.frame, text="Download FFmpeg", command=download_ffmpeg, width=20).pack(pady=5)
            Label(self.frame, text="Restart this program after installation", font=("Arial", 10)).pack(pady=5)
        else:
            Label(self.frame, text="FFmpeg is installed ✓", fg="green", font=("Arial", 12)).pack(pady=5)
            
            # Title
            Label(self.frame, text="PaddleOCR Subtitle Extractor", font=("Arial", 16, "bold")).pack(pady=10)
            Label(self.frame, text="Optimized for Chinese and English subtitles", font=("Arial", 10)).pack(pady=5)
            
            # Subtitle extraction options
            Label(self.frame, text="Subtitle Extraction Options", font=("Arial", 14, "bold")).pack(pady=10)
            
            # OCR enable option
            self.ocr_var = IntVar(value=1)
            button_frame = Frame(self.frame)
            button_frame.pack(pady=10)
            Button(button_frame, text="Select Video File", command=self.select_video, 
                   width=20, font=("Arial", 11)).pack(side=LEFT, padx=5)
            Button(button_frame, text="Use Default Video Path", command=self.use_default, 
                   width=20, font=("Arial", 11)).pack(side=LEFT, padx=5)
            
            # Separator
            Frame(self.frame, height=2, bg="gray").pack(fill="both", pady=15)
            
            # OCR settings
            Label(self.frame, text="PaddleOCR Settings", font=("Arial", 12, "bold")).pack(pady=5)
            Checkbutton(self.frame, text="Enable OCR for hardcoded subtitles (if no embedded subtitles found)", 
                       variable=self.ocr_var, font=("Arial", 10)).pack(anchor="w", pady=5)
            
            # Language selection
            lang_frame = Frame(self.frame)
            lang_frame.pack(fill="both", pady=5)
            Label(lang_frame, text="OCR Language:", width=15).pack(side=LEFT)
            self.lang_var = StringVar(value="ch")
            lang_options = [
                "ch",        # Chinese (Simplified + Traditional)
                "en",        # English
                "chinese_cht",  # Chinese Traditional
                "korean",    # Korean
                "japan",     # Japanese
                "latin"      # Latin
            ]
            OptionMenu(lang_frame, self.lang_var, *lang_options).pack(side=LEFT, fill="both", expand=True)
            
            # Use angle classifier
            angle_frame = Frame(self.frame)
            angle_frame.pack(fill="both", pady=5)
            self.angle_var = IntVar(value=1)
            Checkbutton(angle_frame, text="Use Angle Classifier (helps with rotated text)", 
                       variable=self.angle_var).pack(side=LEFT)
            
            # Confidence threshold
            conf_frame = Frame(self.frame)
            conf_frame.pack(fill="both", pady=5)
            Label(conf_frame, text="Min Confidence:", width=15).pack(side=LEFT)
            self.conf_var = StringVar(value="0.7")
            conf_options = ["0.5", "0.6", "0.7", "0.8", "0.9"]
            OptionMenu(conf_frame, self.conf_var, *conf_options).pack(side=LEFT, fill="both", expand=True)
            
            # Sampling frequency
            sampling_frame = Frame(self.frame)
            sampling_frame.pack(fill="both", pady=5)
            Label(sampling_frame, text="Sampling Interval:", width=15).pack(side=LEFT)
            self.sampling_var = IntVar(value=5)  # Every 5 frames
            Scale(sampling_frame, from_=1, to=30, variable=self.sampling_var, 
                 orient=HORIZONTAL, length=300).pack(side=LEFT, fill="both", expand=True)
            
            # Subtitle area setting
            Label(self.frame, text="Subtitle Area Settings (bottom area of video)", font=("Arial", 12, "bold")).pack(pady=5)
            
            # Get video height (if default video exists)
            self.video_height = 720  # Default height
            if os.path.exists(DEFAULT_VIDEO_PATH):
                try:
                    import cv2
                    cap = cv2.VideoCapture(DEFAULT_VIDEO_PATH)
                    self.video_height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
                    cap.release()
                except:
                    pass
            
            # Upper boundary slider
            y1_frame = Frame(self.frame)
            y1_frame.pack(fill="both", pady=5)
            Label(y1_frame, text="Upper Boundary:", width=15).pack(side=LEFT)
            self.y1_var = IntVar(value=int(self.video_height * 0.8))  # Default: top of bottom 20% area
            self.y1_scale = Scale(y1_frame, from_=0, to=self.video_height, variable=self.y1_var, 
                            orient=HORIZONTAL, length=300)
            self.y1_scale.pack(side=LEFT, fill="both", expand=True)
            
            # Lower boundary slider
            y2_frame = Frame(self.frame)
            y2_frame.pack(fill="both", pady=5)
            Label(y2_frame, text="Lower Boundary:", width=15).pack(side=LEFT)
            self.y2_var = IntVar(value=self.video_height)  # Default: bottom of video
            self.y2_scale = Scale(y2_frame, from_=0, to=self.video_height, variable=self.y2_var, 
                            orient=HORIZONTAL, length=300)
            self.y2_scale.pack(side=LEFT, fill="both", expand=True)
            
            # Image processing options
            process_frame = Frame(self.frame)
            process_frame.pack(fill="both", pady=10)
            Label(process_frame, text="Image Processing:", font=("Arial", 12, "bold")).pack(anchor="w")
            
            # Create a sub-frame for checkboxes
            check_frame = Frame(self.frame)
            check_frame.pack(fill="both", pady=5)
            
            # Image enhancement options
            self.enhance_var = IntVar(value=1)
            self.denoise_var = IntVar(value=1)
            self.adaptive_var = IntVar(value=1)
            self.subtitle_mode_var = IntVar(value=1)
            
            Checkbutton(check_frame, text="Enhance Image", variable=self.enhance_var).pack(anchor="w")
            Checkbutton(check_frame, text="Denoise", variable=self.denoise_var).pack(anchor="w")
            Checkbutton(check_frame, text="Adaptive Threshold", variable=self.adaptive_var).pack(anchor="w")
            Checkbutton(check_frame, text="Subtitle Optimization Mode", variable=self.subtitle_mode_var).pack(anchor="w")
            
            # Text stability threshold
            stability_frame = Frame(self.frame)
            stability_frame.pack(fill="both", pady=5)
            Label(stability_frame, text="Text Stability:", width=15).pack(side=LEFT)
            self.stability_var = IntVar(value=3)  # Default: require 3 consistent readings
            Scale(stability_frame, from_=1, to=10, variable=self.stability_var, 
                 orient=HORIZONTAL, length=300).pack(side=LEFT, fill="both", expand=True)
            
            # Preview and test buttons
            button_frame2 = Frame(self.frame)
            button_frame2.pack(pady=10)
            Button(button_frame2, text="Preview Subtitle Area", command=self.preview_area, width=20).pack(side=LEFT, padx=5)
            Button(button_frame2, text="Test OCR on Current Frame", command=self.test_ocr_on_frame, width=25).pack(side=LEFT, padx=5)
            
            # Progress display
            progress_frame = Frame(self.frame)
            progress_frame.pack(fill="both", pady=10)
            Label(progress_frame, textvariable=self.progress_var, font=("Arial", 10)).pack()
            
            # Tips
            Label(self.frame, text="Tips: First set the subtitle area, then click 'Preview' to confirm it's correct, and 'Test OCR' to verify accuracy", 
                  font=("Arial", 9), fg="gray").pack(pady=5)
    
    def preview_area(self):
        """Preview the selected subtitle area"""
        # First check if default video exists
        if not os.path.exists(DEFAULT_VIDEO_PATH):
            file_path = filedialog.askopenfilename(
                title="Select Video File for Preview",
                filetypes=[("Video Files", "*.mp4 *.mkv *.avi *.mov"), ("All Files", "*.*")]
            )
            if not file_path:
                return
        else:
            file_path = DEFAULT_VIDEO_PATH
        
        # Show preview
        preview_subtitle_area(file_path, self.y1_var.get(), self.y2_var.get())
    
    def test_ocr_on_frame(self):
        """Test OCR on a single frame from the video to verify settings"""
        try:
            import cv2
            import numpy as np
            from paddleocr import PaddleOCR
            
            # First check if default video exists
            if not os.path.exists(DEFAULT_VIDEO_PATH):
                file_path = filedialog.askopenfilename(
                    title="Select Video File for OCR Test",
                    filetypes=[("Video Files", "*.mp4 *.mkv *.avi *.mov"), ("All Files", "*.*")]
                )
                if not file_path:
                    return
            else:
                file_path = DEFAULT_VIDEO_PATH
            
            # Initialize PaddleOCR with appropriate settings
            lang = self.lang_var.get()
            use_angle = self.angle_var.get() == 1
            
            # Update status and show loading message
            self.progress_var.set("Initializing PaddleOCR, please wait...")
            self.root.update()
            
            ocr = PaddleOCR(use_angle_cls=use_angle, lang=lang,
                           det_db_thresh=0.3,
                           det_db_box_thresh=0.5)
            
            # Open video
            cap = cv2.VideoCapture(file_path)
            if not cap.isOpened():
                messagebox.showerror("Error", "Could not open video file")
                return
            
            # Jump to middle of video
            total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
            cap.set(cv2.CAP_PROP_POS_FRAMES, total_frames // 2)
            
            # Read frame
            ret, frame = cap.read()
            if not ret:
                messagebox.showerror("Error", "Could not read video frame")
                cap.release()
                return
            
            # Get subtitle area
            y1 = self.y1_var.get()
            y2 = self.y2_var.get()
            subtitle_region = frame[y1:y2, :]
            
            # Create options dict from GUI settings
            options = {
                'enhance_image': self.enhance_var.get() == 1,
                'denoise': self.denoise_var.get() == 1,
                'adaptive_threshold': self.adaptive_var.get() == 1,
                'subtitle_mode': self.subtitle_mode_var.get() == 1
            }
            
            # Process image
            processed_region = preprocess_image_for_ocr(subtitle_region, options)
            
            # Display processing message
            self.progress_var.set("Performing OCR, please wait...")
            self.root.update()
            
            # Perform OCR
            ocr_result = ocr.ocr(processed_region, cls=use_angle)
            
            # Extract and format text
            extracted_text = ""
            if ocr_result:
                # Handle different PaddleOCR result formats
                try:
                    # For newer versions
                    if isinstance(ocr_result, list) and len(ocr_result) > 0 and isinstance(ocr_result[0], list):
                        for line in ocr_result[0]:
                            if len(line) >= 2:  # Has text and confidence
                                text = line[1][0]  # Text content
                                confidence = line[1][1]  # Confidence score
                                
                                # Only use text with reasonable confidence
                                min_confidence = float(self.conf_var.get())
                                if confidence >= min_confidence:
                                    extracted_text += f"{text} [{confidence:.2f}]\n"
                    # For older versions or different return format
                    else:
                        for line in ocr_result:
                            if isinstance(line, list) and len(line) > 0:
                                for word_info in line:
                                    if len(word_info) >= 2:
                                        text = word_info[1][0]
                                        confidence = word_info[1][1]
                                        min_confidence = float(self.conf_var.get())
                                        if confidence >= min_confidence:
                                            extracted_text += f"{text} [{confidence:.2f}]\n"
                except Exception as e:
                    messagebox.showerror("OCR Processing Error", f"Error processing OCR result: {e}")
                    extracted_text = str(ocr_result)
            
            # Clean up the text
            text = cleanup_text(extracted_text)
            
            # Save both original and processed images for comparison
            temp_dir = os.path.dirname(os.path.abspath(__file__))
            orig_file = os.path.join(temp_dir, "subtitle_orig.jpg")
            proc_file = os.path.join(temp_dir, "subtitle_processed.jpg")
            
            cv2.imwrite(orig_file, subtitle_region)
            cv2.imwrite(proc_file, processed_region)
            
            # Create visualization of detected text areas
            vis_img = subtitle_region.copy()
            
            # Draw bounding boxes if available
            try:
                if ocr_result and isinstance(ocr_result, list) and len(ocr_result) > 0:
                    # Different handling based on result format
                    if isinstance(ocr_result[0], list):
                        for line in ocr_result[0]:
                            if isinstance(line, list) and len(line) > 0:
                                if isinstance(line[0], list) and len(line[0]) == 4:  # Has coordinates
                                    box = line[0]
                                    pts = np.array(box, np.int32).reshape((-1, 1, 2))
                                    cv2.polylines(vis_img, [pts], True, (0, 255, 0), 2)
                                    if len(line) > 1 and isinstance(line[1], list) and len(line[1]) > 0:
                                        text = line[1][0]
                                        cv2.putText(vis_img, text, (int(box[0][0]), int(box[0][1])-10), 
                                                  cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 255), 1)
            except Exception as e:
                # Just continue if visualization fails
                print(f"Visualization error: {e}")
            
            vis_file = os.path.join(temp_dir, "subtitle_detection.jpg")
            cv2.imwrite(vis_file, vis_img)
            
            # Show result
            result = f"Detected Text:\n{text}\n\nImages saved to:\n{orig_file} (original)\n{proc_file} (processed)\n{vis_file} (detection visualization)"
            
            # Reset progress message
            self.progress_var.set("Ready")
            
            messagebox.showinfo("OCR Test Result", result)
            
            # Open the visualization image
            if sys.platform == 'win32':
                os.startfile(vis_file)
            
            cap.release()
            
        except Exception as e:
            messagebox.showerror("Error", f"OCR test failed: {e}")
            self.progress_var.set("Test failed")
    
    def select_video(self):
        """Open file dialog to select video"""
        filename = filedialog.askopenfilename(
            title="Select Video File",
            filetypes=[("Video Files", "*.mp4 *.mkv *.avi *.mov"), ("All Files", "*.*")]
        )
        if filename:
            output_dir = os.path.dirname(filename)
            self.process_video(filename, output_dir)
    
    def use_default(self):
        """Use default video path"""
        if not os.path.exists(DEFAULT_VIDEO_PATH):
            messagebox.showerror("Error", f"Default video path not found: {DEFAULT_VIDEO_PATH}")
            return
        self.process_video(DEFAULT_VIDEO_PATH, DEFAULT_OUTPUT_DIR)
    
    def process_video(self, video_path, output_dir):
        """Process video to extract subtitles"""
        if not os.path.exists(video_path):
            messagebox.showerror("Error", f"Video file does not exist: {video_path}")
            return
        
        # Create output directory if it doesn't exist
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Set OCR options from GUI settings
        ocr_options = {
            'enable_ocr': self.ocr_var.get() == 1,
            'subtitle_y1': self.y1_var.get(),
            'subtitle_y2': self.y2_var.get(),
            'language': self.lang_var.get(),
            'use_angle_cls': self.angle_var.get() == 1,
            'sampling_interval': self.sampling_var.get(),
            'enhance_image': self.enhance_var.get() == 1,
            'denoise': self.denoise_var.get() == 1,
            'adaptive_threshold': self.adaptive_var.get() == 1,
            'subtitle_mode': self.subtitle_mode_var.get() == 1,
            'text_stability': self.stability_var.get(),
            'min_confidence': float(self.conf_var.get())
        }
        
        # Update progress display
        self.progress_var.set("Processing...")
        
        # Process video in a new thread
        self.processing_thread = threading.Thread(
            target=self.process_thread, 
            args=(video_path, output_dir, ocr_options)
        )
        self.processing_thread.daemon = True
        self.processing_thread.start()
    
    def process_thread(self, video_path, output_dir, ocr_options):
        """Process video in a separate thread"""
        def update_progress(percent):
            # Use after method to update UI in main thread
            self.root.after(0, lambda: self.progress_var.set(f"Processing: {percent}%"))
        
        success, message = extract_subtitles(video_path, output_dir, ocr_options, update_progress)
        
        # Use after method to show result in main thread
        if success:
            self.root.after(0, lambda: messagebox.showinfo("Success", message))
            self.root.after(0, lambda: self.progress_var.set("Processing complete"))
            
            # Ask if user wants to open the result
            if messagebox.askyesno("Open Result", "Would you like to open the subtitle file?"):
                base_name = os.path.splitext(os.path.basename(video_path))[0]
                output_path = os.path.join(output_dir, f"{base_name}.srt")
                if os.path.exists(output_path):
                    if sys.platform == 'win32':
                        os.startfile(output_path)
                    else:
                        import subprocess
                        opener = "open" if sys.platform == "darwin" else "xdg-open"
                        subprocess.call([opener, output_path])
        else:
            self.root.after(0, lambda: messagebox.showerror("Error", message))
            self.root.after(0, lambda: self.progress_var.set("Processing failed"))
    
    def on_closing(self):
        """Handle window close event"""
        if hasattr(self, 'processing_thread') and self.processing_thread.is_alive():
            if messagebox.askokcancel("Exit", "Processing is not complete, are you sure you want to exit?"):
                self.root.destroy()
        else:
            self.root.destroy()

def main():
    root = Tk()
    app = SubtitleExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()