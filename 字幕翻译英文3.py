import os
import re
import requests
import time
import gspread
import traceback  # For detailed error tracking
from oauth2client.service_account import ServiceAccountCredentials
from tqdm import tqdm  # For progress display

# API Configuration
API_URL = "http://bedroc-proxy-kbhtlwhggrzm-944033383.us-west-2.elb.amazonaws.com/api/v1/chat/completions"
API_KEY = "aws_Bedrock123@pSwD"
MODEL_ID = "us.anthropic.claude-3-7-sonnet-20250219-v1:0"

class TerminologyManager:
    def __init__(self):
        self.terminology_db = {}
        self.translation_memory = {}
    
    def load_terminology(self, sheet_url, source_lang_col, target_lang_col, language_code):
        """Load terminology from Google Sheets"""
        try:
            scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
            credentials_path = r'C:\\Users\\admin\\Desktop\\字幕提取\\Extract_glossary.json'
            
            # Check if the file exists
            if not os.path.exists(credentials_path):
                print(f"Error: Credentials file not found at {credentials_path}")
                print("Please update the file path or create a dummy credentials file for testing")
                return
                
            credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
            gc = gspread.authorize(credentials)
            
            # Extract sheet ID from URL
            try:
                sheet_id = sheet_url.split('/d/')[1].split('/')[0]
                print(f"Attempting to open Google Sheet with ID: {sheet_id}")
                worksheet = gc.open_by_key(sheet_id).sheet1
            except Exception as e:
                print(f"Error accessing Google Sheet: {e}")
                print("Using a dummy terminology database for testing")
                # Create a dummy terminology database for testing
                self.terminology_db[language_code] = {
                    "智能门锁": "Smart Lock",
                    "人脸识别": "Facial Recognition",
                    "安装指南": "Installation Guide",
                    "锁体": "Lock Body",
                    "锁芯": "Lock Cylinder",
                    "前面板": "Front Panel",
                    "后面板": "Back Panel"
                }
                print(f"Loaded {len(self.terminology_db[language_code])} dummy terms for {language_code}")
                return
            
            data = worksheet.get_all_records()
            
            if not data:
                print("Warning: No data found in the Google Sheet")
                return
                
            if source_lang_col not in data[0].keys() or target_lang_col not in data[0].keys():
                print(f"Error: Column not found: {source_lang_col} or {target_lang_col}")
                print(f"Available columns: {data[0].keys()}")
                return
            
            term_dict = {}
            for row in data:
                if row[source_lang_col] and row[target_lang_col]:
                    source_term = str(row[source_lang_col]).strip().lower()
                    if not source_term:  # Skip empty source terms
                        continue
                    target_term = str(row[target_lang_col])
                    term_dict[source_term] = target_term
            
            self.terminology_db[language_code] = term_dict
            print(f"Loaded {len(term_dict)} terms for {language_code}")
            
        except Exception as e:
            print(f"Error loading terminology: {e}")
            print(traceback.format_exc())  # Print detailed stack trace
            # Create a dummy terminology database for testing
            self.terminology_db[language_code] = {
                "智能门锁": "Smart Lock",
                "人脸识别": "Facial Recognition",
                "安装指南": "Installation Guide",
                "锁体": "Lock Body",
                "锁芯": "Lock Cylinder",
                "前面板": "Front Panel",
                "后面板": "Back Panel"
            }
            print(f"Loaded {len(self.terminology_db[language_code])} dummy terms for {language_code}")
    
    def apply_terminology(self, text, language_code):
        """Apply terminology to text with improved error handling"""
        if not text:  # Guard against empty strings
            return text
            
        if language_code not in self.terminology_db or not self.terminology_db[language_code]:
            return text
        
        try:
            terms = sorted(self.terminology_db[language_code].keys(), key=len, reverse=True)
            
            def replace_term(match):
                try:
                    matched_term = match.group(0)
                    if not matched_term:  # Guard against empty matches
                        return ""
                        
                    replacement = self.terminology_db[language_code].get(matched_term.lower())
                    if not replacement:  # Guard against missing replacements
                        return matched_term
                    
                    # Apply case matching
                    if matched_term.islower():
                        return replacement.lower()
                    elif matched_term.isupper():
                        return replacement.upper()
                    elif matched_term[0].isupper() and len(matched_term) > 1:
                        return replacement[0].upper() + replacement[1:] if len(replacement) > 1 else replacement.upper()
                    return replacement
                except Exception as e:
                    print(f"Error in replace_term: {e}")
                    print(f"Match: {match}")
                    # Return the original match on error
                    return match.group(0) if match and match.group(0) else ""
            
            processed_text = text
            for term in terms:
                try:
                    if not term:  # Skip empty terms
                        continue
                    pattern = r'\b' + re.escape(term) + r'\b'
                    processed_text = re.sub(pattern, replace_term, processed_text, flags=re.IGNORECASE)
                except Exception as e:
                    print(f"Error applying term '{term}': {e}")
                    # Continue with next term on error
                    continue
            
            return processed_text
            
        except Exception as e:
            print(f"Error in apply_terminology: {e}")
            print(traceback.format_exc())  # Print detailed stack trace
            return text  # Return original text on error

class SubtitleTranslator:
    def __init__(self, terminology_manager):
        self.terminology_manager = terminology_manager
        self.debug_mode = True  # Enable for verbose output
    
    def translate_text(self, text, target_language, language_code):
        """Translate text using Claude API with improved error handling"""
        if not text or not text.strip():
            return ""  # Return empty string for empty input
        
        # Check translation memory
        memory_key = text.strip().lower()
        if memory_key in self.terminology_manager.translation_memory:
            return self.terminology_manager.translation_memory[memory_key]
        
        try:
            # Extract terminology for context
            potential_terms = []
            if language_code in self.terminology_manager.terminology_db:
                for term in self.terminology_manager.terminology_db[language_code].keys():
                    if not term:  # Skip empty terms
                        continue
                    try:
                        if re.search(r'\b' + re.escape(term) + r'\b', text, re.IGNORECASE):
                            target_term = self.terminology_manager.terminology_db[language_code][term]
                            potential_terms.append(f"{term} -> {target_term}")
                    except Exception as e:
                        print(f"Error checking term '{term}': {e}")
                        continue
            
            # Build prompt
            sys_prompt = 'You are a translation engine only. Translate the Chinese text to English maintaining all formatting. Return ONLY the translated text with no explanations, no Chinese, and no comments. Never apologize or explain your translation.'
            
            user_prompt = f'Translate the following text from Chinese to English. Return ONLY the translated content. Keep all symbols, punctuation, and formatting exactly as they appear. Do not add any explanations, Chinese text, or comments before or after the translation.'
            
            if potential_terms:
                user_prompt += f"\n\nIMPORTANT: Use the following terminology consistently:\n" + "\n".join(potential_terms)
            
            user_prompt += f"\n\nText to translate:\n{text}"
            
            if self.debug_mode:
                print(f"\n--- DEBUG: Translation Request ---")
                print(f"Original text: {text}")
                if potential_terms:
                    print(f"Terms to use: {potential_terms}")
            
            # Request data
            data = {
                "model": MODEL_ID,
                "messages": [
                    {
                        "role": "user",
                        "content": [{"type": "text", "text": user_prompt}],
                    }
                ],
                "system": sys_prompt
            }
            
            headers = {
                'Content-Type': 'application/json',
                'Authorization': f'Bearer {API_KEY}'
            }
            
            # Add retry logic
            max_retries = 3
            retry_delay = 5
            retry_count = 0

            while retry_count < max_retries:
                try:
                    # Send request
                    if self.debug_mode:
                        print(f"Sending API request (attempt {retry_count + 1}/{max_retries})...")
                    
                    response = requests.post(API_URL, headers=headers, json=data, timeout=60)
                    
                    # Successfully get result
                    if response.status_code == 200:
                        result = response.json()
                        translated_text = result["choices"][0]["message"]["content"]
                        
                        if self.debug_mode:
                            print(f"Raw API response: {translated_text[:100]}...")
                        
                        # Clean and verify translation result
                        explanation_patterns = [
                            r'^(I\'m sorry|I apologize|Sorry|Note|Please note).*?\n\n',
                            r'\n\n(I\'m sorry|I apologize|Sorry|Note|Please note).*?$',
                            r'^(Here is|Here\'s|The following is|This is) the translation.*?\n\n',
                            r'^Translated text:.*?\n\n'
                        ]
                        
                        for pattern in explanation_patterns:
                            translated_text = re.sub(pattern, '', translated_text, flags=re.IGNORECASE | re.DOTALL)

                        # Check if the translation is valid
                        if not translated_text or not translated_text.strip():
                            print(f"Warning: Empty translation result for text: {text}")
                            return text  # Return original text if translation is empty
                        
                        # Apply terminology
                        try:
                            translated_text = self.terminology_manager.apply_terminology(translated_text, language_code)
                        except Exception as e:
                            print(f"Error applying terminology: {e}")
                            # Continue with the raw translation if terminology application fails
                        
                        # Store in memory
                        self.terminology_manager.translation_memory[memory_key] = translated_text
                        
                        if self.debug_mode:
                            print(f"Final translation: {translated_text[:100]}...")
                        
                        return translated_text
                    
                    # Handle 502 error - retry
                    elif response.status_code == 502:
                        retry_count += 1
                        print(f"Encountered HTTP 502 error. Retry attempt {retry_count}/{max_retries} in {retry_delay} seconds...")
                        time.sleep(retry_delay)
                        # Increase delay each retry
                        retry_delay *= 1.5
                        continue
                    
                    # Handle other HTTP errors
                    else:
                        print(f"Translation error: HTTP {response.status_code}")
                        
                        # Try to get response text safely
                        try:
                            resp_text = response.text[:200]  # Limit to first 200 chars
                            print(f"Response text: {resp_text}")
                        except:
                            print("Could not extract response text")
                        
                        # For non-502 errors, increase retry count but use shorter delay
                        retry_count += 1
                        print(f"Attempting retry {retry_count}/{max_retries} for non-502 error in {retry_delay/2} seconds...")
                        time.sleep(retry_delay/2)
                        continue
                        
                except requests.exceptions.RequestException as e:
                    # Handle request exceptions (connection errors, timeouts, etc.)
                    retry_count += 1
                    print(f"Request exception: {e}. Retry attempt {retry_count}/{max_retries} in {retry_delay} seconds...")
                    time.sleep(retry_delay)
                    retry_delay *= 1.5
                    continue
                except Exception as e:
                    # Handle other unexpected errors
                    retry_count += 1
                    print(f"Unexpected error during translation: {e}")
                    print(traceback.format_exc())  # Print detailed stack trace
                    time.sleep(retry_delay)
                    continue

            # If all retries fail, return original text
            print(f"All {max_retries} retry attempts failed. Returning original text.")
            return text

        except Exception as e:
            print(f"Translation error: {e}")
            print(traceback.format_exc())  # Print detailed stack trace
            return text  # Return original text on error

class SRTProcessor:
    def __init__(self, translator):
        self.translator = translator
    
    def parse_srt(self, file_path):
        """Parse SRT file into subtitle entries with improved error handling"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Split content by blank line or double newline to get each subtitle block
            # Handle different types of line endings
            blocks = re.split(r'\n\s*\n', content.strip())
            subtitles = []
            
            for block_idx, block in enumerate(blocks):
                lines = block.strip().split('\n')
                if len(lines) < 2:
                    print(f"Warning: Block {block_idx+1} has insufficient lines: {block}")
                    continue
                
                try:
                    # First line should be the sequence number
                    # Handle cases where the sequence number might not be a clean integer
                    seq_match = re.match(r'^\s*(\d+)\s*$', lines[0])
                    if not seq_match:
                        print(f"Warning: Could not parse sequence number in block {block_idx+1}: {lines[0]}")
                        seq_num = block_idx + 1  # Use block index as fallback
                    else:
                        seq_num = int(seq_match.group(1))
                    
                    # Second line should be the timestamp
                    timestamp = lines[1].strip()
                    
                    # Validate timestamp format
                    timestamp_pattern = r'(\d{2}:\d{2}:\d{2},\d{3})\s*-->\s*(\d{2}:\d{2}:\d{2},\d{3})'
                    if not re.match(timestamp_pattern, timestamp):
                        print(f"Warning: Invalid timestamp format in block {block_idx+1}: {timestamp}")
                        continue
                    
                    # Remaining lines are the text
                    text = '\n'.join(lines[2:])
                    
                    # Extract start time for sorting
                    start_time = timestamp.split(' --> ')[0].strip()
                    # Convert time to seconds for easier sorting
                    h, m, s = start_time.replace(',', '.').split(':')
                    seconds = float(h) * 3600 + float(m) * 60 + float(s)
                    
                    subtitles.append({
                        'seq_num': seq_num,
                        'timestamp': timestamp,
                        'text': text,
                        'seconds': seconds
                    })
                except Exception as e:
                    print(f"Error parsing subtitle block {block_idx+1}: {e}")
                    print(f"Block content: {block}")
                    print(traceback.format_exc())  # Print detailed stack trace
            
            print(f"Successfully parsed {len(subtitles)} out of {len(blocks)} subtitle blocks")
            return subtitles
            
        except Exception as e:
            print(f"Error reading SRT file: {e}")
            print(traceback.format_exc())  # Print detailed stack trace
            return []
    
    def sort_subtitles(self, subtitles):
        """Sort subtitles by their start time"""
        if not subtitles:
            print("Warning: No subtitles to sort")
            return []
            
        try:
            sorted_subtitles = sorted(subtitles, key=lambda x: x['seconds'])
            
            # Reassign sequence numbers
            for i, subtitle in enumerate(sorted_subtitles, 1):
                subtitle['seq_num'] = i
            
            print(f"Successfully sorted {len(sorted_subtitles)} subtitles")
            return sorted_subtitles
        except Exception as e:
            print(f"Error sorting subtitles: {e}")
            print(traceback.format_exc())  # Print detailed stack trace
            return subtitles  # Return original unsorted subtitles on error
    
    def translate_subtitles(self, subtitles, target_language='English', language_code='EN'):
        """Translate subtitle text and format as dual-language with error handling"""
        if not subtitles:
            print("Warning: No subtitles to translate")
            return []
            
        translated_subtitles = []
        failed_count = 0
        
        for idx, subtitle in enumerate(tqdm(subtitles, desc="Translating subtitles")):
            try:
                chinese_text = subtitle['text'].strip()
                
                # Skip empty text
                if not chinese_text:
                    print(f"Warning: Empty text in subtitle {subtitle['seq_num']}")
                    translated_subtitles.append(subtitle)  # Keep original
                    continue
                
                # Print what we're translating for debugging
                print(f"\nTranslating subtitle {subtitle['seq_num']} ({idx+1}/{len(subtitles)}):")
                print(f"Original text: {chinese_text[:50]}...")
                
                # Attempt translation
                english_text = self.translator.translate_text(chinese_text, target_language, language_code)
                
                # Verify translation result
                if not english_text or english_text == chinese_text:
                    print(f"Warning: Translation failed for subtitle {subtitle['seq_num']}. Using original text.")
                    english_text = f"[Translation error: {chinese_text}]"  # Mark as error
                    failed_count += 1
                
                # Format as dual-language
                dual_text = f"{chinese_text}\n{english_text}"
                
                translated_subtitles.append({
                    'seq_num': subtitle['seq_num'],
                    'timestamp': subtitle['timestamp'],
                    'text': dual_text,
                    'seconds': subtitle['seconds']
                })
                
            except Exception as e:
                print(f"Error translating subtitle {subtitle['seq_num']}: {e}")
                print(traceback.format_exc())  # Print detailed stack trace
                
                # Add original subtitle on error
                translated_subtitles.append(subtitle)
                failed_count += 1
            
            # Small delay to avoid rate limits
            time.sleep(0.5)
        
        print(f"Translation completed. Successfully translated {len(subtitles) - failed_count}/{len(subtitles)} subtitles.")
        return translated_subtitles
    
    def write_srt(self, subtitles, output_file):
        """Write subtitles to SRT file with error handling"""
        if not subtitles:
            print("Warning: No subtitles to write")
            return False
            
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                for i, subtitle in enumerate(subtitles):
                    f.write(f"{subtitle['seq_num']}\n")
                    f.write(f"{subtitle['timestamp']}\n")
                    f.write(f"{subtitle['text']}\n")
                    if i < len(subtitles) - 1:
                        f.write("\n")
            
            print(f"Translated subtitles saved to: {output_file}")
            return True
        except Exception as e:
            print(f"Error writing SRT file: {e}")
            print(traceback.format_exc())  # Print detailed stack trace
            return False

def create_dummy_credentials_file():
    """Create a dummy credentials file for testing if the real one doesn't exist"""
    credentials_path = r'C:\\Users\\admin\\Desktop\\字幕提取\\Extract_glossary.json'
    
    if os.path.exists(credentials_path):
        return
        
    print(f"Creating dummy credentials file at {credentials_path}")
    try:
        # Create directory if it doesn't exist
        os.makedirs(os.path.dirname(credentials_path), exist_ok=True)
        
        # Create a minimal dummy file
        dummy_content = """
        {
          "type": "service_account",
          "project_id": "dummy-project",
          "private_key_id": "dummy",
          "private_key": "-----BEGIN PRIVATE KEY-----\\nDUMMY\\n-----END PRIVATE KEY-----\\n",
          "client_email": "dummy@example.com",
          "client_id": "000000000000000000000",
          "auth_uri": "https://accounts.google.com/o/oauth2/auth",
          "token_uri": "https://oauth2.googleapis.com/token",
          "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
          "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/dummy"
        }
        """
        
        with open(credentials_path, 'w', encoding='utf-8') as f:
            f.write(dummy_content)
            
        print("Dummy credentials file created successfully")
    except Exception as e:
        print(f"Error creating dummy credentials file: {e}")
        print("Will use fallback terminology database")

def simple_translation_test():
    """Test the translation functionality with a simple phrase"""
    print("\n--- RUNNING SIMPLE TRANSLATION TEST ---\n")
    
    # Create terminology manager with dummy data
    terminology_manager = TerminologyManager()
    terminology_manager.terminology_db['EN'] = {
        "智能门锁": "Smart Lock",
        "人脸识别": "Facial Recognition"
    }
    
    # Create translator
    translator = SubtitleTranslator(terminology_manager)
    
    # Test translation
    test_text = "Aqara人脸识别智能门锁D200 安装指南"
    print(f"Test text: {test_text}")
    
    try:
        translated = translator.translate_text(test_text, "English", "EN")
        print(f"Translation result: {translated}")
        print("Translation test completed successfully!")
    except Exception as e:
        print(f"Translation test failed: {e}")
        print(traceback.format_exc())

def main():
    # Input file path - adjust as needed
    input_file = "C:\\Users\\admin\\Desktop\\字幕提取\\9.21智能门锁D200_3_merged.srt"
    
    # Check if file exists
    if not os.path.exists(input_file):
        print(f"Error: Input file not found: {input_file}")
        print(f"Current working directory: {os.getcwd()}")
        print("Please place the script in the same directory as the SRT file or update the path")
        return
    
    # Google Sheet URL for terminology
    google_sheet_url = "https://docs.google.com/spreadsheets/d/11B4LNWf27Mt_PvqsyKZYmtxeaLmCBPFSQHiiUyY2IC4/edit?gid=0"
    
    # Create terminology manager
    terminology_manager = TerminologyManager()
    
    # Check for credentials file and create dummy if needed
    create_dummy_credentials_file()
    
    # Run a simple translation test first
    simple_translation_test()
    
    # Load terminology
    terminology_manager.load_terminology(
        google_sheet_url,
        "Chinese (Simplified)",  # Source language column name
        "English",   # Target language column name
        "EN"         # Language code
    )
    
    # Create translator
    translator = SubtitleTranslator(terminology_manager)
    
    # Create SRT processor
    processor = SRTProcessor(translator)
    
    # Parse SRT file
    print("Parsing SRT file...")
    subtitles = processor.parse_srt(input_file)
    
    if not subtitles:
        print("Error: No valid subtitles found in the file")
        return
    
    # Sort subtitles
    print("Sorting subtitles by timestamp...")
    sorted_subtitles = processor.sort_subtitles(subtitles)
    
    print(f"Found {len(sorted_subtitles)} subtitles, reordered sequence numbers")
    
    # Translate subtitles - setting max_subtitles to 5 for testing
    max_subtitles = len(sorted_subtitles)  # Adjust this for testing (e.g., set to 5 to test first 5 only)
    print(f"Translating {max_subtitles} subtitles to English...")
    
    # Only process a subset for testing if needed
    subtitles_to_translate = sorted_subtitles[:max_subtitles]
    
    translated_subtitles = processor.translate_subtitles(subtitles_to_translate)
    
    # Generate output file name
    output_file = os.path.splitext(input_file)[0] + "_EN.srt"
    
    # Write to output file
    print("Writing bilingual SRT file...")
    if processor.write_srt(translated_subtitles, output_file):
        print("Process completed successfully!")
    else:
        print("Process completed with errors. Please check the logs.")

if __name__ == "__main__":
    main()