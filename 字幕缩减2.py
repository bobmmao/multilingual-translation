import re
import os
from dataclasses import dataclass
from typing import List, Dict, Optional


@dataclass
class Subtitle:
    index: int
    start_time: str
    end_time: str
    content: str
    
    def __str__(self) -> str:
        return f"{self.index}\n{self.start_time} --> {self.end_time}\n{self.content}\n"


def normalize_content(content: str) -> str:
    """Normalize subtitle content for comparison purposes."""
    # Remove standalone numbers at the beginning of lines
    content = re.sub(r'(?m)^\s*\d+\s*', '', content)
    
    # Normalize all whitespace (spaces, tabs, newlines) to a single space for comparison
    normalized_content = ' '.join(content.split())
    
    # Remove formatting tags that might differ between otherwise identical content
    normalized_content = re.sub(r'<[^>]+>', '', normalized_content)
    
    return normalized_content


def extract_chinese_text(content: str) -> str:
    """Extract only Chinese characters from the content."""
    # This regular expression matches Chinese characters (including traditional and simplified)
    chinese_chars = re.findall(r'[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]+', content)
    return ''.join(chinese_chars)


def is_time_consecutive(end_time1: str, start_time2: str, max_gap_ms: int = 1000) -> bool:
    """Check if two timestamps are consecutive with a small acceptable gap."""
    def time_to_ms(time_str: str) -> int:
        """Convert SRT time format to milliseconds."""
        h, m, s_ms = time_str.split(':')
        s, ms = s_ms.split(',')
        return int(h) * 3600000 + int(m) * 60000 + int(s) * 1000 + int(ms)
    
    end_ms = time_to_ms(end_time1)
    start_ms = time_to_ms(start_time2)
    
    # Check if times are consecutive (allowing for a gap)
    return start_ms - end_ms <= max_gap_ms


def clean_subtitle_content(content: str) -> str:
    """Clean up subtitle content by removing numbers at the beginning of the last line."""
    # Split content into lines
    lines = content.split('\n')
    
    # Process the last non-empty line
    for i in range(len(lines)-1, -1, -1):
        if lines[i].strip():  # Find the last non-empty line
            # Remove numbers that appear at the beginning of the last line
            lines[i] = re.sub(r'^\s*\d+\s*', '', lines[i])
            # If the line became empty, remove it
            if not lines[i].strip():
                lines.pop(i)
            break
    
    # Rejoin the lines and clean up any empty lines
    content = '\n'.join(lines)
    content = re.sub(r'\n+', '\n', content).strip()
    
    return content


def merge_similar_subtitles(subtitles: List[Subtitle]) -> List[Subtitle]:
    """Merge consecutive subtitles with similar content using more aggressive criteria."""
    if not subtitles:
        return []
    
    merged_subtitles = []
    current_group = [subtitles[0]]
    
    for i in range(1, len(subtitles)):
        current = subtitles[i]
        previous = current_group[-1]
        
        # Check if content is similar (after normalization)
        current_normalized = normalize_content(current.content)
        previous_normalized = normalize_content(previous.content)
        
        # Extract Chinese text for both contents
        current_chinese = extract_chinese_text(current.content)
        previous_chinese = extract_chinese_text(previous.content)
        
        # Check if Chinese content is similar when there's enough Chinese text
        chinese_similar = False
        if current_chinese and previous_chinese:
            # If both have Chinese text, compare them
            if len(current_chinese) > 0 and len(previous_chinese) > 0:
                common_length = min(len(current_chinese), len(previous_chinese))
                if common_length > 0:
                    chinese_similarity = sum(1 for c1, c2 in zip(current_chinese, previous_chinese) if c1 == c2) / common_length
                    chinese_similar = chinese_similarity > 0.8
        
        # Check if timestamps are consecutive
        times_consecutive = is_time_consecutive(previous.end_time, current.start_time, max_gap_ms=1000)
        
        # Special case: handle space differences in otherwise identical content
        space_normalized_current = re.sub(r'\s+', '', current_normalized)
        space_normalized_previous = re.sub(r'\s+', '', previous_normalized)
        space_difference_only = space_normalized_current == space_normalized_previous
        
        # Content similarity check with various criteria
        content_similar = (
            current_normalized == previous_normalized or 
            current_normalized in previous_normalized or 
            previous_normalized in current_normalized or
            space_difference_only or
            # Calculate similarity ratio with a reasonable threshold
            (len(current_normalized) > 5 and len(previous_normalized) > 5 and
             sum(1 for c1, c2 in zip(current_normalized, previous_normalized) if c1 == c2) / 
             max(len(current_normalized), len(previous_normalized)) > 0.8) or
            # Check Chinese similarity  
            chinese_similar
        )
        
        if content_similar and times_consecutive:
            # Extend the time range of the current group
            current_group.append(current)
        else:
            # Create a merged subtitle from current group
            if len(current_group) == 1:
                # Clean the content even for non-merged subtitles
                current_group[0].content = clean_subtitle_content(current_group[0].content)
                merged_subtitles.append(current_group[0])
            else:
                first = current_group[0]
                last = current_group[-1]
                
                # Choose the longer content version
                if len(first.content) >= len(last.content):
                    content = first.content
                else:
                    content = last.content
                
                # Clean up the content
                content = clean_subtitle_content(content)
                
                merged = Subtitle(
                    index=first.index,
                    start_time=first.start_time,
                    end_time=last.end_time,
                    content=content
                )
                merged_subtitles.append(merged)
            
            # Start a new group with the current subtitle
            current_group = [current]
    
    # Handle the last group
    if current_group:
        if len(current_group) == 1:
            # Clean the content even for non-merged subtitles
            current_group[0].content = clean_subtitle_content(current_group[0].content)
            merged_subtitles.append(current_group[0])
        else:
            first = current_group[0]
            last = current_group[-1]
            
            # Choose the longer content version
            if len(first.content) >= len(last.content):
                content = first.content
            else:
                content = last.content
                
            # Clean up the content
            content = clean_subtitle_content(content)
            
            merged = Subtitle(
                index=first.index,
                start_time=first.start_time,
                end_time=last.end_time,
                content=content
            )
            merged_subtitles.append(merged)
    
    # Reindex subtitles
    for i, subtitle in enumerate(merged_subtitles, 1):
        subtitle.index = i
    
    return merged_subtitles


def save_srt_file(subtitles: List[Subtitle], output_path: str) -> None:
    """Save a list of Subtitle objects to an SRT file."""
    with open(output_path, 'w', encoding='utf-8') as f:
        for subtitle in subtitles:
            f.write(str(subtitle) + '\n')


def clean_srt_file(input_path: str) -> List[Subtitle]:
    """Read an SRT file, fix common formatting issues, and return cleaned subtitles."""
    with open(input_path, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()
    
    # 直接提取所有包含时间戳的行及其后面的内容
    time_pattern = re.compile(r'(\d{2}:\d{2}:\d{2},\d{3}) --> (\d{2}:\d{2}:\d{2},\d{3})')
    
    # 替换掉Windows和Mac的不同换行符，统一为\n
    content = content.replace('\r\n', '\n').replace('\r', '\n')
    
    # 将内容分割成行
    lines = content.split('\n')
    
    subtitles = []
    current_index = 0
    i = 0
    
    while i < len(lines):
        line = lines[i].strip()
        # 查找时间戳行
        match = time_pattern.match(line)
        
        if match:
            start_time, end_time = match.groups()
            
            # 收集时间戳行后的所有内容，直到下一个时间戳或文件结束
            content_lines = []
            j = i + 1
            while j < len(lines) and not time_pattern.match(lines[j].strip()):
                if lines[j].strip():  # 只添加非空行
                    content_lines.append(lines[j].strip())
                j += 1
            
            # 创建字幕对象
            current_index += 1
            subtitle_content = '\n'.join(content_lines)
            
            if subtitle_content:  # 只添加有内容的字幕
                subtitles.append(Subtitle(current_index, start_time, end_time, subtitle_content))
            
            i = j  # 跳到下一个时间戳的位置
        else:
            i += 1  # 继续检查下一行
    
    return subtitles


def emergency_parse_file(input_path: str) -> List[Subtitle]:
    """在正常解析失败时尝试最基本的解析方法"""
    try:
        with open(input_path, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        
        # 查找所有时间戳
        time_pattern = re.compile(r'(\d{2}:\d{2}:\d{2},\d{3}) --> (\d{2}:\d{2}:\d{2},\d{3})')
        matches = time_pattern.finditer(content)
        
        subtitles = []
        for i, match in enumerate(matches, 1):
            start_time, end_time = match.groups()
            
            # 找到本次匹配的结束位置
            current_pos = match.end()
            
            # 查找下一个时间戳的开始位置
            next_match = time_pattern.search(content, current_pos)
            if next_match:
                next_pos = next_match.start()
                subtitle_text = content[current_pos:next_pos].strip()
            else:
                subtitle_text = content[current_pos:].strip()
            
            # 去除可能的索引号
            subtitle_text = re.sub(r'^\s*\d+\s*$', '', subtitle_text, flags=re.MULTILINE)
            subtitle_text = subtitle_text.strip()
            
            if subtitle_text:
                subtitles.append(Subtitle(i, start_time, end_time, subtitle_text))
        
        return subtitles
    except Exception as e:
        print(f"Emergency parsing failed: {e}")
        return []


def process_srt_file(input_path: str, output_path: Optional[str] = None) -> None:
    """Process an SRT file by merging similar subtitles with aggressive merging."""
    if output_path is None:
        # Generate output filename based on input filename
        base_name, ext = os.path.splitext(input_path)
        output_path = f"{base_name}_merged{ext}"
    
    try:
        # Parse and clean subtitles
        subtitles = clean_srt_file(input_path)
        print(f"Read {len(subtitles)} subtitles from {input_path}")
        
        if not subtitles:
            print("No valid subtitles found in the file. Attempting emergency parsing...")
            
            # 紧急处理：直接从文件内容中提取时间戳和文本
            subtitles = emergency_parse_file(input_path)
            print(f"Emergency parsing found {len(subtitles)} subtitles")
            
            if not subtitles:
                print("Still no valid subtitles found. Check the file format.")
                return
        
        # Merge similar subtitles with aggressive criteria
        merged_subtitles = merge_similar_subtitles(subtitles)
        print(f"Merged into {len(merged_subtitles)} subtitles")
        
        # Calculate reduction percentage
        reduction = (1 - len(merged_subtitles) / len(subtitles)) * 100
        print(f"Reduced by {reduction:.1f}%")
        
        # Save merged subtitles
        save_srt_file(merged_subtitles, output_path)
        print(f"Saved merged subtitles to {output_path}")
    except Exception as e:
        print(f"Error processing file {input_path}: {e}")
        import traceback
        traceback.print_exc()


def process_directory(directory_path: str) -> None:
    """Process all SRT files in a directory."""
    for file_name in os.listdir(directory_path):
        if file_name.lower().endswith('.srt'):
            input_path = os.path.join(directory_path, file_name)
            name, ext = os.path.splitext(file_name)
            output_path = os.path.join(directory_path, f"{name}_merged{ext}")
            
            print(f"\nProcessing: {file_name}")
            process_srt_file(input_path, output_path)


if __name__ == "__main__":
    # Option 1: Process a specific file
    input_file = r"C:\\Users\\admin\\Desktop\\字幕提取\\001_Installation Video_With Sub.srt"
    
    # Generate output file path by adding "_merged" suffix to the original filename
    base_dir = os.path.dirname(input_file)
    base_name = os.path.basename(input_file)
    name, ext = os.path.splitext(base_name)
    output_file = os.path.join(base_dir, f"{name}_merged{ext}")
    
    # Process the file
    print(f"Processing single file: {base_name}")
    process_srt_file(input_file, output_file)
    
    # Option 2: Uncomment this to process all SRT files in the directory
    # directory_path = os.path.dirname(input_file)  # Use the same directory as the input file
    # print(f"\nProcessing all SRT files in: {directory_path}")
    # process_directory(directory_path)