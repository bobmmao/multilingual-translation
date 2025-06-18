import re
import os

def process_srt(input_file, output_file):
    """
    处理SRT文件，移除时间戳行和多余空行，清理格式
    
    参数:
        input_file (str): 输入SRT文件的路径
        output_file (str): 保存输出的文件路径
    """
    # 读取输入文件
    with open(input_file, 'r', encoding='utf-8') as file:
        content = file.read()
    
    # 步骤1: 移除所有数字索引 (例如 "1", "2", "3")
    content = re.sub(r'^\d+\r?\n', '', content, flags=re.MULTILINE)
    
    # 步骤2: 移除所有时间戳行 (例如 "00:00:01,600 --> 00:00:06,140")
    content = re.sub(r'^\d{2}:\d{2}:\d{2},\d{3}\s-->\s\d{2}:\d{2}:\d{2},\d{3}\r?\n', '', content, flags=re.MULTILINE)
    
    # 步骤3: 移除字体颜色标签
    content = re.sub(r'<font color=#[0-9A-Fa-f]+FF>|</font>', '', content)
    
    # 步骤4: 移除空行 (连续的换行)
    content = re.sub(r'\r?\n\r?\n+', '\n', content)
    
    # 步骤5: 移除行尾的空白字符
    content = re.sub(r'\s+$', '', content, flags=re.MULTILINE)
    
    # 步骤6: 移除行首的空白字符
    content = re.sub(r'^\s+', '', content, flags=re.MULTILINE)
    
    # 将处理后的内容写入输出文件
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write(content)
    
    print(f"处理完成。输出已保存到 {output_file}")


if __name__ == "__main__":
    # 定义输入和输出文件路径
    input_file = "C:\\Users\\admin\\Desktop\\字幕提取\\U200 Lite安装视频.srt"
    output_file = "C:\\Users\\admin\\Desktop\\字幕提取\\U200_Lite_Processed.txt"
    
    # Process the file
    process_srt(input_file, output_file)