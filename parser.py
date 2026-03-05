"""
解析Whisper转录文件，提取时间戳和文案
支持格式: [开始时间-结束时间]\n文案内容
"""
import re


def parse_transcript_file(file_path):
    """
    解析转录文件，返回结构化数据列表
    
    Args:
        file_path: 转录文件路径
        
    Returns:
        list: 每个元素包含 {
            'index': 索引(从0开始),
            'start_time': 开始时间(秒),
            'end_time': 结束时间(秒),
            'duration': 时长(秒),
            'text': 文案内容,
            'gap_after': 与下一句的间隔(秒)
        }
    """
    segments = []
    
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    i = 0
    index = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # 匹配时间戳格式 [0.00-2.68]
        match = re.match(r'\[(\d+\.?\d*)-(\d+\.?\d*)\]', line)
        if match:
            start_time = float(match.group(1))
            end_time = float(match.group(2))
            
            # 获取文案内容（下一行）
            text = ""
            if i + 1 < len(lines):
                text = lines[i + 1].strip()
            
            segments.append({
                'index': index,
                'start_time': start_time,
                'end_time': end_time,
                'duration': end_time - start_time,
                'text': text,
                'gap_after': 0  # 将在后面计算
            })
            
            index += 1
            i += 2  # 跳过时间戳和文案两行
        else:
            i += 1
    
    # 计算句间间隔
    for i in range(len(segments) - 1):
        segments[i]['gap_after'] = segments[i + 1]['start_time'] - segments[i]['end_time']
    
    # 最后一句没有间隔
    if segments:
        segments[-1]['gap_after'] = 0
    
    return segments


def format_time(seconds):
    """将秒数格式化为 MM:SS.ms 格式"""
    minutes = int(seconds // 60)
    secs = seconds % 60
    return f"{minutes:02d}:{secs:05.2f}"


if __name__ == "__main__":
    # 测试解析
    segments = parse_transcript_file("RA2_English_Colloquial.txt")
    print(f"共解析 {len(segments)} 个句子")
    print("\n前3个句子:")
    for seg in segments[:3]:
        print(f"[{seg['index']}] {format_time(seg['start_time'])}-{format_time(seg['end_time'])}")
        print(f"    {seg['text']}")
        print(f"    时长: {seg['duration']:.2f}秒, 间隔: {seg['gap_after']:.2f}秒\n")
