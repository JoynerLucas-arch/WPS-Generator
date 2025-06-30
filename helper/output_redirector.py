from queue import Queue


class OutputRedirector:
    def __init__(self, queue):
        self.queue = queue
        self.buffer = ""  # 用于处理跨多次write调用的消息
        
    def write(self, text):
        # 将新文本添加到缓冲区
        self.buffer += text
        
        # 检查是否有完整行
        if '\n' in self.buffer:
            lines = self.buffer.split('\n')
            # 最后一部分可能是不完整的行，保留在缓冲区
            self.buffer = lines[-1]
            # 处理完整的行
            for line in lines[:-1]:
                self._process_line(line + '\n')  # 添加换行符，因为split移除了它们
        elif text.endswith('\n'):
            # 如果文本以换行符结束，处理整个缓冲区
            self._process_line(self.buffer)
            self.buffer = ""
            
    def _process_line(self, line):
        # 过滤掉调试和警告信息
        if any(line.strip().startswith(prefix) for prefix in [
            "调试 -", 
            "调试信息", 
            "警告：", 
            "Warning:", 
            "Error:", 
            "错误：", 
            "未检测到图片数据"
        ]):
            return
            
        # 检查特殊情况：模板校验成功信息通常很长，包含很多标签列表
        if "模板校验成功" in line and "个标签实例" in line:
            # 只保留第一行，不显示标签列表
            summary_line = line.split('\n')[0] if '\n' in line else line
            self.queue.put(summary_line + "\n")
            return
            
        # 对于正常消息，添加到队列
        self.queue.put(line)

    def flush(self):
        # 如果缓冲区中还有内容，处理它
        if self.buffer:
            self._process_line(self.buffer)
            self.buffer = ""
        pass
