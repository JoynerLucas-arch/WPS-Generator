from abc import ABCMeta, abstractmethod
from typing import Union
import json
import re


class DataLoader(metaclass=ABCMeta):
    @abstractmethod
    def load_data(self) -> [dict, None]:
        """加载单组数据，返回内容标签名称和内容，多次加载分别返回每组数据，若无数据，返回 None"""
        pass


class StaticDataLoader(DataLoader):
    """静态数据加载器，存储所有数据并依次加载"""
    def __init__(self, datas: list[dict] = None):
        if datas is None:
            datas = []
        # 确保datas至少包含一个字典
        if not datas:
            datas = [{}]
        self._datas = datas
        self._index = 0

    def load_data(self) -> [dict, None]:
        if self._index >= len(self._datas):
            return None
        data = self._datas[self._index]
        self._index += 1
        return data


class LLMDataLoader(DataLoader):
    """处理大模型输出的数据加载器"""

    def __init__(self, llm_output: str):
        # 提取JSON部分内容
        json_match = re.search(r'```json\s*(.*?)\s*```', llm_output, re.DOTALL)
        if json_match:
            # 找到了JSON代码块
            json_str = json_match.group(1)
        else:
            # 没有找到JSON代码块，使用整个输出
            json_str = llm_output
        
        # 将Python风格的单引号转换为双引号的JSON
        json_str = json_str.replace("'", '"')
        
        try:
            # 首先尝试标准JSON解析
            data = json.loads(json_str)
            # 转换数据格式以匹配原有的格式
            self.data = {}
            for tag, value in data.items():
                if isinstance(value, list) and len(value) == 2:
                    self.data[tag] = tuple(value)
                else:
                    self.data[tag] = value
        except json.JSONDecodeError:
            # 如果标准JSON解析失败，尝试使用eval（更安全的做法）
            try:
                self.data = eval(json_str)
            except Exception as e:
                # 如果都失败，打印错误并使用空字典
                print(f"无法解析LLM输出内容: {str(e)}")
                self.data = {}
        
        # 添加默认图片数据（仅当大模型输出中没有相应图片数据时）
        self._add_default_image_data()
        
        self.loaded = False
        
    def _add_default_image_data(self):
        """添加图片数据，只在没有对应数据时添加"""
        image_data = {
            '焊接': ('', 'imgs/焊接.png'),
            '组装': ('', 'imgs/组装.png'),
            '安装': ('', 'imgs/安装.png')
        }
        # # 检查是否已有图片数据
        # has_image_data = False
        # for key, value in self.data.items():
        #     if isinstance(value, tuple) and len(value) == 2 and isinstance(value[1], str):
        #         if value[1].endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
        #             has_image_data = True
        #             break
        
        # # 如果没有图片数据，则添加默认图片数据
        # if not has_image_data:
        #     # print("未检测到图片数据，添加默认图片数据")
        self.data.update(image_data)

    def load_data(self) -> Union[dict, None]:
        """单次加载所有数据"""
        if self.loaded:
            return None
        self.loaded = True
        return self.data
