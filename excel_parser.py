#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import re
from typing import Dict, Any, Optional

class ExcelParser:
    """Excel文件解析器，用于解析焊接接头清单数据"""
    
    def __init__(self):
        self.data = None
        self.parsed_dict = None
        self.current_row_index = 0  # 添加当前行索引
    
    def clean_column_names(self, df: pd.DataFrame) -> pd.DataFrame:
        """清理列名中的空格和回车"""
        cleaned_columns = {}
        for col in df.columns:
            if isinstance(col, str):
                # 移除所有空白字符（空格、回车、换行符、制表符等）
                cleaned_col = re.sub(r'[\s\r\n\t]+', '', str(col).strip())
                # 进一步清理，确保没有残留的特殊字符
                cleaned_col = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', cleaned_col)
                cleaned_columns[col] = cleaned_col
            else:
                cleaned_columns[col] = col
        
        return df.rename(columns=cleaned_columns)
    
    def standardize_thickness_material(self, value) -> str:
        """标准化厚度/材质格式为 XXmm XXXXX-XX"""
        if pd.isna(value) or value is None:
            return str(value) if value is not None else ""
        
        # 转换为字符串并清理多余空格
        value_str = re.sub(r'\s+', ' ', str(value).strip())
        
        # 使用正则表达式匹配厚度和材质
        # 匹配模式：数字+mm 材质代码
        pattern = r'(\d+(?:\.\d+)?)(mm)?\s*([A-Za-z0-9\-]+)'
        match = re.search(pattern, value_str)
        
        if match:
            thickness = match.group(1)
            material = match.group(3)
            return f"{thickness}mm {material}"
        else:
            # 如果无法匹配，返回清理后的原值
            return value_str
    
    def load_excel_data(self, file_path: str, sheet_name_or_index: int = 2, header_row: int = 2) -> bool:
        """
        加载Excel文件数据
        
        Args:
            file_path: Excel文件路径
            sheet_name_or_index: 工作表名称或索引（默认第3个工作表，索引为2）
            header_row: 标题行位置（默认第3行，索引为2）
            
        Returns:
            bool: 是否成功加载
        """
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path, sheet_name=sheet_name_or_index, header=header_row)
            
            # 清理列名
            df = self.clean_column_names(df)
            
            # 填充'部件/图纸号及版本'列的空值
            if '部件/图纸号及版本' in df.columns:
                df['部件/图纸号及版本'] = df['部件/图纸号及版本'].ffill()
            
            # 标准化厚度/材质列的格式
            thickness_columns = ['厚度t1/材质', '厚度t2/材质']
            for col in thickness_columns:
                if col in df.columns:
                    df[col] = df[col].apply(self.standardize_thickness_material)
            
            self.data = df
            return True
            
        except Exception as e:
            print(f"加载Excel文件时发生错误: {str(e)}")
            return False
    
    def extract_row_data(self, row_index: int = 0) -> Optional[Dict[str, Any]]:
        """
        提取指定行数据并转换为字典
        
        Args:
            row_index: 要提取的行索引（默认为0，即第一行）
            
        Returns:
            Dict: 包含指定行数据的字典，如果没有数据则返回None
        """
        if self.data is None or len(self.data) <= row_index:
            return None
        
        # 定义需要提取的列名
        target_columns = [
            'WPS', '焊接工艺', '接头类型', '焊接位置', 'WPQR', 
            '厚度t1/材质', '厚度t2/材质', '接头坡口形式', '焊接填充材料','保护气体类型'
        ]
        
        # 提取指定行数据并存储到字典中
        row_dict = {}
        for col in target_columns:
            if col in self.data.columns:
                value = self.data.iloc[row_index][col]
                # 处理NaN值
                if pd.isna(value):
                    row_dict[col] = None
                else:
                    row_dict[col] = str(value)
            else:
                row_dict[col] = None
        
        self.parsed_dict = row_dict
        self.current_row_index = row_index
        return row_dict
    
    def extract_first_row_data(self) -> Optional[Dict[str, Any]]:
        """
        提取第一行数据并转换为字典
        
        Returns:
            Dict: 包含第一行数据的字典，如果没有数据则返回None
        """
        return self.extract_row_data(0)
    
    def extract_next_row_data(self) -> Optional[Dict[str, Any]]:
        """
        提取下一行数据并转换为字典
        
        Returns:
            Dict: 包含下一行数据的字典，如果没有更多数据则返回None
        """
        next_index = self.current_row_index + 1
        return self.extract_row_data(next_index)
    
    def has_next_row(self) -> bool:
        """
        检查是否还有下一行数据
        
        Returns:
            bool: 如果还有下一行数据返回True，否则返回False
        """
        if self.data is None:
            return False
        return (self.current_row_index + 1) < len(self.data)
    
    def get_total_rows(self) -> int:
        """
        获取总行数
        
        Returns:
            int: 总行数
        """
        if self.data is None:
            return 0
        return len(self.data)
    
    def get_current_row_index(self) -> int:
        """
        获取当前行索引
        
        Returns:
            int: 当前行索引
        """
        return self.current_row_index
    
    def get_all_data(self) -> Optional[pd.DataFrame]:
        """获取完整的DataFrame数据"""
        return self.data
    
    def get_parsed_dict(self) -> Optional[Dict[str, Any]]:
        """获取解析后的数据字典"""
        return self.parsed_dict
    
    def format_data_for_prompt(self) -> str:
        """
        将解析的数据格式化为适合作为提示词的字符串
        
        Returns:
            str: 格式化后的数据字符串
        """
        if self.parsed_dict is None:
            return "未解析到Excel数据"
        
        formatted_lines = ["焊接参数信息："]
        for key, value in self.parsed_dict.items():
            if value is not None and str(value).strip():
                formatted_lines.append(f"- {key}: {value}")
        
        return "\n".join(formatted_lines)
    
    def parse_file(self, file_path: str) -> Optional[Dict[str, Any]]:
        """
        一键解析Excel文件并返回第一行数据字典
        
        Args:
            file_path: Excel文件路径
            
        Returns:
            Dict: 解析后的数据字典，失败返回None
        """
        if self.load_excel_data(file_path):
            return self.extract_first_row_data()
        return None