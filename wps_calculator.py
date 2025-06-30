# -*- coding: utf-8 -*-

import math
import re
from typing import Dict, Any, Tuple, Optional

class WPSCalculator:
    """焊接工艺参数计算器"""
    
    def __init__(self):
        # 基础参数配置表
        self.thickness_ranges = {
            "0-3mm": {
                "current_min": 160,
                "current_max": 180,
                "voltage_min": 22.0,
                "voltage_max": 23.2,
                "wire_speed_min": 10.1,
                "wire_speed_max": 11.6,
                "reference_heat_input": 0.18
            },
            "3-5mm": {
                "current_min": 180,
                "current_max": 200,
                "voltage_min": 23.2,
                "voltage_max": 23.5,
                "wire_speed_min": 11.6,
                "wire_speed_max": 12.7,
                "reference_heat_input": 0.42
            },
            "5mm+": {
                "current_min": 200,
                "current_max": 220,
                "voltage_min": 23.5,
                "voltage_max": 23.8,
                "wire_speed_min": 12.7,
                "wire_speed_max": 13.8,
                "reference_heat_input": 0.43
            }
        }
    
    def extract_thickness_from_material(self, material_str: str) -> Optional[float]:
        """从材质字符串中提取厚度值
        
        Args:
            material_str: 格式如"3mm 6005A-T6"的字符串
            
        Returns:
            float: 厚度值，如果提取失败返回None
        """
        if not material_str or material_str == "None":
            return None
            
        # 使用正则表达式提取数字部分
        pattern = r'(\d+(?:\.\d+)?)mm'
        match = re.search(pattern, str(material_str))
        
        if match:
            return float(match.group(1))
        return None
    
    def get_min_thickness(self, t1_material: str, t2_material: str) -> Optional[float]:
        """获取两个材质厚度的较小值
        
        Args:
            t1_material: t1材质字符串
            t2_material: t2材质字符串
            
        Returns:
            float: 较小的厚度值
        """
        t1 = self.extract_thickness_from_material(t1_material)
        t2 = self.extract_thickness_from_material(t2_material)
        
        if t1 is None and t2 is None:
            return None
        elif t1 is None:
            return t2
        elif t2 is None:
            return t1
        else:
            return min(t1, t2)
    
    def get_thickness_range_key(self, thickness: float) -> str:
        """根据厚度值获取对应的参数范围键
        
        Args:
            thickness: 厚度值
            
        Returns:
            str: 参数范围键
        """
        if thickness <= 3:
            return "0-3mm"
        elif thickness <= 5:
            return "3-5mm"
        else:
            return "5mm+"
    
    def calculate_welding_parameters(self, excel_data: Dict[str, Any]) -> Dict[str, Any]:
        """计算焊接工艺参数
        
        Args:
            excel_data: Excel解析的数据字典
            
        Returns:
            Dict: 包含计算结果的字典
        """
        # 提取厚度信息
        t1_material = excel_data.get('厚度t1/材质', '')
        t2_material = excel_data.get('厚度t2/材质', '')
        
        min_thickness = self.get_min_thickness(t1_material, t2_material)
        
        if min_thickness is None:
            raise ValueError("无法从材质信息中提取有效的厚度值")
        
        # 获取对应的参数范围
        range_key = self.get_thickness_range_key(min_thickness)
        params = self.thickness_ranges[range_key]
        
        # 基础参数
        current_min = params["current_min"]
        current_max = params["current_max"]
        voltage_min = params["voltage_min"]
        voltage_max = params["voltage_max"]
        wire_speed_min = params["wire_speed_min"]
        wire_speed_max = params["wire_speed_max"]
        reference_heat_input = params["reference_heat_input"]
        
        # 计算焊接速度
        # 焊接速度小 = (0.8 * 电流强度大 * 电弧电压大 * 0.001) / (参考热输入 * 1.25)（按0.5向上取整）
        welding_speed_min_raw = (0.8 * current_max * voltage_max * 0.001) / (reference_heat_input * 1.25)
        welding_speed_min = math.ceil(welding_speed_min_raw * 2) / 2
        
        # 焊接速度大 = (0.8 * 电流强度小 * 电弧电压小 * 0.001) / (参考热输入 * 0.75)（按0.5向下取整）
        welding_speed_max_raw = (0.8 * current_min * voltage_min * 0.001) / (reference_heat_input * 0.75)
        welding_speed_max = math.floor(welding_speed_max_raw * 2) / 2
        
        # 计算热输入
        # 热输入小 = round((0.8 * 电流强度小 * 电弧电压小 * 0.001) / 焊接速度大)（保留两位小数）
        heat_input_min = round(
            (0.8 * current_min * voltage_min * 0.001) / welding_speed_max, 2
        )
        
        # 热输入大 = round((0.8 * 电流强度大 * 电弧电压大 * 0.001) / 焊接速度小)（保留两位小数）
        heat_input_max = round(
            (0.8 * current_max * voltage_max * 0.001) / welding_speed_min, 2
        )
        
        # 构建结果字典
        calculated_params = {
            "电流强度小": current_min,
            "电流强度大": current_max,
            "电弧电压小": voltage_min,
            "电弧电压大": voltage_max,
            "送丝速度小": wire_speed_min,
            "送丝速度大": wire_speed_max,
            "焊接速度小": welding_speed_min,
            "焊接速度大": welding_speed_max,
            "热输入小": heat_input_min,
            "热输入大": heat_input_max,
            "参考厚度": min_thickness,
            "厚度范围": range_key
        }
        
        return calculated_params
    
    def merge_data_for_llm(self, excel_data: Dict[str, Any], calculated_params: Dict[str, Any]) -> Dict[str, Any]:
        """合并Excel数据和计算参数，准备发送给大模型
        
        Args:
            excel_data: Excel解析的原始数据
            calculated_params: 计算得出的焊接参数
            
        Returns:
            Dict: 合并后的完整数据字典
        """
        # 创建合并后的数据字典
        merged_data = excel_data.copy()
        
        # 添加计算参数
        merged_data.update(calculated_params)
        
        return merged_data
    
    def process_excel_data(self, excel_data: Dict[str, Any]) -> Dict[str, Any]:
        """处理Excel数据的主要方法
        
        Args:
            excel_data: Excel解析的数据字典
            
        Returns:
            Dict: 包含原始数据和计算参数的完整字典
        """
        try:
            # 计算焊接工艺参数
            calculated_params = self.calculate_welding_parameters(excel_data)
            
            # 合并数据
            merged_data = self.merge_data_for_llm(excel_data, calculated_params)
            
            return merged_data
            
        except Exception as e:
            print(f"处理Excel数据时发生错误: {str(e)}")
            # 如果计算失败，返回原始数据
            return excel_data
    
    def format_parameters_for_display(self, calculated_params: Dict[str, Any]) -> str:
        """格式化参数用于显示
        
        Args:
            calculated_params: 计算得出的参数字典
            
        Returns:
            str: 格式化后的参数字符串
        """
        lines = ["计算得出的焊接工艺参数："]
        lines.append(f"参考厚度: {calculated_params.get('参考厚度', 'N/A')}mm")
        lines.append(f"厚度范围: {calculated_params.get('厚度范围', 'N/A')}")
        lines.append("")
        lines.append("基础参数：")
        lines.append(f"电流强度: {calculated_params.get('电流强度小', 'N/A')}A ~ {calculated_params.get('电流强度大', 'N/A')}A")
        lines.append(f"电弧电压: {calculated_params.get('电弧电压小', 'N/A')}V ~ {calculated_params.get('电弧电压大', 'N/A')}V")
        lines.append(f"送丝速度: {calculated_params.get('送丝速度小', 'N/A')}m/min ~ {calculated_params.get('送丝速度大', 'N/A')}m/min")
        lines.append("")
        lines.append("计算参数：")
        lines.append(f"焊接速度: {calculated_params.get('焊接速度小', 'N/A')}mm/s ~ {calculated_params.get('焊接速度大', 'N/A')}mm/s")
        lines.append(f"热输入: {calculated_params.get('热输入小', 'N/A')}KJ/mm ~ {calculated_params.get('热输入大', 'N/A')}KJ/mm")
        
        return "\n".join(lines)


# 使用示例和测试函数
if __name__ == "__main__":
    # 创建计算器实例
    calculator = WPSCalculator()
    
    # 测试数据
    test_data = {
        '厚度t1/材质': '4mm 6005A-T6',
        '厚度t2/材质': '8mm 6082-T6',
        'WPS': 'WPS-001',
        '焊接工艺': '131(MIG-t)',
        '接头类型': '角接',
        '焊接位置': 'PB+PD',
        'WPQR': 'WPQR-001',
        '接头坡口形式': 'a3'
    }
    
    # 处理数据
    result = calculator.process_excel_data(test_data)
    
    # 显示结果
    print("原始数据:")
    for key, value in test_data.items():
        print(f"  {key}: {value}")
    
    print("\n" + "="*50)
    
    # 提取计算参数
    calc_params = {k: v for k, v in result.items() if k not in test_data}
    print(calculator.format_parameters_for_display(calc_params))
    
    print("\n" + "="*50)
    print("完整合并数据:")
    for key, value in result.items():
        print(f"  {key}: {value}")