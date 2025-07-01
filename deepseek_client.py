import json
from typing import Optional, Dict, Any
from openai import OpenAI
from wps_calculator import WPSCalculator

class DeepSeekClient:
    def __init__(self, api_key: str, base_url: str = "https://api.deepseek.com"):
        """
        初始化DeepSeek客户端
        
        Args:
            api_key: DeepSeek API密钥
            base_url: API基础URL
        """
        self.client = OpenAI(
            api_key=api_key,
            base_url=base_url
        )
        self.conversation_history = []
        self.wps_calculator = WPSCalculator()  # 初始化焊接工艺参数计算器
        self.system_prompt = '''请基于以下输入数据字典，按照指定的映射规则生成焊接工艺参数，输出格式必须为JSON对象结构：

映射规则说明：
- 工艺规程编号：直接提取输入数据中"WPS"字段的值
- 工艺编号：格式为"[焊接工艺数字] P/P [接头类型] [合金代码] S t[t1厚度]+t[t2厚度] [焊接位置]"
    * [焊接工艺数字]：提取"焊接工艺"字段中括号前的数字部分
    * P/P：固定值
    * [接头类型]：直接提取"接头类型"字段的值
    * [合金代码]：默认是23.1，代表6系铝合金（6005和6082），也就是母材1或2牌号开头的前四个字符
    * S：固定值
    * t[t1厚度]+t[t2厚度]：提取厚度t1/材质和厚度t2/材质中的厚度数字部分
    * [焊接位置]：直接提取"焊接位置"字段的值
- 接头：直接提取"接头坡口形式"字段的值
- 工艺评定名称：直接提取"WPQR"字段的值
- 焊接过程：格式为"[焊后缀][焊接工艺数字] [[焊接方法]] [[焊接标准]]"
    * [焊后缀]：提取"焊接工艺"字段括号内-后面的字母（如MIG-t中的t）
    * [焊接工艺数字]：提取"焊接工艺"字段中括号前的数字部分
    * [焊接方法]：提取"焊接工艺"字段括号内的英文部分（如MIG）
    * [焊接标准]：固定填写ISO 4063
- 接头名称：固定输出"角接接头"（暂无规则）
- 母材1牌号：格式为"[材质] [标准]"
    * [材质]：提取"厚度t1/材质"字段中材质部分（去除厚度和mm）
    * [标准]：固定值TB/T3260.1-2011（该值目前代表动车铁标）
- 母材2牌号：格式为"[材质] [标准]"
    * [材质]：提取"厚度t2/材质"字段中材质部分（去除厚度和mm）
    * [标准]：固定值TB/T3260.1-2011（该值目前代表动车铁标）
- 母材厚度：格式为"[t1厚度]/[t2厚度]"
    * [t1厚度]：提取"厚度t1/材质"字段中的数字部分
    * [t2厚度]：提取"厚度t2/材质"字段中的数字部分
- 焊接接头形式参数：格式为"单位:mm
t1: [t1厚度]
t2: [t2厚度]"
    * t1厚度：提取"厚度t1/材质"字段中的数字部分
    * t2厚度：提取"厚度t2/材质"字段中的数字部分
- 焊角厚度：根据"接头坡口形式"计算
    * 如果包含"a"：直接使用a后的数字（如a3则为3）
    * 如果包含"z"：使用z后的数字乘以0.7（如z4则为4*0.7=2.8）
- 焊接位置：直接提取输入数据中"焊接位置"字段的值
- 焊前准备：固定值"用清洗剂去除油污/用打磨方法去除氧化膜"
- 焊接准备细节：根据接头类型和厚度判断
    * 角接：值为"焊前装配间隙要求为0mm，最大不超过1mm"
    * 对接：参考t1和t2厚度的较小值
        - 厚度≤8mm：值为"焊前装配间隙要求为0mm，最大不超过3mm"
        - 厚度>8mm：值为"焊前装配间隙要求为0mm，最大不超过4mm"
    * 搭接：值为"/"
- 填充金属类别：固定值"S"
- 层道：根据"接头坡口形式"判断
    * a3、a4、a5：值为"1"
    * a5：值为"1\n2"（1和2之间回车换行）
    * a8：值为"1\n2\n3"（1，2，3之间回车换行）
    * 2V、3V、4V：值为"1\n2"（1和2之间回车换行）
- 填充金属名称：根据层道的行数匹配
    * 单行：值为"ISO 18273-S Al 5087 [AlMg4.5MnZr]"
    * 双行：值为"ISO 18273-S Al 5087 [AlMg4.5MnZr]\nISO 18273-S Al 5087 [AlMg4.5MnZr]"
    * 三行：值为"ISO 18273-S Al 5087 [AlMg4.5MnZr]\nISO 18273-S Al 5087 [AlMg4.5MnZr]\nISO 18273-S Al 5087 [AlMg4.5MnZr]"
- 焊材烘干规定：固定值"/"
- 保护气体：格式为"[工艺标准] [[保护气体类型]]"
    * [工艺标准]：固定值"ISO14175-I1"
    * [保护气体类型]：直接提取输入中"保护气体类型"字段的值
- 保护气体流量：根据保护气体类型判断
    * 纯Ar气(99.99%Ar)：值为"14~16 [内直径为Φ13的喷嘴]"
    * 二元(30%He+70%Ar)或三元气体（30%He+70%Ar+0.015%N2）：值为"16~18 [内直径为Φ13的喷嘴]"
- 根部保护气体：固定值"/"
- 根部保护气体流量：固定值"/"
- 预热温度：固定值"80~100"
- 层间温度：固定值"/"
- 焊后热处理：固定值"/"
- 加热和冷却速度：固定值"/"
- 焊丝干伸长度：固定值"12~15"
- 摆动：固定值"/"
- 脉冲焊接情况：固定值"脉冲焊"
- 等离子焊接情况：固定值"/"
- 焊枪角度：固定值"前倾角0~10°"
- 焊工或操作者：固定值"/"
- 证书名称：固定值"/"
- 接头长度：固定值"/"
- 根部开槽衬垫情况：固定值"/"
- 焊接工艺参数：二维数组格式，表头固定，数据行根据焊接位置生成

### 焊接工艺参数表格填写规则

**第一列（焊道）：**
- 根据焊接位置中P开头的组合数量确定表格行数
- 命名格式：1-[位置]（如：1-PB、1-PD、1-PF）
- 目前默认层级为1层，预留多层焊道扩展

**各列填写内容：**
1. 焊道：根据焊接位置生成（如上述规则）
2. 焊接方法：[焊后缀][焊接工艺数字]
3. 焊材规格：固定为"Φ1.2"
4. 电流强度：[电流强度小]~[电流强度大]
5. 电弧电压：[电弧电压小]~[电弧电压大]
6. 电流种类/极性：固定为"DCEP/+"
7. 送丝速度：[送丝速度小]~[送丝速度大]
8. 焊接速度：[焊接速度小]~[焊接速度大]
9. 热输入：[热输入小]~[热输入大]

输出JSON结构：
```json
{
    "工艺规程编号": string,
    "工艺编号": string,
    "接头": string,
    "工艺评定名称": string,
    "焊接过程": string,
    "接头名称": string,
    "母材1牌号": string,
    "母材2牌号": string,
    "母材厚度": string,
    "焊接接头形式参数": string,
    "焊角厚度": string,
    "焊接位置": string,
    "焊前准备": string,
    "焊接准备细节": string,
    "填充金属类别": string,
    "层道": string,
    "填充金属名称": string,
    "焊材烘干规定": string,
    "保护气体": string,
    "保护气体流量": string,
    "根部保护气体": string,
    "根部保护气体流量": string,
    "预热温度": string,
    "层间温度": string,
    "焊后热处理": string,
    "加热和冷却速度": string,
    "焊丝干伸长度": string,
    "摆动": string,
    "脉冲焊接情况": string,
    "等离子焊接情况": string,
    "焊枪角度": string,
    "焊工或操作者": string,
    "证书名称": string,
    "接头长度": string,
    "根部开槽衬垫情况": string,
    "焊接工艺参数": [
        ["焊道", "焊接方法", "焊材规格(mm)", "电流强度(A)", "电弧电压(V)", "电流种类/极性", "送丝速度(m/min)", "焊接速度*(mm/s)", "热输入*(KJ/mm)"],
        // 数据行根据焊接位置动态生成
    ]
}
```

要求：
1. 严格按照上述映射规则进行数据转换
2. 对于标注"暂无规则"的字段，使用示例中的固定值
3. 焊角厚度计算规则：
   - 接头坡口形式包含"a"时：取a后数字作为焊角厚度
   - 接头坡口形式包含"z"时：取z后数字乘以0.7作为焊角厚度
4. 焊接准备细节判断规则：
   - 角接类型：使用"焊前装配间隙要求为0mm，最大不超过1mm"
   - 对接类型：根据t1和t2厚度较小值判断
   - 搭接类型：使用"/"
5. 焊接过程中的焊后缀提取：从焊接工艺字段括号内-后提取字母
6. 焊接工艺参数表中的焊接方法使用提取的焊后缀+工艺数字组合
7. 输出必须是一个完整的JSON对象，不要包含```json标记
8. 不要输出JSON以外的任何内容，例如思维链或解释文字
9. 确保所有字符串值正确提取和格式化
10. 焊接工艺参数必须按照新的计算规则生成，确保数值准确
11. 层道字段根据接头坡口形式生成相应行数，填充金属名称需要匹配层道的行数
12. 保护气体流量根据保护气体类型进行判断和填写

## 修改模式说明
如果用户提供了修改指令，请按照以下要求执行：
1. 仅对用户明确要求修改的字段或参数进行调整
2. 未提及的字段和参数保持原有数据不变
3. 严格按照用户的修改要求进行精确修改，不要自行推断或扩展
4. 修改后仍需遵循原有的数据格式和结构要求
5. 最终输出完整的JSON对象，确保所有字段完整
6. 不要输出修改说明、解释文字或其他额外内容
7. 修改后的输出结果必须同样依照上述的完整JSON对象格式

请基于输入的数据字典生成对应的焊接工艺参数JSON对象。'''
    
    def set_system_prompt(self, prompt: str):
        """设置系统提示词"""
        self.system_prompt = prompt
    
    def reset_conversation(self):
        """重置对话历史"""
        self.conversation_history = []
    
    def chat(self, message: str, excel_data: Optional[Dict[str, Any]] = None, stream: bool = False) -> str:
        """
        发送消息到DeepSeek API
        
        Args:
            message: 用户消息
            excel_data: Excel解析的数据字典
            stream: 是否使用流式响应
            
        Returns:
            AI回复内容
        """
        # 构建完整的用户消息
        full_message = message
        if excel_data:
            # 使用焊接工艺参数计算器处理Excel数据
            try:
                processed_data = self.wps_calculator.process_excel_data(excel_data)
                excel_info = "\n\n以下是Excel文件解析的焊接参数数据（已包含计算的工艺参数）：\n"
                excel_info += json.dumps(processed_data, ensure_ascii=False, indent=2)
                
                # 添加计算参数的说明
                calc_params = {k: v for k, v in processed_data.items() if k not in excel_data}
                if calc_params:
                    excel_info += "\n\n计算得出的焊接工艺参数说明：\n"
                    excel_info += self.wps_calculator.format_parameters_for_display(calc_params)
                    
            except Exception as e:
                print(f"计算焊接工艺参数时发生错误: {str(e)}")
                # 如果计算失败，使用原始数据
                excel_info = "\n\n以下是Excel文件解析的焊接参数数据：\n"
                excel_info += json.dumps(excel_data, ensure_ascii=False, indent=2)
                
            full_message += excel_info
        
        # 构建消息列表
        messages = [
            {"role": "system", "content": self.system_prompt}
        ]
        
        # 添加历史对话
        messages.extend(self.conversation_history)
        
        # 添加当前消息
        messages.append({"role": "user", "content": full_message})
        
        try:
            response = self.client.chat.completions.create(
                model="deepseek-reasoner",
                messages=messages,
                stream=stream,
                temperature=0.2,
                max_tokens=8192
            )
            
            if stream:
                # 处理流式响应 - 支持deepseek-reasoner的思维链
                reasoning_content = ""
                content = ""
                reasoning_started = False  # 标记是否已开始显示思维链
                answer_started = False     # 标记是否已开始显示最终回答
                
                for chunk in response:
                    # 处理思维链内容
                    if chunk.choices[0].delta.reasoning_content:
                        reasoning_part = chunk.choices[0].delta.reasoning_content
                        reasoning_content += reasoning_part
                        
                        # 只在第一次显示思维链标识
                        if not reasoning_started:
                            print("\n[思维过程]\n", end="", flush=True)
                            reasoning_started = True
                        
                        # 显示思维链内容
                        print(reasoning_part, end="", flush=True)
                        
                    # 处理最终回答内容
                    elif chunk.choices[0].delta.content is not None:
                        content_part = chunk.choices[0].delta.content
                        content += content_part
                        
                        # 只在第一次显示最终回答标识
                        if not answer_started:
                            # 如果有思维链，先换行分隔
                            if reasoning_content:
                                print("\n\n[最终回答]\n", end="", flush=True)
                            answer_started = True
                        
                        # 显示最终回答内容
                        print(content_part, end="", flush=True)
                print()  # 最后换行
                
                # 组合完整响应（思维链 + 最终回答）
                full_response = ""
                if reasoning_content:
                    full_response += f"[思维过程]\n{reasoning_content}\n\n"
                if content:
                    full_response += f"[最终回答]\n{content}"
                
                # 保存到对话历史（只保存最终回答部分）
                self.conversation_history.append({"role": "user", "content": message})
                self.conversation_history.append({"role": "assistant", "content": content})
                
                return full_response
            else:
                # 处理非流式响应
                assistant_message = response.choices[0].message.content
                
                # 保存到对话历史
                self.conversation_history.append({"role": "user", "content": message})
                self.conversation_history.append({"role": "assistant", "content": assistant_message})
                
                return assistant_message
                
        except Exception as e:
            error_msg = f"调用DeepSeek API时发生错误: {str(e)}"
            print(error_msg)
            return error_msg
    
    def get_last_response(self) -> Optional[str]:
        """获取最后一次AI回复"""
        if self.conversation_history and self.conversation_history[-1]["role"] == "assistant":
            return self.conversation_history[-1]["content"]
        return None
    
    def get_conversation_history(self) -> list:
        """获取完整对话历史"""
        return self.conversation_history.copy()
    
    def generate_document(self, template_path: str, save_path: str, json_text: str, image_path: str = None, selected_images: dict = None):
        """生成文档
        
        Args:
            template_path: 模板文件路径
            save_path: 保存文件路径
            json_text: JSON格式的数据文本
            image_path: 图片路径（可选，兼容旧版本）
            selected_images: 从图片素材库选择的图片信息（可选）
        """
        from data_loader import LLMDataLoader
        import os
        
        # 创建数据加载器
        data_loader = LLMDataLoader(json_text)
        data = data_loader.load_data()
        
        if data is None:
            raise ValueError("无法解析数据")
        
        # 优先使用图片素材库选择的图片
        if selected_images:
            # 处理图片素材库选择的图片
            for key, image_info in selected_images.items():
                category = image_info['category']
                number = image_info['number']
                
                # 添加焊接接头形式图片
                if image_info['joint_form']:
                    joint_form_key = f"{category}焊接接头形式{number}"
                    data[joint_form_key] = ("", image_info['joint_form'])
                    # 同时添加通用的焊接接头形式标签
                    if "焊接接头形式" not in data:
                        data["焊接接头形式"] = ("", image_info['joint_form'])
                
                # 添加焊接顺序图片
                if image_info['sequence']:
                    sequence_key = f"{category}焊接顺序{number}"
                    data[sequence_key] = ("", image_info['sequence'])
                    # 同时添加通用的焊接顺序标签
                    if "焊接顺序" not in data:
                        data["焊接顺序"] = ("", image_info['sequence'])
        
        # 兼容旧版本的图片路径参数
        elif image_path:
            # 添加焊接接头形式图片
            welding_joint_image = os.path.join(image_path, "焊接接头形式.png")
            data["焊接接头形式"] = ("", welding_joint_image)
            
            # 添加焊接顺序图片
            welding_sequence_image = os.path.join(image_path, "焊接顺序.png")
            data["焊接顺序"] = ("", welding_sequence_image)
        
        # 调用main.py中的match函数生成文档
        from main import match
        match(template_path, save_path, data)