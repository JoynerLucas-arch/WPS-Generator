# 焊接工艺规程生成器 V4

## 项目简介

这是一个基于DeepSeek API的智能焊接工艺规程文档生成器，集成了Excel数据导入、焊接工艺参数自动计算和AI文档生成功能，为焊接工程师提供一站式的工艺规程制作解决方案。

## 主要功能

- 🤖 **AI对话**: 基于DeepSeek API的智能对话系统
- 📊 **Excel导入**: 支持焊接参数Excel文件解析和导入
- 🧮 **参数计算**: 智能焊接工艺参数计算器，根据材料厚度自动计算焊接参数
- 📝 **文档生成**: 自动生成标准化的焊接工艺规程文档
- 🎨 **图形界面**: 友好的GUI界面，操作简单直观
- 📋 **提示词管理**: 内置提示词模板库，支持自定义提示词
- 🖼️ **图像支持**: 丰富的焊接接头形式和焊接顺序图库
- 📦 **可执行文件**: 支持PyInstaller打包为独立可执行文件

## 重构内容

### 版本更新

1. **API替换**: 从ragflow迁移到DeepSeek官方API
2. **Excel集成**: 整合Excel解析功能到主界面
3. **参数计算**: 新增焊接工艺参数自动计算功能
4. **界面优化**: 新增Excel数据导入模块和参数计算显示
5. **代码重构**: 模块化设计，提高代码可维护性
6. **图像资源**: 完善的焊接接头图像库

### 新增模块

- `deepseek_client.py`: DeepSeek API客户端
- `excel_parser.py`: Excel文件解析器
- `wps_calculator.py`: 焊接工艺参数计算器
- `labels.py`: 标签和常量定义
- `helper/`: 辅助工具模块集合
- 更新的GUI界面支持Excel数据导入和参数计算

## 安装依赖

```bash
pip install -r requirements.txt
```
### 依赖包说明
- openai>=1.0.0 : DeepSeek API客户端
- pandas>=1.5.0 : Excel数据处理
- openpyxl>=3.0.0 : Excel文件读写
- python-docx>=0.8.11 : Word文档处理
- Pillow>=9.0.0 : 图像处理
- requests>=2.28.0 : HTTP请求

## 配置说明
### 1. DeepSeek API配置
方法一：环境变量配置（推荐）
```bash
# Windows
set 
DEEPSEEK_API_KEY=your-deepseek-api-key

# Linux/Mac
export 
DEEPSEEK_API_KEY=your-deepseek-api-key
```
方法二：直接修改代码 在 main.py 中修改以下配置：
```python
DEEPSEEK_API_KEY = "your-deepseek-api-key"  # 替换为你的API Key
DEEPSEEK_BASE_URL = "https://api.deepseek.com"
```

### 2. 获取DeepSeek API Key
1. 访问 DeepSeek官网
2. 注册账号并登录
3. 在API管理页面创建新的API Key
4. 将API Key配置到环境变量或代码中
## 使用方法
### 1. 启动程序
python main.py
### 2. Excel数据导入
1. 点击"Excel数据导入"区域的"浏览"按钮
2. 选择包含焊接参数的Excel文件
3. 点击"解析Excel"按钮解析数据
4. 系统会自动计算焊接工艺参数
5. 可以点击"预览数据"查看解析和计算结果
### 3. 焊接工艺参数计算
系统会根据Excel中的材料厚度信息自动计算：

- 电流强度范围
- 电弧电压范围
- 送丝速度范围
- 焊接速度范围
- 热输入范围
### 4. 生成文档
1. 选择合适的提示词模板
2. 在输入框中描述需求
3. 点击"发送"与AI对话
4. 满意后点击"生成文档"创建Word文档
### 5. 打包可执行文件
```bash
pyinstaller app.spec
```
生成的可执行文件位于 dist/ 目录下。

## Excel文件格式要求
支持的Excel文件应包含以下字段（第3行为标题行，第3个工作表）：

- WPS : 焊接工艺规程编号
- 焊接工艺 : 焊接工艺类型
- 接头类型 : 焊接接头类型
- 焊接位置 : 焊接位置代码
- WPQR : 焊接工艺评定记录
- 缺欠质量等级 : 质量等级要求
- 厚度t1/材质 : 第一材料厚度和材质（如："3mm 6005A-T6"）
- 厚度t2/材质 : 第二材料厚度和材质
- 接头坡口形式 : 坡口形式描述
## 焊接工艺参数计算规则
### 厚度分层
- 0-3mm : 电流160-180A，电压22-23.2V，送丝速度10.1-11.6m/min
- 3-5mm : 电流180-200A，电压23.2-23.5V，送丝速度11.6-12.7m/min
- 5mm以上 : 电流200-220A，电压23.5-23.8V，送丝速度12.7-13.8m/min
### 计算公式
- 焊接速度 = (0.8 × 电流 × 电压 × 0.001) ÷ 热输入
- 热输入 = (0.8 × 电流 × 电压 × 0.001) ÷ 焊接速度
## 项目结构
```
WPS样机(API)/
├── main.py                    # 主程序入口
├── deepseek_client.py         # DeepSeek API客户端
├── excel_parser.py            # Excel解析器
├── wps_calculator.py          # 焊接工艺参数计算器
├── document_generator_gui.py  # GUI界面
├── doc_processor.py           # 文档处理器
├── template_analyzer.py       # 模板分析器
├── data_loader.py             # 数据加载器
├── labels.py                  # 标签和常量定义
├── llm_response.py            # LLM响应处理
├── requirements.txt           # 依赖包列表
├── app.spec                   # PyInstaller配置文件
├── 焊接工艺参数计算器说明.md   # 计算器详细说明
├── data/                      # 数据文件夹
│   ├── 焊接规程书模板.docx    # 文档模板
│   ├── prompt_templates.json  # 提示词模板
│   ├── 焊接工艺规程.png       # 示例图片
│   └── 底架焊接接头清单.xlsx  # 示例Excel文件
├── generated_docs/            # 生成的文档
│   └── 焊接工艺规程/          # 焊接工艺规程文档
├── imgs/                      # 焊接接头图像库
│   ├── 板T形接头/             # T形接头图像
│   ├── 板对接接头/            # 对接接头图像
│   ├── 板搭接接头/            # 搭接接头图像
│   ├── 管板对接/              # 管板对接图像
│   └── 角接接头/              # 角接接头图像
├── helper/                    # 辅助工具模块
│   ├── docx_helper.py         # Word文档辅助工具
│   ├── image_size.py          # 图像尺寸处理
│   ├── os_helper.py           # 操作系统辅助工具
│   ├── output_redirector.py   # 输出重定向
│   └── type_helper.py         # 类型辅助工具
├── build/                     # 构建文件
└── dist/                      # 打包输出文件
```

## 注意事项
1. API费用 : DeepSeek API按使用量计费，请注意控制使用成本
2. 网络连接 : 需要稳定的网络连接访问DeepSeek API
3. 文件格式 : Excel文件需要符合指定的格式要求
4. API限制 : 注意API的调用频率限制
5. 环境变量 : 推荐使用环境变量配置API Key，避免代码泄露
6. 图像资源 : 确保imgs目录下的图像文件完整

## 故障排除
### 常见问题
1. API连接失败
   
   - 检查网络连接
   - 验证API Key是否正确
   - 确认API余额是否充足
   - 检查环境变量配置
2. Excel解析失败
   
   - 检查文件格式是否正确
   - 确认工作表索引和标题行位置
   - 验证必要字段是否存在
   - 检查材料厚度格式（如："3mm 6005A-T6"）
3. 参数计算异常
   
   - 确认材料厚度信息格式正确
   - 检查厚度值是否为有效数字
   - 验证计算公式参数范围
4. 文档生成失败
   
   - 检查模板文件是否存在
   - 确认保存目录权限
   - 验证AI回复内容格式
5. 打包问题
   
   - 确保所有依赖包已安装
   - 检查app.spec配置文件
   - 验证资源文件路径
## 开发说明
### 扩展功能
- 可以在 excel_parser.py 中修改字段映射
- 可以在 deepseek_client.py 中调整AI参数
- 可以在 wps_calculator.py 中修改计算规则
- 可以在GUI中添加更多操作按钮
- 可以在 imgs/ 目录中添加新的接头图像
### 自定义模板
- 修改 data/焊接规程书模板.docx 自定义文档模板
- 在 data/prompt_templates.json 中添加新的提示词
- 在 labels.py 中定义新的标签常量
### 参数计算定制
- 在 wps_calculator.py 中修改厚度分层规则
- 调整计算公式和参数范围
- 添加新的材料类型支持
## 版本历史
- V3.0 : DeepSeek API集成，Excel导入功能，焊接工艺参数计算器
- V2.0 : ragflow集成版本
- V1.0 : 基础版本
## 许可证
本项目仅供学习和研究使用。

## 联系方式
如有问题或建议，请通过项目Issues反馈。