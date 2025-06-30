# 焊接工艺规程生成器 V3

## 项目简介

这是一个基于DeepSeek API的焊接工艺规程文档生成器，支持Excel数据导入和智能文档生成。

## 主要功能

- 🤖 **AI对话**: 基于DeepSeek API的智能对话系统
- 📊 **Excel导入**: 支持焊接参数Excel文件解析和导入
- 📝 **文档生成**: 自动生成标准化的焊接工艺规程文档
- 🎨 **图形界面**: 友好的GUI界面，操作简单直观
- 📋 **提示词管理**: 内置提示词模板库，支持自定义提示词

## 重构内容

### V3版本更新

1. **API替换**: 从ragflow迁移到DeepSeek官方API
2. **Excel集成**: 整合test2.py的Excel解析功能到主界面
3. **界面优化**: 新增Excel数据导入模块
4. **代码重构**: 模块化设计，提高代码可维护性

### 新增模块

- `deepseek_client.py`: DeepSeek API客户端
- `excel_parser.py`: Excel文件解析器
- 更新的GUI界面支持Excel数据导入

## 安装依赖

```bash
pip install -r requirements.txt
```

## 配置说明

### 1. DeepSeek API配置

在 `main.py` 中修改以下配置：

```python
DEEPSEEK_API_KEY = "your-deepseek-api-key"  # 替换为你的API Key
DEEPSEEK_BASE_URL = "https://api.deepseek.com"
```

### 2. 获取DeepSeek API Key

1. 访问 [DeepSeek官网](https://platform.deepseek.com/)
2. 注册账号并登录
3. 在API管理页面创建新的API Key
4. 将API Key填入配置文件

## 使用方法

### 1. 启动程序

```bash
python main.py
```

### 2. Excel数据导入

1. 点击"Excel数据导入"区域的"浏览"按钮
2. 选择包含焊接参数的Excel文件
3. 点击"解析Excel"按钮解析数据
4. 可以点击"预览数据"查看解析结果

### 3. 生成文档

1. 选择合适的提示词模板
2. 在输入框中描述需求
3. 点击"发送"与AI对话
4. 满意后点击"生成文档"创建Word文档

## Excel文件格式要求

支持的Excel文件应包含以下字段（第3行为标题行，第3个工作表）：

- WPS: 焊接工艺规程编号
- 焊接工艺: 焊接工艺类型
- 接头类型: 焊接接头类型
- 焊接位置: 焊接位置代码
- WPQR: 焊接工艺评定记录
- 缺欠质量等级: 质量等级要求
- 厚度t1/材质: 第一材料厚度和材质
- 厚度t2/材质: 第二材料厚度和材质
- 接头坡口形式: 坡口形式描述

## 项目结构

```
demo-V3/
├── main.py                    # 主程序入口
├── deepseek_client.py         # DeepSeek API客户端
├── excel_parser.py            # Excel解析器
├── document_generator_gui.py  # GUI界面
├── doc_processor.py           # 文档处理器
├── template_analyzer.py       # 模板分析器
├── data_loader.py             # 数据加载器
├── requirements.txt           # 依赖包列表
├── data/                      # 数据文件夹
│   ├── 焊接规程书模板.docx    # 文档模板
│   ├── prompt_templates.json  # 提示词模板
│   └── *.xlsx                 # Excel数据文件
├── generated_docs/            # 生成的文档
└── helper/                    # 辅助工具模块
```

## 注意事项

1. **API费用**: DeepSeek API按使用量计费，请注意控制使用成本
2. **网络连接**: 需要稳定的网络连接访问DeepSeek API
3. **文件格式**: Excel文件需要符合指定的格式要求
4. **API限制**: 注意API的调用频率限制

## 故障排除

### 常见问题

1. **API连接失败**
   - 检查网络连接
   - 验证API Key是否正确
   - 确认API余额是否充足

2. **Excel解析失败**
   - 检查文件格式是否正确
   - 确认工作表索引和标题行位置
   - 验证必要字段是否存在

3. **文档生成失败**
   - 检查模板文件是否存在
   - 确认保存目录权限
   - 验证AI回复内容格式

## 开发说明

### 扩展功能

- 可以在 `excel_parser.py` 中修改字段映射
- 可以在 `deepseek_client.py` 中调整AI参数
- 可以在GUI中添加更多操作按钮

### 自定义模板

- 修改 `data/焊接规程书模板.docx` 自定义文档模板
- 在 `data/prompt_templates.json` 中添加新的提示词

## 版本历史

- **V3.0**: DeepSeek API集成，Excel导入功能
- **V2.0**: ragflow集成版本
- **V1.0**: 基础版本