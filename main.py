from document_generator_gui import DocumentGeneratorGUI
from helper.os_helper import *
from data_loader import StaticDataLoader
from template_analyzer import TemplateAnalyzer
from doc_processor import DocumentProcessor
from deepseek_client import DeepSeekClient
import os


def match(file_path: str, save_path: str, datas: dict):

    # 注册每次模板生成过程的静态插入数据
    TemplateAnalyzer.register_static_datas()

    # 模板检查与预处理
    check_result = TemplateAnalyzer.check_template(file_path, DocumentProcessor.insert_data_to_no_content_point)

    # 模板校验失败直接退出
    if check_result['code'].is_error():
        return
    insert_points = check_result['data']['insert_points']
    document = check_result['data']['document']

    # 处理有内容类型插入点，检查并插入数据，返回没有对应数据的插入点
    no_data_points = DocumentProcessor.solve_content_labels(insert_points, datas)

    # 保存文件
    document.save(save_path)


def main():
    # DeepSeek API配置信息
    DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")  # 请替换为实际的API Key
    DEEPSEEK_BASE_URL = "https://api.deepseek.com"
    
    # 模板路径配置
    TEMPLATE_PATHS = {
        "焊接工艺规程": "data/焊接规程书模板.docx"
    }
    
    # 保存路径配置
    SAVE_DIR = {
        "焊接工艺规程": "generated_docs/焊接工艺规程",
    }
    INITIAL_QUESTION = "你好！请问需要生成什么类型的焊接工艺规程？"

    # 创建DeepSeek客户端
    deepseek_client = DeepSeekClient(
        api_key=DEEPSEEK_API_KEY,
        base_url=DEEPSEEK_BASE_URL
    )
    
    # 创建并运行GUI
    app = DocumentGeneratorGUI(
        chat_assistant=deepseek_client,
        template_paths=TEMPLATE_PATHS,
        save_dir=SAVE_DIR["焊接工艺规程"],
        initial_question=INITIAL_QUESTION
    )
    app.run()


if __name__ == "__main__":
    main()
