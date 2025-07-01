import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, simpledialog
import threading
from queue import Queue
import sys
from helper.output_redirector import OutputRedirector
import os
import json
from PIL import Image, ImageTk
from excel_parser import ExcelParser


class DocumentGeneratorGUI:
    def __init__(self, chat_assistant, template_paths: dict, save_dir: str, initial_question: str):
        self.root = tk.Tk()
        self.root.title("焊接规程书生成器")
        self.root.geometry("1200x800")

        self.chat_assistant = chat_assistant
        self.template_paths = template_paths
        self.save_dir = save_dir
        self.current_template = "焊接工艺规程"  # 默认选择工艺卡片
        self.image_path = ""  # 图片路径
        
        # Excel解析器
        self.excel_parser = ExcelParser()
        self.excel_data = None  # 存储解析的Excel数据
        self.excel_file_path = ""  # Excel文件路径
        
        # 确保保存目录存在
        self.ensure_save_directories()
        
        # 模板预览图片路径
        self.template_preview_paths = {
            "焊接工艺规程": "data/焊接工艺规程.png",
        }
        
        # 提示词模板库
        self.prompt_templates = {
            "焊接工艺规程": ["请参照现有知识，生成焊接工艺规程。"],
            # "作业指导书": ["请参照现有电气布线知识，生成CRRC-1型车的电气布线作业指导。"]
        }
        
        # 尝试从文件加载已保存的提示词
        self.prompt_file = "data/prompt_templates.json"
        self.load_prompt_templates()
        
        # 当前选择的提示词
        self.current_prompt = ""
        
        # 存储预览图片的引用
        self.preview_images = {}
        self.current_preview = None
        
        # 图片素材库相关变量
        self.imgs_dir = "imgs"  # 图片素材库目录
        self.selected_images = {}  # 存储用户选择的图片 {category: {joint_type: image_path, sequence: image_path}}
        self.image_categories = ["板T形接头", "板对接接头", "板搭接接头", "管板对接", "角接接头"]

        # Queue for handling output redirection
        self.output_queue = Queue()
        self.old_stdout = sys.stdout
        sys.stdout = OutputRedirector(self.output_queue)

        self.initial_question = initial_question
        self.create_widgets()
        self.setup_output_handling()
        # 显示欢迎信息，但不自动发送消息
        self.output_text.insert(tk.END, "欢迎使用焊接工艺规程生成器！\n")
        self.output_text.insert(tk.END, "请先上传并解析Excel文件，然后选择提示词进行对话。\n\n")

    def create_widgets(self):
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # 创建左侧控制面板
        control_frame = ttk.LabelFrame(main_frame, text="控制面板", padding="5")
        control_frame.grid(row=0, column=0, rowspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        # 创建滚动区域
        # 创建Canvas和Scrollbar
        canvas = tk.Canvas(control_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(control_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        # 配置滚动
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 布局Canvas和Scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 绑定鼠标滚轮事件到主界面Canvas
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def bind_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def unbind_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
        
        canvas.bind('<Enter>', bind_mousewheel)
        canvas.bind('<Leave>', unbind_mousewheel)
        
        # 创建内容区域框架（现在是scrollable_frame的子框架）
        content_frame = ttk.Frame(scrollable_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 模板选择
        template_frame = ttk.Frame(content_frame)
        template_frame.pack(fill=tk.X, pady=5)
        ttk.Label(template_frame, text="选择模板:").pack(side=tk.LEFT)
        self.template_var = tk.StringVar(value=self.current_template)
        template_combo = ttk.Combobox(template_frame, textvariable=self.template_var, values=list(self.template_paths.keys()), state="readonly")
        template_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        template_combo.bind('<<ComboboxSelected>>', self.on_template_change)

        # 模板预览
        preview_frame = ttk.LabelFrame(content_frame, text="模板预览", padding="5")
        preview_frame.pack(fill=tk.X, pady=5)
        
        # 创建预览标签
        self.preview_label = ttk.Label(preview_frame)
        self.preview_label.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 设置预览框的初始最小宽度，以便能够容纳更大的预览图片
        preview_frame.configure(width=340)  # 设置足够容纳320px宽图片的框宽度
        
        # 加载并显示初始预览图片
        self.update_preview_image()

        # Excel文件上传
        excel_frame = ttk.LabelFrame(content_frame, text="Excel数据导入", padding="5")
        excel_frame.pack(fill=tk.X, pady=5)
        
        # Excel文件路径选择
        excel_path_frame = ttk.Frame(excel_frame)
        excel_path_frame.pack(fill=tk.X, pady=2)
        ttk.Label(excel_path_frame, text="Excel文件:").pack(side=tk.LEFT)
        self.excel_path_var = tk.StringVar()
        excel_entry = ttk.Entry(excel_path_frame, textvariable=self.excel_path_var)
        excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        excel_browse_btn = ttk.Button(excel_path_frame, text="浏览", command=self.select_excel_file)
        excel_browse_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        # Excel操作按钮
        excel_btn_frame = ttk.Frame(excel_frame)
        excel_btn_frame.pack(fill=tk.X, pady=2)
        parse_excel_btn = ttk.Button(excel_btn_frame, text="解析Excel", command=self.parse_excel_file)
        parse_excel_btn.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        preview_excel_btn = ttk.Button(excel_btn_frame, text="预览数据", command=self.preview_excel_data)
        preview_excel_btn.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        clear_excel_btn = ttk.Button(excel_btn_frame, text="清除数据", command=self.clear_excel_data)
        clear_excel_btn.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        
        # Excel数据状态显示
        self.excel_status_var = tk.StringVar(value="未导入Excel数据")
        excel_status_label = ttk.Label(excel_frame, textvariable=self.excel_status_var, foreground="gray")
        excel_status_label.pack(fill=tk.X, pady=2)

        # 图片素材库选择
        image_frame = ttk.Frame(content_frame)
        image_frame.pack(fill=tk.X, pady=5)
        image_library_btn = ttk.Button(image_frame, text="图片素材库", command=self.open_image_library)
        image_library_btn.pack(fill=tk.X)
        
        # 提示词模板库
        prompt_frame = ttk.LabelFrame(content_frame, text="提示词模板库", padding="5")
        prompt_frame.pack(fill=tk.X, pady=5)
        
        # 提示词选择区域
        prompt_select_frame = ttk.Frame(prompt_frame)
        prompt_select_frame.pack(fill=tk.X, pady=5)
        ttk.Label(prompt_select_frame, text="选择提示词:").pack(side=tk.LEFT)
        
        # 默认加载当前模板的提示词
        self.prompt_var = tk.StringVar()
        if self.prompt_templates[self.current_template]:
            self.prompt_var.set(self.prompt_templates[self.current_template][0])
            self.current_prompt = self.prompt_templates[self.current_template][0]
        
        self.prompt_combo = ttk.Combobox(prompt_select_frame, textvariable=self.prompt_var, 
                                        values=self.prompt_templates[self.current_template])
        self.prompt_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.prompt_combo.bind('<<ComboboxSelected>>', self.on_prompt_change)
        
        # 提示词操作按钮区 - 分为两行
        prompt_btn_frame1 = ttk.Frame(prompt_frame)
        prompt_btn_frame1.pack(fill=tk.X, pady=(5, 2))
        
        prompt_btn_frame2 = ttk.Frame(prompt_frame)
        prompt_btn_frame2.pack(fill=tk.X, pady=(2, 5))
        
        # 第一行按钮: 使用、添加和预览（新增）
        use_prompt_btn = ttk.Button(prompt_btn_frame1, text="使用此提示词", command=self.use_current_prompt)
        use_prompt_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        add_prompt_btn = ttk.Button(prompt_btn_frame1, text="添加新提示词", command=self.add_prompt_template)
        add_prompt_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        preview_prompt_btn = ttk.Button(prompt_btn_frame1, text="预览完整提示词", command=self.preview_prompt)
        preview_prompt_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 第二行按钮: 删除和管理
        delete_prompt_btn = ttk.Button(prompt_btn_frame2, text="删除当前提示词", command=self.delete_prompt_template)
        delete_prompt_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 可以添加更多管理功能的按钮
        reset_prompt_btn = ttk.Button(prompt_btn_frame2, text="重置为默认提示词", command=self.reset_prompt_templates)
        reset_prompt_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 空白填充区域，确保按钮在底部
        filler = ttk.Frame(content_frame)
        filler.pack(fill=tk.BOTH, expand=True)

        # 按钮区域 - 放在滚动内容的底部
        button_frame = ttk.Frame(content_frame)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        
        self.reset_button = ttk.Button(button_frame, text="重置", command=self.reset_conversation)
        self.reset_button.pack(fill=tk.X, pady=2)
        
        self.send_button = ttk.Button(button_frame, text="发送", command=self.send_message)
        self.send_button.pack(fill=tk.X, pady=2)
        
        self.generate_button = ttk.Button(button_frame, text="生成文档", command=self.generate_document)
        self.generate_button.pack(fill=tk.X, pady=2)

        # Model Output Area
        output_label = ttk.Label(main_frame, text="主面板:")
        output_label.grid(row=0, column=1, sticky=tk.W, pady=(0, 5))

        self.output_text = scrolledtext.ScrolledText(main_frame, height=15, wrap=tk.WORD)
        self.output_text.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))

        # User Input Area
        input_label = ttk.Label(main_frame, text="用户输入:")
        input_label.grid(row=2, column=1, sticky=tk.W, pady=(10, 5))

        self.input_text = scrolledtext.ScrolledText(main_frame, height=8, wrap=tk.WORD)
        self.input_text.grid(row=3, column=1, sticky=(tk.W, tk.E))

        # Bind Enter key to send message
        self.input_text.bind('<Control-Return>', lambda e: self.send_message())

    def load_prompt_templates(self):
        """从文件加载提示词模板"""
        try:
            if os.path.exists(self.prompt_file):
                with open(self.prompt_file, 'r', encoding='utf-8') as f:
                    saved_prompts = json.load(f)
                    # 合并已保存的提示词
                    for template_type, prompts in saved_prompts.items():
                        if template_type in self.prompt_templates:
                            # 保留默认提示词并添加保存的提示词
                            default_prompt = self.prompt_templates[template_type][0]
                            self.prompt_templates[template_type] = [default_prompt]
                            for prompt in prompts:
                                if prompt != default_prompt and prompt not in self.prompt_templates[template_type]:
                                    self.prompt_templates[template_type].append(prompt)
        except Exception as e:
            print(f"加载提示词模板时出错: {str(e)}")
    
    def save_prompt_templates(self):
        """保存提示词模板到文件"""
        try:
            os.makedirs(os.path.dirname(self.prompt_file), exist_ok=True)
            with open(self.prompt_file, 'w', encoding='utf-8') as f:
                json.dump(self.prompt_templates, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存提示词模板时出错: {str(e)}")
    
    def add_prompt_template(self):
        """添加新的提示词模板"""
        # 创建自定义对话框，替代简单的askstring对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("添加提示词")
        dialog.geometry("500x300")  # 设置更大的窗口尺寸
        dialog.transient(self.root)  # 设置为主窗口的临时窗口
        dialog.grab_set()  # 模态窗口
        
        # 添加说明标签
        ttk.Label(dialog, text="请输入新的提示词:", padding=(10, 10)).pack(anchor=tk.W)
        
        # 添加文本输入区域（多行文本框，滚动条）
        text_frame = ttk.Frame(dialog)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        text_scroll = ttk.Scrollbar(text_frame)
        text_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        text_entry = tk.Text(text_frame, height=10, wrap=tk.WORD, yscrollcommand=text_scroll.set)
        text_entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        text_scroll.config(command=text_entry.yview)
        
        # 添加按钮区域
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        result = [None]  # 使用列表存储结果，以便在回调中修改
        
        def on_ok():
            result[0] = text_entry.get("1.0", tk.END).strip()
            dialog.destroy()
            
        def on_cancel():
            dialog.destroy()
        
        ttk.Button(btn_frame, text="确定", command=on_ok).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side=tk.RIGHT)
        
        # 居中显示对话框
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        # 设置焦点
        text_entry.focus_set()
        
        # 等待对话框关闭
        self.root.wait_window(dialog)
        
        # 处理结果
        new_prompt = result[0]
        if new_prompt and new_prompt.strip():
            if new_prompt not in self.prompt_templates[self.current_template]:
                self.prompt_templates[self.current_template].append(new_prompt)
                self.prompt_combo['values'] = self.prompt_templates[self.current_template]
                self.prompt_var.set(new_prompt)
                self.current_prompt = new_prompt
                self.save_prompt_templates()
                messagebox.showinfo("成功", "已添加新的提示词")
            else:
                messagebox.showinfo("提示", "该提示词已存在")
    
    def delete_prompt_template(self):
        """删除当前选择的提示词模板"""
        if not self.current_prompt:
            messagebox.showinfo("提示", "请先选择一个提示词")
            return
            
        # 不允许删除默认提示词
        if self.current_prompt == self.prompt_templates[self.current_template][0]:
            messagebox.showinfo("提示", "不能删除默认提示词")
            return
            
        # 确认删除
        confirm = messagebox.askyesno("确认删除", f"确定要删除提示词: \n\n{self.current_prompt}\n\n吗?")
        if confirm:
            self.prompt_templates[self.current_template].remove(self.current_prompt)
            self.prompt_combo['values'] = self.prompt_templates[self.current_template]
            # 选择第一个提示词
            self.prompt_var.set(self.prompt_templates[self.current_template][0])
            self.current_prompt = self.prompt_templates[self.current_template][0]
            self.save_prompt_templates()
            messagebox.showinfo("成功", "已删除提示词")
    
    def reset_prompt_templates(self):
        """重置为默认提示词"""
        confirm = messagebox.askyesno("确认重置", "确定要重置为默认提示词吗? 这将删除所有自定义提示词。")
        if confirm:
            # 恢复默认提示词
            default_templates = {
                "焊接工艺规程": ["请参照现有知识，生成焊接工艺规程。"]
            }
            self.prompt_templates = default_templates
            self.prompt_combo['values'] = self.prompt_templates[self.current_template]
            self.prompt_var.set(self.prompt_templates[self.current_template][0])
            self.current_prompt = self.prompt_templates[self.current_template][0]
            self.save_prompt_templates()
            messagebox.showinfo("成功", "已重置为默认提示词")
    
    def on_prompt_change(self, event=None):
        """处理提示词切换"""
        selected_prompt = self.prompt_var.get()
        if selected_prompt:
            self.current_prompt = selected_prompt
            # 重置聊天 - 开始新的会话
            self.chat_assistant.reset_conversation()
            # 清空输出区域
            self.output_text.delete("1.0", tk.END)
            # 将提示词填入输入框，但不自动发送
            self.input_text.delete("1.0", tk.END)
            self.input_text.insert("1.0", self.current_prompt)
            # 显示提示信息
            self.output_text.insert(tk.END, f"已选择提示词: {selected_prompt}\n")
            self.output_text.insert(tk.END, "Excel数据将在发送消息时自动包含。\n")
            self.output_text.insert(tk.END, "请点击发送按钮开始对话。\n")
            self.output_text.insert(tk.END, "注意：切换提示词已开始新的对话会话。\n\n")
            self.output_text.see(tk.END)
    
    def preview_prompt(self):
        """预览完整提示词对话框"""
        if not self.current_prompt:
            messagebox.showinfo("提示", "请先选择一个提示词")
            return
            
        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("提示词预览")
        dialog.geometry("600x400")  # 设置较大的窗口尺寸以显示完整内容
        dialog.transient(self.root)  # 设置为主窗口的临时窗口
        dialog.grab_set()  # 模态窗口
        
        # 添加说明标签
        ttk.Label(dialog, text="完整提示词内容:", padding=(10, 10)).pack(anchor=tk.W)
        
        # 添加文本显示区域（多行文本框，滚动条）
        text_frame = ttk.Frame(dialog)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        text_scroll = ttk.Scrollbar(text_frame)
        text_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        text_display = tk.Text(text_frame, height=15, wrap=tk.WORD, yscrollcommand=text_scroll.set)
        text_display.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        text_scroll.config(command=text_display.yview)
        
        # 插入当前提示词
        text_display.insert("1.0", self.current_prompt)
        text_display.config(state="disabled")  # 设置为只读
        
        # 添加按钮区域
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        # 使用此提示词按钮
        def use_and_close():
            self.use_current_prompt()
            dialog.destroy()
            
        ttk.Button(btn_frame, text="使用此提示词", command=use_and_close).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="关闭", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        
        # 居中显示对话框
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def use_current_prompt(self):
        """使用当前选择的提示词"""
        if self.current_prompt:
            self.input_text.delete("1.0", tk.END)
            self.input_text.insert("1.0", self.current_prompt)
            messagebox.showinfo("提示", "提示词已填入输入框，可以直接发送或修改后发送")

    def update_preview_image(self):
        try:
            preview_path = self.template_preview_paths[self.current_template]
            if preview_path not in self.preview_images:
                # 加载图片
                image = Image.open(preview_path)
                # 调整图片大小以适应预览区域
                # 计算缩放比例，保持宽高比
                preview_width = 320  # 进一步增加预览图的宽度（从300增加到320）
                ratio = preview_width / image.width
                new_size = (preview_width, int(image.height * ratio))
                image = image.resize(new_size, Image.Resampling.LANCZOS)
                # 转换为PhotoImage
                self.preview_images[preview_path] = ImageTk.PhotoImage(image)
            
            # 更新预览标签
            self.preview_label.configure(image=self.preview_images[preview_path])
            # 设置预览框的最小高度，确保图片显示完整
            current_image = self.preview_images[preview_path]
            preview_frame_height = current_image.height() + 30  # 增加一些额外空间给边框和标题
            self.preview_label.master.configure(height=preview_frame_height, width=current_image.width() + 20)  # 确保框宽度适应图片宽度
        except Exception as e:
            print(f"加载预览图片时出错: {str(e)}")
            self.preview_label.configure(image='')

    def on_template_change(self, event=None):
        self.current_template = self.template_var.get()
        # 更新预览图片
        self.update_preview_image()
        
        # 更新提示词列表
        self.prompt_combo['values'] = self.prompt_templates[self.current_template]
        if self.prompt_templates[self.current_template]:
            self.prompt_var.set(self.prompt_templates[self.current_template][0])
            self.current_prompt = self.prompt_templates[self.current_template][0]
        else:
            self.prompt_var.set("")
            self.current_prompt = ""
            
        # 重置对话 - 开始新的会话
        self.chat_assistant.reset_conversation()
        self.output_text.delete("1.0", tk.END)
        self.input_text.delete("1.0", tk.END)
        
        # 恢复到最初状态
        self.excel_data = None
        self.excel_file_path = ""
        self.excel_path_var.set("")
        self.excel_status_var.set("未导入Excel数据")
        self.image_path = ""
        
        # 显示欢迎信息
        self.output_text.insert(tk.END, "欢迎使用焊接工艺规程生成器！\n")
        self.output_text.insert(tk.END, "请先上传并解析Excel文件，然后选择提示词进行对话。\n\n")
        self.output_text.insert(tk.END, f"已切换到模板: {self.current_template}\n\n")
        self.output_text.see(tk.END)
        messagebox.showinfo("提示", f"已切换到{self.current_template}模板，系统已重置到初始状态")

    def setup_output_handling(self):
        def check_output():
            while True:
                try:
                    # Get output from queue without blocking
                    output = self.output_queue.get_nowait()
                    self.output_text.insert(tk.END, output)
                    self.output_text.see(tk.END)
                except:
                    break
            # Schedule next check
            self.root.after(100, check_output)

        # Start checking for output
        self.root.after(100, check_output)

    def send_message(self):
        # 检查是否已解析Excel数据
        if not self.excel_data:
            messagebox.showwarning("提示", "请先上传并解析Excel文件后再进行对话！")
            return
            
        user_input = self.input_text.get("1.0", tk.END).strip()
        if not user_input:
            return

        # Disable buttons while processing
        self.toggle_buttons(False)

        def process_message():
            try:
                # 显示用户输入
                self.root.after(0, lambda: self.output_text.insert(tk.END, f"\n用户: {user_input}\n\n"))
                self.root.after(0, lambda: self.output_text.insert(tk.END, "焊接工艺编写助手: "))
                self.root.after(0, lambda: self.output_text.see(tk.END))
                
                # Get response from DeepSeek client with Excel data
                response = self.chat_assistant.chat(
                    message=user_input,
                    excel_data=self.excel_data,
                    stream=True
                )

                # Clear input after successful send
                self.root.after(0, lambda: self.input_text.delete("1.0", tk.END))
                self.root.after(0, lambda: self.output_text.insert(tk.END, "\n\n"))
                self.root.after(0, lambda: self.output_text.see(tk.END))

            except Exception as e:
                error_msg = f"发送消息失败: {str(e)}"
                self.root.after(0, lambda: self.output_text.insert(tk.END, f"\n错误: {error_msg}\n\n"))
                messagebox.showerror("错误", error_msg)
            finally:
                # Re-enable buttons
                self.root.after(0, lambda: self.toggle_buttons(True))

        # Run in separate thread to prevent GUI freezing
        threading.Thread(target=process_message, daemon=True).start()

    def reset_conversation(self):
        # 重置对话 - 开始新的会话
        self.chat_assistant.reset_conversation()
        self.output_text.delete("1.0", tk.END)
        self.input_text.delete("1.0", tk.END)
        
        # 恢复到最初状态
        self.excel_data = None
        self.excel_file_path = ""
        self.excel_path_var.set("")
        self.excel_status_var.set("未导入Excel数据")
        self.image_path = ""
        
        # 重置模板选择到默认值
        self.current_template = "焊接工艺规程"
        self.template_var.set(self.current_template)
        self.update_preview_image()
        
        # 重置提示词到默认值
        self.prompt_combo['values'] = self.prompt_templates[self.current_template]
        if self.prompt_templates[self.current_template]:
            self.prompt_var.set(self.prompt_templates[self.current_template][0])
            self.current_prompt = self.prompt_templates[self.current_template][0]
        else:
            self.prompt_var.set("")
            self.current_prompt = ""
        
        # 显示欢迎信息
        self.output_text.insert(tk.END, "欢迎使用焊接工艺规程生成器！\n")
        self.output_text.insert(tk.END, "请先上传并解析Excel文件，然后选择提示词进行对话。\n\n")
        self.output_text.see(tk.END)
        
        messagebox.showinfo("重置", "系统已重置到初始状态，大模型对话会话已重新开始")

    def generate_document(self):
        last_response = self.chat_assistant.get_last_response()
        if not last_response:
            messagebox.showwarning("警告", "没有可用的回复内容。请先进行对话并获得有效回复。")
            return

        # Disable buttons while processing
        self.toggle_buttons(False)

        def process_generation():
            try:
                # Clean and normalize JSON text
                json_text = last_response.strip()
                
                # 确保目标目录存在
                save_dir = self.save_dir
                # 如果save_dir是字典类型，根据当前模板类型选择正确的保存路径
                if isinstance(self.save_dir, dict):
                    save_dir = self.save_dir[self.current_template]
                
                # 确保目录存在
                os.makedirs(save_dir, exist_ok=True)
                
                # 从Excel数据中获取工艺规程编号(WPS)作为文件名
                filename = "generated_doc"
                if self.excel_data and 'WPS' in self.excel_data:
                    wps_number = str(self.excel_data['WPS']).strip()
                    if wps_number and wps_number != 'nan':
                        # 处理文件名中的特殊字符，将斜杠替换为合法字符
                        filename = wps_number.replace('/', '').replace('\\', '')
                        # 特殊处理：G/TS 改为 GTS
                        filename = filename.replace('G/TS', 'GTS')
                    else:
                        import uuid
                        filename = f"generated_doc_{str(uuid.uuid4())[:8]}"
                else:
                    import uuid
                    filename = f"generated_doc_{str(uuid.uuid4())[:8]}"
                
                save_path = f"{save_dir}/{filename}.docx"

                # # 如果设置了图片路径，更新图片路径
                # if self.image_path:
                #     import json
                #     try:
                #         # 注释掉JSON解析部分，直接使用原始文本
                #         '''
                #         data = json.loads(json_text)
                #         # 更新所有图片路径
                #         for key, value in data.items():
                #             if isinstance(value, tuple) and len(value) == 2 and isinstance(value[1], str):
                #                 if value[1].endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                #                     # 保持文件名不变，只更新路径
                #                     filename = os.path.basename(value[1])
                #                     data[key] = (value[0], os.path.join(self.image_path, filename))
                #         json_text = json.dumps(data, ensure_ascii=False)
                #         '''
                #         # 直接使用原始文本
                #         pass
                #     except Exception as e:
                #         print(f"更新图片路径时出错: {str(e)}")

                # 获取选择的图片信息
                selected_image_info = self.get_selected_image_paths_for_generation()
                
                # Generate document
                self.chat_assistant.generate_document(
                    self.template_paths[self.current_template],
                    save_path,
                    json_text,
                    self.image_path if self.image_path else None,
                    selected_image_info  # 传递图片素材库选择的图片信息
                )

                # Show success message and ask for continuation
                self.root.after(0, lambda: self.show_generation_success_and_ask_continue(save_path))

            except Exception as e:
                import traceback
                traceback.print_exc()  # 打印完整的错误堆栈跟踪信息
                messagebox.showerror("Error", f"Failed to generate document: {str(e)}")
            finally:
                # Re-enable buttons
                self.root.after(0, lambda: self.toggle_buttons(True))

        # Run in separate thread to prevent GUI freezing
        threading.Thread(target=process_generation, daemon=True).start()

    def toggle_buttons(self, enabled: bool):
        state = 'normal' if enabled else 'disabled'
        self.reset_button.configure(state=state)
        self.send_button.configure(state=state)
        self.generate_button.configure(state=state)
    
    def show_generation_success_and_ask_continue(self, save_path: str):
        """
        显示生成成功消息并询问是否继续生成下一个文档
        
        Args:
            save_path: 已生成文档的保存路径
        """
        # 显示成功消息
        success_msg = f"文档生成成功：{save_path}\n\n是否继续生成下一个文档？"
        
        # 检查是否还有下一行数据
        if not self.excel_data or not self.excel_parser.has_next_row():
            # 没有更多数据
            messagebox.showinfo("生成完成", f"文档生成成功：{save_path}\n\n已处理完所有Excel数据，没有更多行可以处理。")
            return
        
        # 询问用户是否继续
        result = messagebox.askyesno("生成成功", success_msg)
        
        if result:  # 用户选择继续
            self.continue_next_document_generation()
        else:
            messagebox.showinfo("完成", "文档生成已完成。")
    
    def continue_next_document_generation(self):
        """
        继续生成下一个文档
        """
        try:
            # 1. 重置图片素材选择
            self.reset_image_materials()
            
            # 2. 重置大模型session
            self.reset_llm_session()
            
            # 3. 解析下一行Excel数据
            next_data = self.excel_parser.extract_next_row_data()
            
            if next_data:
                self.excel_data = next_data
                current_row = self.excel_parser.get_current_row_index() + 1
                total_rows = self.excel_parser.get_total_rows()
                
                # 更新状态显示
                self.excel_status_var.set(f"已解析Excel数据 (第{current_row}行/共{total_rows}行)")
                
                # 在输出区域显示新的数据信息
                self.output_text.insert(tk.END, f"\n=== 开始处理第{current_row}行数据 ===\n")
                self.output_text.insert(tk.END, f"WPS编号: {next_data.get('WPS', 'N/A')}\n")
                self.output_text.insert(tk.END, f"接头类型: {next_data.get('接头类型', 'N/A')}\n")
                self.output_text.insert(tk.END, "请选择图片素材并使用提示词进行对话生成。\n\n")
                self.output_text.see(tk.END)
                
                messagebox.showinfo("准备就绪", f"已加载第{current_row}行数据，请选择图片素材并开始新的对话。")
            else:
                messagebox.showinfo("完成", "没有更多Excel数据可以处理。")
                
        except Exception as e:
            messagebox.showerror("错误", f"处理下一行数据时发生错误: {str(e)}")
    
    def reset_image_materials(self):
        """
        重置图片素材选择
        """
        self.selected_images = {}
        self.image_path = ""
        self.output_text.insert(tk.END, "已重置图片素材选择。\n")
    
    def reset_llm_session(self):
        """
        重置大模型session
        """
        try:
            # 调用聊天助手的重置方法
            if hasattr(self.chat_assistant, 'reset_conversation'):
                self.chat_assistant.reset_conversation()
            elif hasattr(self.chat_assistant, 'reset'):
                self.chat_assistant.reset()
            
            self.output_text.insert(tk.END, "已重置大模型会话。\n")
        except Exception as e:
            self.output_text.insert(tk.END, f"重置大模型会话时出错: {str(e)}\n")
    
    def select_excel_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_file_path = file_path
            self.excel_path_var.set(file_path)
    
    def parse_excel_file(self):
        """解析Excel文件"""
        if not self.excel_file_path:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
        
        try:
            # 解析Excel文件
            self.excel_data = self.excel_parser.parse_file(self.excel_file_path)
            if self.excel_data:
                total_rows = self.excel_parser.get_total_rows()
                self.excel_status_var.set(f"已成功解析Excel数据 (第1行/共{total_rows}行)")
                messagebox.showinfo("成功", f"Excel文件解析成功！共找到{total_rows}行数据。")
            else:
                self.excel_status_var.set("Excel解析失败")
                messagebox.showerror("错误", "Excel文件解析失败，请检查文件格式")
        except Exception as e:
            self.excel_status_var.set("Excel解析出错")
            messagebox.showerror("错误", f"解析Excel文件时发生错误: {str(e)}")
    
    def preview_excel_data(self):
        """预览Excel数据"""
        if not self.excel_data:
            messagebox.showwarning("警告", "请先解析Excel文件")
            return
        
        # 创建预览窗口
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Excel数据预览")
        preview_window.geometry("800x600")
        preview_window.transient(self.root)
        preview_window.grab_set()
        
        # 创建主框架
        main_frame = ttk.Frame(preview_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建Notebook用于分页显示
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # 第一页：原始Excel数据
        excel_frame = ttk.Frame(notebook)
        notebook.add(excel_frame, text="原始Excel字段")
        
        # 添加说明标签
        ttk.Label(excel_frame, text="解析的Excel数据:", padding=(10, 10)).pack(anchor=tk.W)
        
        # 添加文本显示区域
        excel_text_frame = ttk.Frame(excel_frame)
        excel_text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        excel_scroll = ttk.Scrollbar(excel_text_frame)
        excel_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        excel_text_display = tk.Text(excel_text_frame, height=15, wrap=tk.WORD, yscrollcommand=excel_scroll.set)
        excel_text_display.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        excel_scroll.config(command=excel_text_display.yview)
        
        # 格式化并插入原始Excel数据
        formatted_data = self.excel_parser.format_data_for_prompt()
        excel_text_display.insert("1.0", formatted_data)
        excel_text_display.config(state="disabled")
        
        # 第二页：计算的焊接工艺参数
        params_frame = ttk.Frame(notebook)
        notebook.add(params_frame, text="计算的焊接工艺参数")
        
        # 添加说明标签
        ttk.Label(params_frame, text="计算出的焊接工艺参数:", padding=(10, 10)).pack(anchor=tk.W)
        
        # 添加文本显示区域
        params_text_frame = ttk.Frame(params_frame)
        params_text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        params_scroll = ttk.Scrollbar(params_text_frame)
        params_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        params_text_display = tk.Text(params_text_frame, height=15, wrap=tk.WORD, yscrollcommand=params_scroll.set)
        params_text_display.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        params_scroll.config(command=params_text_display.yview)
        
        # 计算并显示焊接工艺参数
        try:
            from wps_calculator import WPSCalculator
            calculator = WPSCalculator()
            calculated_params = calculator.calculate_welding_parameters(self.excel_data)
            formatted_params = calculator.format_parameters_for_display(calculated_params)
            params_text_display.insert("1.0", formatted_params)
        except Exception as e:
            error_msg = f"计算焊接工艺参数时发生错误:\n{str(e)}\n\n请检查Excel数据中的厚度信息是否正确。"
            params_text_display.insert("1.0", error_msg)
        
        params_text_display.config(state="disabled")
        
        # 添加关闭按钮
        ttk.Button(main_frame, text="关闭", command=preview_window.destroy).pack(pady=10)
        
        # 居中显示窗口
        preview_window.update_idletasks()
        width = preview_window.winfo_width()
        height = preview_window.winfo_height()
        x = (preview_window.winfo_screenwidth() // 2) - (width // 2)
        y = (preview_window.winfo_screenheight() // 2) - (height // 2)
        preview_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def clear_excel_data(self):
        """清除Excel数据"""
        self.excel_data = None
        self.excel_file_path = ""
        self.excel_path_var.set("")
        self.excel_status_var.set("未导入Excel数据")
        messagebox.showinfo("提示", "Excel数据已清除")

    def run(self):
        self.root.mainloop()
        # Restore original stdout when application closes
        sys.stdout = self.old_stdout

    def ensure_save_directories(self):
        """确保所有需要的保存目录都存在"""
        if isinstance(self.save_dir, dict):
            # 如果是字典类型，为每个模板类型创建对应的目录
            for template_type, directory in self.save_dir.items():
                os.makedirs(directory, exist_ok=True)
                print(f"已确保目录存在: {directory}")
        else:
            # 如果是字符串类型，创建单个目录
            os.makedirs(self.save_dir, exist_ok=True)
            print(f"已确保目录存在: {self.save_dir}")
    
    def open_image_library(self):
        """打开图片素材库窗口"""
        library_window = tk.Toplevel(self.root)
        library_window.title("图片素材库")
        library_window.geometry("1000x700")
        library_window.transient(self.root)
        library_window.grab_set()
        
        # 创建主框架
        main_frame = ttk.Frame(library_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 类别选择区域
        category_frame = ttk.LabelFrame(main_frame, text="选择接头类别", padding="10")
        category_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.category_var = tk.StringVar(value=self.image_categories[0])
        category_combo = ttk.Combobox(category_frame, textvariable=self.category_var, 
                                    values=self.image_categories, state="readonly", width=20)
        category_combo.pack(side=tk.LEFT, padx=(0, 10))
        
        refresh_btn = ttk.Button(category_frame, text="刷新图片", 
                               command=lambda: self.load_category_images(images_frame, self.category_var.get()))
        refresh_btn.pack(side=tk.LEFT)
        
        # 图片显示区域
        images_frame = ttk.LabelFrame(main_frame, text="图片选择", padding="10")
        images_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        cancel_btn = ttk.Button(button_frame, text="取消", command=library_window.destroy)
        cancel_btn.pack(side=tk.RIGHT)
        
        confirm_btn = ttk.Button(button_frame, text="确认选择", 
                               command=lambda: self.confirm_image_selection(library_window))
        confirm_btn.pack(side=tk.RIGHT, padx=(0, 10))
        
        # 绑定类别选择事件
        category_combo.bind('<<ComboboxSelected>>', 
                          lambda e: self.load_category_images(images_frame, self.category_var.get()))
        
        # 初始加载第一个类别的图片
        self.load_category_images(images_frame, self.category_var.get())
        
        # 为整个窗口绑定滚轮事件
        def on_window_mousewheel(event):
            # 查找当前活动的canvas
            for widget in images_frame.winfo_children():
                if isinstance(widget, tk.Canvas):
                    widget.yview_scroll(int(-1*(event.delta/120)), "units")
                    break
        
        library_window.bind("<MouseWheel>", on_window_mousewheel)
        
        # 居中显示窗口
        library_window.update_idletasks()
        width = library_window.winfo_width()
        height = library_window.winfo_height()
        x = (library_window.winfo_screenwidth() // 2) - (width // 2)
        y = (library_window.winfo_screenheight() // 2) - (height // 2)
        library_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def load_category_images(self, parent_frame, category):
        """加载指定类别的图片"""
        # 清空现有内容
        for widget in parent_frame.winfo_children():
            widget.destroy()
        
        # 创建滚动区域
        canvas = tk.Canvas(parent_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 布局Canvas和Scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 获取图片文件
        category_path = os.path.join(self.imgs_dir, category)
        if not os.path.exists(category_path):
            ttk.Label(scrollable_frame, text=f"未找到类别: {category}").pack(pady=20)
            return
        
        # 获取并分组图片
        image_pairs = self.get_image_pairs(category_path, category)
        
        if not image_pairs:
            ttk.Label(scrollable_frame, text="该类别下没有找到图片").pack(pady=20)
            return
        
        # 创建图片选择界面
        self.create_image_selection_ui(scrollable_frame, category, image_pairs)
        
        # 绑定鼠标滚轮事件到主界面Canvas
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def bind_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def unbind_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
        
        canvas.bind('<Enter>', bind_mousewheel)
        canvas.bind('<Leave>', unbind_mousewheel)
    
    def get_image_pairs(self, category_path, category):
        """获取图片配对信息"""
        files = os.listdir(category_path)
        image_pairs = {}
        
        for file in files:
            if file.endswith('.png'):
                # 解析文件名
                if '-焊接接头形式' in file:
                    # 提取序号
                    parts = file.split('-焊接接头形式')
                    if len(parts) == 2:
                        number = parts[1].replace('.png', '')
                        if number not in image_pairs:
                            image_pairs[number] = {}
                        image_pairs[number]['joint_form'] = os.path.join(category_path, file)
                elif '-焊接顺序' in file:
                    # 提取序号
                    parts = file.split('-焊接顺序')
                    if len(parts) == 2:
                        number = parts[1].replace('.png', '')
                        if number not in image_pairs:
                            image_pairs[number] = {}
                        image_pairs[number]['sequence'] = os.path.join(category_path, file)
        
        return image_pairs
    
    def create_image_selection_ui(self, parent, category, image_pairs):
        """创建图片选择界面"""
        # 创建单选按钮变量
        self.current_selection_var = tk.StringVar()
        
        # 检查是否已经有选择（全局检查，只有当前类别有选择时才设置）
        current_selection = None
        if category in self.selected_images and self.selected_images[category]:
            current_selection = list(self.selected_images[category].keys())[0]
            self.current_selection_var.set(current_selection)
        # 如果其他类别有选择，当前类别不设置任何选择
        elif any(self.selected_images.values()):
            self.current_selection_var.set("")  # 清空当前类别的选择
        
        for number, paths in sorted(image_pairs.items()):
            # 创建每组图片的框架
            group_frame = ttk.LabelFrame(parent, text=f"组合 {number}", padding="10")
            group_frame.pack(fill=tk.X, pady=5)
            
            # 单选按钮
            select_rb = ttk.Radiobutton(group_frame, text="选择此组合", 
                                      variable=self.current_selection_var,
                                      value=number,
                                      command=lambda n=number: 
                                      self.toggle_image_selection(category, n, True, image_pairs[n]))
            select_rb.pack(anchor=tk.W, pady=(0, 10))
            
            # 图片显示区域
            images_container = ttk.Frame(group_frame)
            images_container.pack(fill=tk.X)
            
            # 显示焊接接头形式图片
            if 'joint_form' in paths:
                joint_frame = ttk.Frame(images_container)
                joint_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
                
                ttk.Label(joint_frame, text="焊接接头形式", font=('Arial', 10, 'bold')).pack()
                self.display_thumbnail(joint_frame, paths['joint_form'])
            
            # 显示焊接顺序图片
            if 'sequence' in paths:
                sequence_frame = ttk.Frame(images_container)
                sequence_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                
                ttk.Label(sequence_frame, text="焊接顺序", font=('Arial', 10, 'bold')).pack()
                self.display_thumbnail(sequence_frame, paths['sequence'])
        
        # 存储选择变量以便后续使用
        # self.current_selection_var 已在方法开头定义
    
    def display_thumbnail(self, parent, image_path):
        """显示图片缩略图"""
        try:
            # 打开并调整图片大小
            image = Image.open(image_path)
            # 计算缩放比例，保持宽高比
            max_size = (200, 150)
            image.thumbnail(max_size, Image.Resampling.LANCZOS)
            
            # 转换为Tkinter可用的格式
            photo = ImageTk.PhotoImage(image)
            
            # 创建标签显示图片
            img_label = ttk.Label(parent, image=photo)
            img_label.image = photo  # 保持引用
            img_label.pack(pady=5)
            
            # 显示文件名
            filename = os.path.basename(image_path)
            ttk.Label(parent, text=filename, font=('Arial', 8)).pack()
            
        except Exception as e:
            ttk.Label(parent, text=f"无法加载图片: {str(e)}").pack()
    
    def toggle_image_selection(self, category, number, selected, paths):
        """切换图片选择状态（全局单选模式）"""
        # 清空所有类别的选择（确保全局只能选择一组）
        self.selected_images = {}
        # 设置新的选择
        if selected:
            self.selected_images[category] = {number: paths}
    
    def confirm_image_selection(self, window):
        """确认图片选择"""
        # 统计选择的图片数量
        total_selected = sum(len(selections) for selections in self.selected_images.values())
        
        if total_selected == 0:
            messagebox.showwarning("提示", "请至少选择一组图片")
            return
        
        # 更新主界面显示
        self.update_main_interface_with_selections()
        
        messagebox.showinfo("成功", f"已选择 {total_selected} 组图片")
        window.destroy()
    
    def update_main_interface_with_selections(self):
        """在主界面显示选择的图片信息"""
        # 查找并删除现有的图片选择内容
        content = self.output_text.get(1.0, tk.END)
        start_marker = "=== 已选择的图片素材 ==="
        end_marker = "=== 图片选择完成 ==="
        
        insert_position = tk.END  # 默认插入位置
        
        start_pos = content.find(start_marker)
        if start_pos != -1:
            # 找到开始位置，查找结束位置
            end_pos = content.find(end_marker, start_pos)
            if end_pos != -1:
                # 找到结束位置，计算在文本框中的行列位置
                start_line = content[:start_pos].count('\n') + 1
                end_line = content[:end_pos + len(end_marker)].count('\n') + 1
                
                # 记录插入位置（删除前的开始位置）
                insert_position = f"{start_line}.0"
                
                # 删除旧的图片选择内容
                self.output_text.delete(f"{start_line}.0", f"{end_line}.end")
        
        # 构建完整的图片选择内容
        image_content = "=== 已选择的图片素材 ===\n\n"
        image_paths_to_insert = []
        
        for category, selections in self.selected_images.items():
            if selections:
                image_content += f"{category}:\n"
                for number, paths in selections.items():
                    image_content += f"  组合 {number}:\n"
                    if 'joint_form' in paths:
                        image_content += f"    焊接接头形式: {os.path.basename(paths['joint_form'])}\n"
                        image_paths_to_insert.append((insert_position, paths['joint_form']))
                        image_content += "\n"  # 为图片预留空行
                    if 'sequence' in paths:
                        image_content += f"    焊接顺序: {os.path.basename(paths['sequence'])}\n"
                        image_paths_to_insert.append((insert_position, paths['sequence']))
                        image_content += "\n"  # 为图片预留空行
                image_content += "\n"
        
        image_content += "=== 图片选择完成 ===\n"
        
        # 在指定位置插入完整内容
        if insert_position == tk.END:
            self.output_text.insert(tk.END, "\n" + image_content)
            # 记录当前插入位置用于后续图片插入
            current_insert_pos = self.output_text.index(tk.END + "-1c")
        else:
            self.output_text.insert(insert_position, image_content)
            current_insert_pos = insert_position
        
        # 在正确位置插入图片缩略图
        self._insert_images_at_position(current_insert_pos, image_paths_to_insert)
        
        self.output_text.see(tk.END)
    
    def _insert_images_at_position(self, base_position, image_paths):
        """在指定位置插入图片缩略图"""
        # 由于图片插入会改变文本位置，我们需要从后往前插入
        # 或者使用更简单的方法：找到图片应该插入的文本位置
        content = self.output_text.get(1.0, tk.END)
        
        for _, image_path in image_paths:
            try:
                # 打开并调整图片大小
                image = Image.open(image_path)
                max_size = (150, 100)
                image.thumbnail(max_size, Image.Resampling.LANCZOS)
                
                # 转换为Tkinter可用的格式
                photo = ImageTk.PhotoImage(image)
                
                # 查找对应的图片文件名在文本中的位置
                filename = os.path.basename(image_path)
                current_content = self.output_text.get(1.0, tk.END)
                
                # 找到文件名后的换行位置
                filename_pos = current_content.find(filename)
                if filename_pos != -1:
                    # 计算行号
                    line_num = current_content[:filename_pos].count('\n') + 1
                    # 在该行末尾插入图片
                    insert_pos = f"{line_num}.end"
                    self.output_text.insert(insert_pos, "\n")
                    self.output_text.image_create(insert_pos + "+1c", image=photo)
                    
                    # 保持图片引用，防止被垃圾回收
                    if not hasattr(self, 'output_images'):
                        self.output_images = []
                    self.output_images.append(photo)
                    
            except Exception as e:
                print(f"插入图片失败: {str(e)}")
    
    def insert_image_to_output(self, image_path):
        """在输出区域插入图片缩略图"""
        try:
            # 打开并调整图片大小
            image = Image.open(image_path)
            # 计算缩放比例，保持宽高比
            max_size = (150, 100)
            image.thumbnail(max_size, Image.Resampling.LANCZOS)
            
            # 转换为Tkinter可用的格式
            photo = ImageTk.PhotoImage(image)
            
            # 在文本框中插入图片
            self.output_text.image_create(tk.END, image=photo)
            self.output_text.insert(tk.END, "\n")
            
            # 保持图片引用，防止被垃圾回收
            if not hasattr(self, 'output_images'):
                self.output_images = []
            self.output_images.append(photo)
            
        except Exception as e:
            self.output_text.insert(tk.END, f"    [无法显示图片: {str(e)}]\n")
    
    def get_selected_image_paths_for_generation(self):
        """获取用于文档生成的图片路径信息"""
        image_info = {}
        
        for category, selections in self.selected_images.items():
            for number, paths in selections.items():
                key = f"{category}-{number}"
                image_info[key] = {
                    'joint_form': paths.get('joint_form', ''),
                    'sequence': paths.get('sequence', ''),
                    'category': category,
                    'number': number
                }
        
        return image_info
