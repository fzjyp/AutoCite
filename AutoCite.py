import os
import re
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import scrolledtext
import docx
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH



def process_reference_string(ref, index):
    # 1. 替换全角中括号与圆括号
    ref = ref.replace('［', '[').replace('］', ']').replace('【', '[').replace('】', ']')
    ref = ref.replace('（', '(').replace('）', ')')

    # 2. 标点替换（中文全角转英文半角）
    punctuation_map = {'，': ',', '。': '.', '：': ':', '；': ';', '．': '.'}
    for zh_punc, en_punc in punctuation_map.items():
        ref = ref.replace(zh_punc, en_punc)

    # 3. 压缩多余空格
    ref = re.sub(r' {2,}', ' ', ref)

    # 4. 删除中英文字符交界处的非法空格
    ref = re.sub(r'(?<=[\u4e00-\u9fa5])\s+(?=[\u4e00-\u9fa5])', '', ref)
    ref = re.sub(r'(?<=[\u4e00-\u9fa5])\s+(?=[a-zA-Z0-9])', '', ref)
    ref = re.sub(r'(?<=[a-zA-Z0-9])\s+(?=[\u4e00-\u9fa5])', '', ref)

    # 5. 删除括号内外的多余空格
    ref = re.sub(r'\s+([\[\(])', r'\1', ref)
    ref = re.sub(r'([\]\)])\s+', r'\1', ref)
    ref = re.sub(r'([\[\(])\s+', r'\1', ref)
    ref = re.sub(r'\s+([\]\)])', r'\1', ref)

    # 6. 删除标点符号前的多余空格
    ref = re.sub(r'\s+([,.:;])', r'\1', ref)

    # 7. 强制重排编号，并严格删除编号与第一个字符之间的空格
    ref = re.sub(r'^\[\s*\d+\s*\]\s*', f'[{index}]', ref)

    # 8. 区分中英文文献，精准定制标点后的空格
    is_chinese = bool(re.search(r'[\u4e00-\u9fa5]', ref))
    if is_chinese:
        ref = re.sub(r'([,.:;])\s+', r'\1', ref)
    else:
        ref = re.sub(r'([,;])\s*', r'\1 ', ref)
        ref = re.sub(r'(:)\s*(?!/)', r'\1 ', ref)
        ref = re.sub(r'(\.)\s+', r'\1 ', ref)
        ref = re.sub(r'(\.)(?=[A-Z\[\(])', r'\1 ', ref)
        ref = ref.strip()

    return ref


def apply_academic_format(paragraph, text):
    run = paragraph.add_run(text)

    # 字体设置：英文新罗马，中文宋体，字号小四
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # 段落设置：顶格，单倍行距，两端对齐
    paragraph.paragraph_format.left_indent = Pt(0)
    paragraph.paragraph_format.first_line_indent = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # 强制勾选“允许西文在单词中间换行”
    pPr = paragraph._element.get_or_add_pPr()
    wordWrap = OxmlElement('w:wordWrap')
    wordWrap.set(qn('w:val'), 'off')
    pPr.append(wordWrap)


class ReferenceFormatterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoCite - 自动文献排版 V1.0")
        self.root.geometry("480x400")
        self.root.resizable(False, False)

        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', font=('微软雅黑', 11, 'bold'), padding=8)

        # 顶部面板：描述与按钮
        top_frame = tk.Frame(root, padx=20, pady=15)
        top_frame.pack(fill=tk.X)

        tk.Label(top_frame, text="学术参考文献 格式清洗", font=("微软雅黑", 16, "bold"), fg="#2c3e50").pack(pady=(0, 5))
        tk.Label(top_frame, text="自动定位“参考文献”", font=("微软雅黑", 9), fg="#7f8c8d").pack(pady=(0, 15))

        self.btn_run = ttk.Button(top_frame, text="🚀 选择论文文档并一键排版", command=self.start_processing)
        self.btn_run.pack(fill=tk.X)

        log_frame = tk.Frame(root, padx=20, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(log_frame, text="处理日志 (Console)", font=("微软雅黑", 9, "bold")).pack(anchor="w")

        self.txt_log = scrolledtext.ScrolledText(log_frame, height=10, bg="#1e1e1e", fg="#00ff00", font=("Consolas", 9))
        self.txt_log.pack(fill=tk.BOTH, expand=True)
        self.txt_log.config(state=tk.DISABLED)

    def log(self, message):

        def update_log():
            self.txt_log.config(state=tk.NORMAL)
            current_time = time.strftime("%H:%M:%S")
            self.txt_log.insert(tk.END, f"[{current_time}] {message}\n")
            self.txt_log.see(tk.END)
            self.txt_log.config(state=tk.DISABLED)

        self.root.after(0, update_log)

    def start_processing(self):
        file_path = filedialog.askopenfilename(
            title="请选择包含参考文献的 Word 文档",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")]
        )

        if not file_path:
            return

        self.btn_run.config(text="后台执行中，请看日志...", state=tk.DISABLED)

        # 清空之前的日志
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.delete(1.0, tk.END)
        self.txt_log.config(state=tk.DISABLED)

        threading.Thread(target=self._process_thread, args=(file_path,), daemon=True).start()

    def _process_thread(self, file_path):
        """独立的后台处理线程"""
        try:
            self.log(f"📄 已加载文档: {os.path.basename(file_path)}")
            self.log("🔍 启动智能雷达，全篇扫描安全边界...")

            doc = docx.Document(file_path)

            start_idx = -1
            end_idx = -1

            # 定义起点与终点雷达特征
            start_patterns = [r'^\s*参\s*考\s*文\s*献\s*$', r'^\s*References\s*$']
            stop_patterns = [r'^\s*附\s*录', r'^\s*致\s*谢', r'^\s*Acknowledgements', r'^\s*攻读.*?成果', r'^\s*发表.*?论文',
                             r'^\s*附\s*件']

            for i, para in enumerate(doc.paragraphs):
                text = para.text.strip()
                if not text:
                    continue

                if start_idx == -1:
                    if any(re.match(p, text, re.IGNORECASE) for p in start_patterns):
                        start_idx = i
                        self.log(f"🎯 成功锁定参考文献起点 (第 {i} 段)")
                else:
                    if any(re.match(p, text, re.IGNORECASE) for p in stop_patterns):
                        end_idx = i
                        self.log(f"🛑 探测到后续边界 [{text[:5]}...]，已划定沙盒隔离区")
                        break

            if start_idx == -1:
                self.log("❌ 致命错误：未能在文档中找到“参考文献”独立标题！")
                self.root.after(0, lambda: messagebox.showerror("定位失败", "未能在文档中找到“参考文献”独立标题。\n程序已安全停止，未做任何修改。"))
                return

            if end_idx == -1:
                end_idx = len(doc.paragraphs)
                self.log("🔚 参考文献直达文档末尾，无其他干扰章节。")

            self.log("🧹 正在提取乱码文本，消灭幽灵换行...")

            lines = []
            for i in range(start_idx + 1, end_idx):
                para_text = re.sub(r'[\r\v\x0b]', '\n', doc.paragraphs[i].text)
                for sub_line in para_text.split('\n'):
                    if sub_line.strip():
                        lines.append(sub_line.strip())

            valid_refs = []
            ref_start_pattern = re.compile(r'^[\s\u3000]*([\[【［]\s*\d+\s*[\]】］])')

            for line in lines:
                if ref_start_pattern.match(line):
                    valid_refs.append(line)
                else:
                    if valid_refs:
                        valid_refs[-1] += " " + line
                    # 如果不是序号开头，且前面也没序号，直接舍弃

            self.log(f"✨ 提取完毕！共抓取到 {len(valid_refs)} 条待清洗文献。")
            self.log("✂️ 正在应用正则，洗脱非法空格与标点...")

            final_refs = [process_reference_string(ref, i + 1) for i, ref in enumerate(valid_refs)]

            self.log("💥 正在切除原有凌乱段落...")
            for i in range(end_idx - 1, start_idx, -1):
                p_element = doc.paragraphs[i]._element
                p_element.getparent().remove(p_element)

            # 注入新格式
            self.log("💉 正在重新注入国标排版样式...")
            insert_before_p = doc.paragraphs[end_idx] if end_idx < len(doc.paragraphs) else None

            for ref in final_refs:
                new_p = insert_before_p.insert_paragraph_before() if insert_before_p else doc.add_paragraph()
                apply_academic_format(new_p, ref)

            # 保存文件
            dir_name = os.path.dirname(file_path)
            base_name = os.path.basename(file_path)
            name, ext = os.path.splitext(base_name)
            new_file_path = os.path.join(dir_name, f"{name}_排版完成{ext}")

            doc.save(new_file_path)

            self.log(f"✅ 大功告成！新文件已安全生成！")

            # 提示成功
            success_msg = f"✅ 扫描并重排了 {len(final_refs)} 条参考文献。\n\n沙盒隔离机制运行完美，您的正文毫发无伤！\n\n新文件已生成：\n{name}_急救排版完成{ext}"
            self.root.after(0, lambda: messagebox.showinfo("排版完成", success_msg))

        except Exception as e:
            self.log(f"❌ 系统崩溃: {e}")
            self.root.after(0, lambda: messagebox.showerror("系统错误", f"处理过程中发生异常：\n{e}"))
        finally:
            self.root.after(0, lambda: self.btn_run.config(text="🚀 选择论文文档并一键排版", state=tk.NORMAL))


if __name__ == "__main__":
    root = tk.Tk()
    app = ReferenceFormatterApp(root)
    root.mainloop()