import pandas as pd
import requests
from requests.exceptions import RequestException
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import concurrent.futures
from urllib.parse import urlparse
import re
import docx
import os
import openpyxl
import urllib3

# 禁用 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 删除URL中的中文字符
def remove_chinese_characters(url):
    return re.sub(r'[\u4e00-\u9fff]+', '', url)

# 定义一个函数来检查 URL
def check_url(url, min_content_length=100):
    try:
        url = remove_chinese_characters(url)  # 处理URL中的中文字符
        response = requests.get(url, allow_redirects=True, timeout=5, verify=False)
        final_url = response.url
        content_type = response.headers.get('Content-Type', '').lower()
        content_length = len(response.content)

        if 'html' in content_type:
            soup = BeautifulSoup(response.content, 'html.parser')
            tags = [tag for tag in soup.find_all() if tag.name not in ['style', 'script', '[document]', 'head', 'title'] and not tag.string]
            login_form = soup.find('form', {'id': 'login'}) or soup.find('input', {'name': 'username'}) or soup.find('input', {'type': 'password'})

            if response.status_code == 200:
                if content_length < min_content_length:
                    return '可疑', response.status_code, content_type, content_length, f'内容过短 ({content_length} bytes)', final_url
                elif len(tags) < 5:
                    return '可疑', response.status_code, content_type, content_length, '页面可能功能点过少', final_url
                elif login_form:
                    return '可用', response.status_code, content_type, content_length, '页面包含登录功能', final_url
                else:
                    return '可用', response.status_code, content_type, content_length, '内容多半正常', final_url
            elif 400 <= response.status_code < 500:
                return '不可用', response.status_code, content_type, content_length, '客户端错误', final_url
            else:
                return '可疑', response.status_code, content_type, content_length, '可能的服务器错误', final_url
        else:
            return '可疑', response.status_code, content_type, content_length, '返回内容不是HTML', final_url
    except RequestException as e:
        error_type = type(e).__name__
        return '不可用', None, 'N/A', 0, f'{error_type}: {str(e)}', url

# 提取URL的域名
def get_domain(url):
    parsed_url = urlparse(url)
    return parsed_url.netloc

# 从 TXT 文件中提取 URL
def extract_urls_from_txt(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    urls = re.findall(r'\b(?:https?://|www\.)?(?:[\w-]+\.)+[\w-]+(?:/[\w./?%&=-]*)?', content)
    return urls

# 从 Word 文件中提取 URL
def extract_urls_from_docx(file_path):
    doc = docx.Document(file_path)
    urls = []
    for paragraph in doc.paragraphs:
        urls.extend(re.findall(r'\b(?:https?://|www\.)?(?:[\w-]+\.)+[\w-]+(?:/[\w./?%&=-]*)?', paragraph.text))
    return urls

# 拆分单元格中包含多个URL的情况
def split_cell_urls(cell_content):
    potential_delimiters = r'[;, \n]+'
    urls = re.split(potential_delimiters, cell_content)
    return [url.strip() for url in urls if url.strip()]

# 从文件中读取URL，并处理单元格中包含多个URL的情况
def read_urls_from_file(file_path):
    all_urls = []
    if file_path.endswith('.xlsx') or file_path.endswith('.csv'):
        df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
        if 'URL' in df.columns:
            for cell in df['URL'].dropna():
                all_urls.extend(split_cell_urls(str(cell)))
    elif file_path.endswith('.txt'):
        all_urls = extract_urls_from_txt(file_path)
    elif file_path.endswith('.docx'):
        all_urls = extract_urls_from_docx(file_path)
    else:
        raise ValueError("不支持的文件格式")
    return all_urls

# 定义一个函数来分类 URL 并保存结果
def categorize_urls(input_file, output_dir, log_text_widget, root, max_threads):
    try:
        urls = read_urls_from_file(input_file)

        # 提取域名并去重
        unique_urls = list(set(urls))
        unique_domains = set()
        filtered_urls = []

        for url in unique_urls:
            domain = get_domain(url)
            if domain not in unique_domains:
                unique_domains.add(domain)
                filtered_urls.append(url)

        # 初始化一个字典来存储结果
        results = {'可用': [], '不可用': [], '可疑': []}

        url_count = tk.IntVar(value=0)

        def update_log_text(message):
            log_text_widget.config(state='normal')
            log_text_widget.insert(tk.END, message)
            log_text_widget.see(tk.END)
            log_text_widget.config(state='disabled')
            log_text_widget.update()

        def process_url(url):
            try:
                if not url.startswith(('http://', 'https://')):
                    url = 'http://' + url

                status, status_code, content_type, content_length, note, final_url = check_url(url)

                if status in results:
                    results[status].append((final_url, status_code, content_type, content_length, note))
                else:
                    update_log_text(f"未知状态: {status}\n")
                    return

                update_log_text(
                    f'检查结果: {status}, 最终URL: {final_url}, 状态码: {status_code}, 内容类型: {content_type}, 页面长度: {content_length}, 备注: {note}\n')

                url_count.set(url_count.get() + 1)
                update_log_text(f'已检查 URL 数量: {url_count.get()}\n')

            except Exception as e:
                update_log_text(f'检查URL时出错: {url}, 错误信息: {str(e)}\n')

        # 初次检测
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_threads) as executor:
            futures = []
            for url in filtered_urls:
                futures.append(executor.submit(process_url, url))

            concurrent.futures.wait(futures)

        # 保存初次检测结果
        for status, urls in results.items():
            result_df = pd.DataFrame(urls, columns=['URL', '状态码', '内容类型', '页面长度', '备注'])
            output_file = f'{output_dir}/{status}_url.xlsx'
            result_df.to_excel(output_file, index=False)
            update_log_text(f'已保存 {status} 的结果到 {output_file}\n')

        # # 复测不可用的URL
        # if results['不可用']:
        #     update_log_text("\n开始复测不可用的 URL...\n")
        #     recheck_results = {'可用': [], '不可用': [], '可疑': []}
        #
        #     def recheck_url(url_data):
        #         url, _, _, _, _, _ = url_data
        #         status, status_code, content_type, content_length, note, final_url = check_url(url)
        #         recheck_results[status].append((final_url, status_code, content_type, content_length, note))
        #         update_log_text(
        #             f'复测结果: {status}, 最终URL: {final_url}, 状态码: {status_code}, 内容类型: {content_type}, 页面长度: {content_length}, 备注: {note}\n')
        #
        #     with concurrent.futures.ThreadPoolExecutor(max_workers=max_threads) as executor:
        #         futures = []
        #         for url_data in results['不可用']:
        #             futures.append(executor.submit(recheck_url, url_data))
        #
        #         concurrent.futures.wait(futures)
        #
        #     # 保存复测结果
        #     recheck_file = f'{output_dir}/复测结果.xlsx'
        #     with pd.ExcelWriter(recheck_file) as writer:
        #         for status, urls in recheck_results.items():
        #             recheck_df = pd.DataFrame(urls, columns=['URL', '状态码', '内容类型', '页面长度', '备注'])
        #             recheck_df.to_excel(writer, sheet_name=status, index=False)
        #     update_log_text(f'\n已保存复测结果到 {recheck_file}\n')

        # 显示结果总数
        final_message = (
            f"URL 分类完成:\n"
            f"可用的 URL 总数: {len(results['可用']) + len(recheck_results['可用'])}\n"
            f"不可用的 URL 总数: {len(results['不可用']) + len(recheck_results['不可用'])}\n"
            f"可疑的 URL 总数: {len(results['可疑']) + len(recheck_results['可疑'])}\n"
        )
        update_log_text(f'\n{final_message}')
        root.after(0, messagebox.showinfo, "完成", final_message)

    except Exception as e:
        root.after(0, messagebox.showerror, "错误", f"出现错误: {str(e)}")

# 选择输入文件
def select_input_file():
    file_path = filedialog.askopenfilename(filetypes=[("All supported files", "*.xlsx;*.csv;*.txt;*.docx"),
                                                      ("Excel files", "*.xlsx"),
                                                      ("CSV files", "*.csv"),
                                                      ("Text files", "*.txt"),
                                                      ("Word files", "*.docx")])
    input_file_var.set(file_path)

# 选择输出文件夹
def select_output_directory():
    directory = filedialog.askdirectory()
    output_dir_var.set(directory)

# 创建 GUI
root = tk.Tk()
root.title("URL 分类器")
root.geometry("900x500")
root.resizable(True, True)

# 设置整体布局
frame = tk.Frame(root, padx=10, pady=10)
frame.pack(fill=tk.BOTH, expand=True)

# 输入文件选择
input_frame = tk.LabelFrame(frame, text="输入文件", padx=10, pady=10)
input_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

input_file_var = tk.StringVar()
tk.Entry(input_frame, textvariable=input_file_var, width=50).grid(row=0, column=0, padx=5, pady=5)
tk.Button(input_frame, text="浏览", command=select_input_file).grid(row=0, column=1, padx=5, pady=5)

# 输出文件夹选择
output_frame = tk.LabelFrame(frame, text="输出文件夹", padx=10, pady=10)
output_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

output_dir_var = tk.StringVar()
tk.Entry(output_frame, textvariable=output_dir_var, width=50).grid(row=0, column=0, padx=5, pady=5)
tk.Button(output_frame, text="浏览", command=select_output_directory).grid(row=0, column=1, padx=5, pady=5)

# 并发线程数选择
thread_frame = tk.LabelFrame(frame, text="线程设置", padx=10, pady=10)
thread_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)

tk.Label(thread_frame, text="线程数量:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
thread_count_var = tk.IntVar(value=5)  # 默认线程数
tk.Spinbox(thread_frame, from_=1, to=100, textvariable=thread_count_var, width=5).grid(row=0, column=1, padx=5, pady=5, sticky='w')

# 日志输出区域，放在右边并允许调整大小
log_frame = tk.LabelFrame(frame, text="日志输出", padx=10, pady=10)
log_frame.grid(row=0, column=1, rowspan=4, sticky="nsew", padx=5, pady=5)

log_text = ScrolledText(log_frame, width=50, height=25, state='disabled')
log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

# 开始处理按钮
button_frame = tk.Frame(frame)
button_frame.grid(row=3, column=0, pady=10)

tk.Button(button_frame, text="开始处理", command=lambda: threading.Thread(target=categorize_urls, args=(
    input_file_var.get(), output_dir_var.get(), log_text, root, thread_count_var.get())).start()).pack()

# 设置行列权重，以便窗口调整时保持布局
frame.grid_rowconfigure(3, weight=1)
frame.grid_columnconfigure(0, weight=1)
frame.grid_columnconfigure(1, weight=2)  # 调整右侧日志窗口的布局权重

# 运行 GUI 主循环
root.mainloop()
