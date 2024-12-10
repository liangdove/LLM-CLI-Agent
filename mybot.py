import argparse
from ollama import chat
import ollama
import pyfiglet
import keyboard  # 用于监听键盘事件
import os  # 用于文件路径验证
from docx import Document  # 引入 python-docx
import PyPDF2  # 用于处理 PDF 文件
import pandas as pd  # 用于处理 Excel 和 CSV 文件
import subprocess

# 全局变量：历史对话存储
history = []

# 默认模型名称
MODEL_NAME = 'qwen2.5:3B'

# 可用模型列表
MODELS = {
    "1": "qwen2.5:3B",
    "2": "lamma3.2:70B"
}

# ANSI 转义序列，用于文本高亮和颜色
HIGHLIGHT_USER = "\033[1;32m"  # 绿色，高亮 User
HIGHLIGHT_BOT = "\033[1;34m"   # 蓝色，高亮 Bot
HIGHLIGHT_WARNING = "\033[1;31m"   # 红色，高亮 Bot
RESET = "\033[0m"              # 重置颜色和样式

def start_conversation():
    """启动对话框"""
    global history, MODEL_NAME
    print(f"欢迎使用 MYBOT ！")
    print(f"""
输入{HIGHLIGHT_BOT}'exit' / 'close'{RESET} 退出对话
输入{HIGHLIGHT_BOT}'S' {RESET}             切换模型
输入{HIGHLIGHT_BOT}'p' {RESET}             停止对话
输入{HIGHLIGHT_BOT}'file' {RESET}          文件操作
输入{HIGHLIGHT_BOT}'summarize' {RESET}     总结文档
          """)
    while True:
        user_input = input(f"{HIGHLIGHT_USER}User：{RESET}")
        if user_input.lower() in ['exit', 'close']:
            print(f"{HIGHLIGHT_WARNING}对话已关闭{RESET}")
            break
        if user_input == "S":
            switch_model()
            continue
        if user_input == "summarize":
            summarize_document()
            continue
        if user_input == "file":
            process_file_command()
            continue
        try:
            # 调用 Ollama API，启用流式响应
            print(f"{HIGHLIGHT_BOT}Bot", end='', flush=True)
            print(f"（{MODEL_NAME}）{RESET}：", end='', flush=True)
            stream = chat(
                model=MODEL_NAME,
                messages=[
                    {"role": "system", "content": "你是一个友好耐心的助手，专门服务北京邮电大学的我。"},
                    *history,
                    {"role": "user", "content": user_input},
                ],
                stream=True
            )
            reply_chunks = []
            for chunk in stream:
                # 监听 p 键，若按下则停止当前输出并开启下一轮对话
                if keyboard.is_pressed('p'):
                    print(f"\n{HIGHLIGHT_WARNING}已停止当前对话{RESET}")
                    break  # 跳出当前的输出循环
                content = chunk['message']['content']
                print(content, end='', flush=True)  # 仅显示Bot的回应内容
                reply_chunks.append(content)
            else:
                print()  # 打印换行
                # 将完整回复组合
                reply = ''.join(reply_chunks)

                # 更新历史对话
                history.append({"role": "user", "content": user_input})
                history.append({"role": "assistant", "content": reply})
        except Exception as e:
            print(f"{HIGHLIGHT_WARNING}发生错误：{e}{RESET}")


def extract_pdf_text(file_path):
    """从 PDF 文件中提取文本"""
    content = []
    try:
        with open(file_path, "rb") as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            # 遍历每一页提取文本
            for page in reader.pages:
                content.append(page.extract_text())
    except Exception as e:
        print(f"PDF 文件处理出错：{e}")
    return "\n".join(content)


def summarize_document():
    """总结文档内容"""
    global history, MODEL_NAME
    file_path = input(f"{HIGHLIGHT_USER}请输入文件路径：{RESET}")

    try:  
        # 检查文件是否存在
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"{HIGHLIGHT_WARNING}文件未找到，请确认路径是否正确{RESET}")

        # 根据文件类型读取内容
        if file_path.endswith(".docx"):
            # 处理 Word 文档
            document = Document(file_path)
            content = "\n".join([para.text for para in document.paragraphs])
        elif file_path.endswith(".pdf"):
            # 处理 PDF 文档
            content = extract_pdf_text(file_path)
        elif file_path.endswith(".xlsx"):
            # 处理 Excel 文件
            df = pd.read_excel(file_path, engine="openpyxl")
            content = df.to_string(index=False)  # 将 DataFrame 转为字符串
        elif file_path.endswith(".csv"):
            # 处理 CSV 文件
            df = pd.read_csv(file_path)
            content = df.to_string(index=False)  # 将 DataFrame 转为字符串
        else:
            # 处理普通文本文件
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()

        print(f"正在生成文档总结，请稍候...")

        # 调用 Ollama API 流式生成总结
        stream = chat(
            model=MODEL_NAME,
            messages=[
                {"role": "system", "content": "你现在是一个帮助生成文档总结的助手；对于文字文档，请总结文档的文字内容。对于数据文档，请总结相应的数据特征。"},
                {"role": "user", "content": f"请总结以下文档内容：\n\n{content}"}
            ],
            stream=True  # 启用流式输出
        )

        # 实时显示生成的内容
        print(f"{HIGHLIGHT_BOT}文档总结{RESET}：", end='', flush=True)
        summary_chunks = []
        for chunk in stream:
            content = chunk['message']['content']
            print(content, end='', flush=True)
            summary_chunks.append(content)
        
        # 合并总结内容
        summary = ''.join(summary_chunks)

        print(f"{HIGHLIGHT_BOT}\n总结生成完成！{RESET}")  # 输出换行以美化控制台显示
        
        # 保存总结到历史对话中
        history.append({"role": "user", "content": f"请总结以下文档内容：\n\n{content}"})
        history.append({"role": "assistant", "content": summary})
        
    except FileNotFoundError as e:
        print(f"{HIGHLIGHT_WARNING}{e}{RESET}")
    except Exception as e:
        print(f"{HIGHLIGHT_WARNING}发生错误：{e}{RESET}")


import subprocess

def process_file_command():
    """
    通过自然语言解析用户意图，并执行文件操作命令
    """
    global history, MODEL_NAME
    print(f"输入任何你想对文件做的操作\n比如你可以说：“{HIGHLIGHT_BOT}将A.txt的文件内容复制到B.txt。{RESET}”")
    print(f"输入{HIGHLIGHT_BOT}'退出'{RESET} 退出文件操作")

    while True:
        user_input = input("请输入文件操作指令：")
        
        # 如果用户输入 '退出'，退出循环
        if user_input.lower() == '退出':
            print(f"{HIGHLIGHT_WARNING}退出文件操作模式{RESET}")
            break
        
        try:
            # 调用 Ollama 模型解析自然语言为具体命令
            response = chat(
                model=MODEL_NAME,
                messages=[
                    {"role": "system", "content": """
你是一个文件管理助手，请将用户的自然语言输入映射为Windows的CMD命令并返回对应的命令行，
你极有可能用到'dir','mkdir','type','more','echo','del','copy',rename等命令，
你只需要给出命令即可，不要多余的回答，不要多余的回答！
                     """},
                    {"role": "user", "content": f"将以下自然语言转换为CMD的命令：{user_input}"}
                ]
            )
            
            # 从模型响应中提取映射的命令
            mapped_command = response["message"]["content"]
            print(f"{HIGHLIGHT_BOT}映射的命令：\n{RESET}{mapped_command}")

            # 将命令拆分为多条指令
            # commands_list = mapped_command.split(';')  # 如果命令是用分号分隔的
            commands_list = mapped_command.split('\n')  # 如果是换行分隔，可以改为此行
            
            # 使用 subprocess 执行每条指令
            for cmd in commands_list:
                cmd = cmd.strip()  # 去掉多余空格
                if not cmd:  # 跳过空命令
                    continue
                print(f"{HIGHLIGHT_BOT}执行命令：{cmd}{RESET}")
                result = subprocess.run(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                
                # 获取命令的输出
                command_output = result.stdout
                command_error = result.stderr

                if command_output.strip():
                    print(f"{HIGHLIGHT_BOT}命令执行结果{RESET}：\n{command_output}")
                elif command_error.strip():
                    print(f"{HIGHLIGHT_WARNING}命令执行时发生错误：{RESET}\n{command_error}")
                else:
                    print(f"{HIGHLIGHT_BOT}命令未产生输出，请确认该命令是否有效{RESET}")

        except Exception as e:
            print(f"命令执行失败：{e}")


def switch_model():
    """切换模型"""
    global MODEL_NAME
    print("可用模型列表：")
    for key, model_name in MODELS.items():
        print(f"{key}. {model_name}")
    selected = input("请选择模型编号：")
    if selected in MODELS:
        MODEL_NAME = MODELS[selected]
        print(f"{HIGHLIGHT_WARNING}模型已切换为：{MODEL_NAME}{RESET}")
    else:
        print(f"{HIGHLIGHT_WARNING}输入无效，继续使用当前模型{RESET}")


def show_help():
    """显示帮助信息"""
    ascii_art = pyfiglet.figlet_format("BUPT OS MYBOT")
    print(f"{HIGHLIGHT_BOT}----------------------------------------------------------------------\n")
    print(f"{ascii_art}")
    print(f"----------------------------------------------------------------------{RESET}")
    help_text = """
MYBOT  一个基于 Ollama 的大语言模型交互工具\n
使用方法：
全局：
  mybot help        帮助信息
  mybot start       开启对话
  mybot info        模型信息
  mybot summarize   总结文档
  mybot file        文件操作
对话:
  S                 切换模型
  p                 停止对话
  summarize         总结文档
  file              文件操作
  exit/close        关闭对话
    """
    print(help_text)


def show_info():
    """显示模型信息"""
    global MODEL_NAME
    print("可用模型列表：\n")
    key_width = max(len(key) for key in MODELS.keys())  # 编号的最大宽度
    name_width = max(len(name) for name in MODELS.values())  # 模型名称的最大宽度
    
    # 打印模型列表
    header = f"{'编号'.ljust(key_width)}  {'模型名称'.ljust(name_width)}"
    print(header)
    print("-" * len(header))
    for key, model_name in MODELS.items():
        print(f"{key.ljust(key_width)}  {model_name.ljust(name_width)}")
    
    selected = input("请选择模型编号：")
    if selected in MODELS:
        MODEL_NAME = MODELS[selected]
        try:
            info = ollama.show(MODEL_NAME)
            print("\n模型信息：")
            print(info)
        except Exception as e:
            print(f"{HIGHLIGHT_WARNING}无法获取模型信息，错误：{e}{RESET}")
    else:
        print(f"{HIGHLIGHT_WARNING}没有此模型信息{RESET}")


# 主函数
def main():
    """主函数：解析命令行参数并调用相应功能"""
    parser = argparse.ArgumentParser(description="mybot - Ollama LLM 终端工具")
    parser.add_argument("command", choices=["help", "start", "info", "summarize", "file"], help="命令")
    args = parser.parse_args()

    if args.command == "help":
        show_help()
    elif args.command == "start":
        start_conversation()
    elif args.command == "info":
        show_info()
    elif args.command == "summarize":
        summarize_document()
    elif args.command == "file":
        process_file_command()
    else:
        print(f"{HIGHLIGHT_WARNING}无效命令，请使用 'mybot help' 查看帮助信息{RESET}")

if __name__ == "__main__":
    main()
