import pyperclip
import pypandoc
import os
import datetime


def save_clipboard_md_to_word():
    """
    从剪贴板读取Markdown内容，并将其转换为一个带时间戳的Word文档。
    该函数能够精准处理格式、LaTeX公式和样式。
    它会自动检测是否存在名为 'my_template.docx' 的样式模板，并根据情况决定是否使用。
    """
    try:
        # --- 配置 ---
        # 模板文件名 (可自定义)
        reference_doc_file = "my_template.docx"

        # 1. 从剪贴板读取内容
        markdown_content = pyperclip.paste()
        if not markdown_content:
            print("剪贴板为空，没有内容可以转换。")
            print("Clipboard is empty. Nothing to convert.")
            return
        print("成功从剪贴板读取到内容。 (Successfully read content from clipboard.)")

        # 2. 定义输出文件名
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"output_{timestamp}.docx"

        # 3. 构建 Pandoc 转换参数
        extra_args = ['--standalone']

        # 检查模板文件是否存在
        if os.path.exists(reference_doc_file):
            print(f"检测到样式模板 '{reference_doc_file}'，将使用它进行转换。")
            print(f"Found style template '{reference_doc_file}', applying it for conversion.")
            extra_args.append(f'--reference-doc={reference_doc_file}')
        else:
            print(f"未检测到样式模板 '{reference_doc_file}'。")
            print("将使用 Pandoc 默认样式进行转换。")
            print(f"Style template '{reference_doc_file}' not found. Using Pandoc's default styles.")

        # 4. 使用 pypandoc 进行转换
        pypandoc.convert_text(
            markdown_content,
            'docx',
            format='md',
            outputfile=output_filename,
            extra_args=extra_args
        )

        print("\n" + "=" * 50)
        print(f"✅ 文件转换成功！已保存为: {os.path.abspath(output_filename)}")
        print(f"✅ File converted successfully! Saved as: {os.path.abspath(output_filename)}")
        print("=" * 50)


    except FileNotFoundError:
        print("\n[错误] Pandoc 未找到。 (Error: Pandoc not found.)")
        print("请确保您已经安装了 Pandoc 并且它在系统的 PATH 环境变量中。")
        print("Please ensure Pandoc is installed and accessible in your system's PATH.")

    except Exception as e:
        print(f"\n处理过程中发生未知错误: {e}")
        print(f"An unexpected error occurred: {e}")


if __name__ == "__main__":
    save_clipboard_md_to_word()
