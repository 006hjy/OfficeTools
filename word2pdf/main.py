import os
import sys
import comtypes.client

def get_unique_pdf_path(directory, filename_no_ext):
    """
    生成唯一的PDF文件路径。
    如果文件已存在，则在文件名后添加下划线，直到文件名唯一。
    """
    base_name = filename_no_ext
    pdf_name = f"{base_name}.pdf"
    pdf_path = os.path.join(directory, pdf_name)
    
    while os.path.exists(pdf_path):
        base_name += "_"
        pdf_name = f"{base_name}.pdf"
        pdf_path = os.path.join(directory, pdf_name)
        
    return pdf_path

def main():
    # 获取脚本所在的当前目录 (兼容打包后的 exe)
    if getattr(sys, 'frozen', False):
        current_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 定义 Word 导出 PDF 的格式代码
    wdFormatPDF = 17
    
    # 获取当前目录下所有的 .doc 和 .docx 文件 (排除临时文件)
    files = [f for f in os.listdir(current_dir) 
             if f.lower().endswith(('.doc', '.docx')) and not f.startswith('~')]
    
    if not files:
        print("当前目录下没有找到 .doc 或 .docx 文件。")
        input("按回车键退出...")
        return

    print("=" * 60)
    print("欢迎使用 Word 转 PDF 工具 (doc2pdf)")
    print("注意：本程序必须依赖 Microsoft Word 才能正常工作。")
    print("请确保您已安装 Microsoft Office。")
    print("-" * 60)
    print(f"当前扫描目录: {current_dir}")
    print(f"共发现 {len(files)} 个可转换的文档 (.doc/.docx)")
    print("=" * 60)
    
    input("\n按回车键(Enter)开始转换...")

    print(f"\n正在启动 Word 应用程序...")

    word = None
    try:
        # 启动 Word 应用程序
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
    except Exception as e:
        print("无法启动 Word 应用程序，请确保已安装 Microsoft Word 并且安装了 'comtypes' 库。")
        print(f"错误详情: {e}")
        return

    try:
        for filename in files:
            file_path = os.path.join(current_dir, filename)
            # 获取不带后缀的文件名
            file_name_no_ext = os.path.splitext(filename)[0]
            
            # 获取唯一的输出路径
            output_pdf_path = get_unique_pdf_path(current_dir, file_name_no_ext)
            
            try:
                print(f"正在转换: {filename} -> {os.path.basename(output_pdf_path)}")
                # 打开文档
                doc = word.Documents.Open(file_path)
                # 另存为 PDF
                doc.SaveAs(output_pdf_path, FileFormat=wdFormatPDF)
                # 关闭文档
                doc.Close()
            except Exception as e:
                print(f"转换 {filename} 失败: {e}")
    finally:
        # 确保 Word 应用程序退出
        if word:
            word.Quit()
    
    print("所有任务已完成。")
    input("按回车键退出...")

if __name__ == "__main__":
    main()
