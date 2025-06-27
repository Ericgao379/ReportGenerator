import os
import comtypes.client
import sys

def get_exe_dir():
    # 如果是打包成exe，使用 sys._MEIPASS 或 sys.argv[0]
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(__file__)

def get_path_name():
    base_dir = get_exe_dir()
    input_path = os.path.join(base_dir, "wordFiles")
    output_path = os.path.join(base_dir, "pdfFiles")

    if not os.path.exists(output_path):
        os.makedirs(output_path)

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"未找到输入文件夹: {input_path}")

    filename_list = os.listdir(input_path)
    wordname_list = [name for name in filename_list if name.endswith('.docx')]

    for wordname in wordname_list:
        word_full_path = os.path.join(input_path, wordname)
        pdfname = os.path.splitext(wordname)[0] + '.pdf'
        pdf_full_path = os.path.join(output_path, pdfname)
        yield word_full_path, pdf_full_path

def word_convert_to_pdf():
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = 0
    for word_path, pdf_path in get_path_name():
        doc = word.Documents.Open(word_path)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
    word.Quit()

if __name__ == '__main__':
    word_convert_to_pdf()

print("done")