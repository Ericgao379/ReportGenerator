import os
import comtypes.client
def get_path_name():
    # 当前脚本所在目录
    base_dir = os.path.dirname(__file__)
    # 拼接 wordFiles 路径
    path = os.path.join(base_dir, "wordFiles")
    filename_list = os.listdir(path)
    wordname_list = [wordname for wordname in filename_list if wordname.endswith('docx')]
    for wordname in wordname_list:
        pdfname = os.path.splitext(wordname)[0] + '.pdf'
        word_full_path =  os.path.join(path, wordname)
        pdf_full_path = os.path.join(path, pdfname)
        #生成器
        yield word_full_path, pdf_full_path
def word_convert_to_pdf():
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = 0
    for wordname,pdfname in get_path_name():
        wordname = word.Documents.Open(wordname)
        wordname.SaveAs(pdfname, FileFormat = 17)
        wordname.Close()

if __name__ == '__main__':
    word_convert_to_pdf()