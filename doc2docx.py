import win32com.client as win32
import os
import glob
import re

paths = glob.glob('C:\\Users\\bsp\\python\\doc2docx\\files\\*.doc', recursive=True)
print('Преобразовываем .doc в .docx файл:')


def save_as_docx(path):
    print(path)
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path)
# word.Visible = True
    new_file_abs = os.path.abspath(path)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
    doc.Activate()
    word.ActiveDocument.SaveAs(new_file_abs, FileFormat=win32.constants.wdFormatXMLDocument)
    doc.Close(True)


for path in paths:
    save_as_docx(path)

#win32.gencache.EnsureDispatch('Word.Application').Quit()

"""word = win32.gencache.EnsureDispatch('Word.Application')
doc = word.Documents.Open("C:\\Users\\bsp\\python\\aosr\\1.doc")
word.Visible = True
doc.Activate()
word.ActiveDocument.SaveAs("C:\\Users\\bsp\\python\\aosr\\2.docx", FileFormat=win32.constants.wdFormatXMLDocument)
doc.Close(True)
word.Quit()"""
