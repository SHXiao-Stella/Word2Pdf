#在前人代码（https://www.jianshu.com/p/fed949abe811）的基础上修正了“被呼叫方拒绝接收呼叫”问题

from win32com.client import constants,gencache,pywintypes
import os
import time

def Word_to_Pdf(Word_path,Pdf_path): # Word转pdf方法,第一个参数代表word文档路径，第二个参数代表pdf文档路径
    word = gencache.EnsureDispatch('Word.Application')
    word.Visible = 0  #不显示
    doc = word.Documents.Open(Word_path,ReadOnly = 1)
    doc.ExportAsFixedFormat(Pdf_path,constants.wdExportFormatPDF)
    word.Documents.Close(SaveChanges=0)
    word.Quit()
    

Word_files = []
for file in os.listdir('.'):
    if file.endswith(('.doc','.docx')):
        Word_files.append(file)
#or
#path='D:\\文件夹'
#for file in os.listdir(path):
#    if file.endswith(('.doc','.docx')):
#        Word_files.append(file)
        
        
count=0
for file in Word_files:
    count += 1
    try:
        file_path = os.path.abspath(file)
        index = file_path.rindex('.')
        pdf_path = file_path[:index] + '.pdf'
        Word_to_Pdf(file_path, pdf_path)
        print(count)
        time.sleep(4)
        #if count % 100 == 0: #每转换100个print1次
        #    print 'did 100'
    except pywintypes.com_error:
        print(pdf_path)
        continue
