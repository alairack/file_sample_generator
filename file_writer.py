import os
import random

import pypandoc
import win32com.client


class FileWriter(object):
    def __init__(self, file_content: str, file_name: str, folder: str, is_secret=False):
        self.sample_path = "./sample/" + folder
        self.file_name = file_name
        self.file_content = file_content
        if not os.path.exists(self.sample_path):
            os.makedirs(self.sample_path)
        self.output_file_path = os.path.abspath(os.path.join(self.sample_path, self.file_name))
        file_type = file_name.split(".")[1]

        if is_secret:
            self.add_secret_label()

        func_dict = {"doc": self.write_doc_and_docx_and_ofd, "docx": self.write_doc_and_docx_and_ofd, "pptx": self.write_pptx,
                     'pdf': self.write_pdf, 'ppt': self.write_ppt, "xls": self.write_xls_and_xlsx,
                     'xlsx': self.write_xls_and_xlsx, 'ofd': self.write_doc_and_docx_and_ofd, 'wps': self.write_wps,
                     'et': self.write_et}
        func_dict.get(file_type)()

    def write_to_file(self, file_content):
        with open(self.output_file_path, 'w') as f:
            f.write(file_content)

    def add_secret_label(self):
        """
        添加密标
        :return:
        """
        secret_label = ["秘密 10年", "秘密★5年", "机密 20年", "机密★长期", "绝密", "绝密★30年"]
        random_index = random.randrange(len(secret_label))
        mark_text = secret_label[random_index]
        self.file_content = mark_text + "\n" + self.file_content


    def write_doc_and_docx_and_ofd(self):
        word = win32com.client.Dispatch("Word.Application")
        # 让文档可创建
        word.Visible = True
        # 创建文档
        doc = word.Documents.Add()

        # 写内容,定位都最开始
        r = doc.Range(0, 0)
        # 插入内容
        r.InsertAfter(self.file_content)

        # 存储文件
        doc.SaveAs(self.output_file_path)
        doc.Close()

    def write_pptx(self):
        output = pypandoc.convert_text(self.file_content, 'pptx', format='html', outputfile=self.output_file_path)

    def write_ppt(self):
        Application = win32com.client.Dispatch("PowerPoint.Application")

        # 创建一个新的演示文稿
        Presentation = Application.Presentations.Add()

        slide = Presentation.Slides.Add(1, 11)

        # 添加文本框并设置文本
        left = top = width = height = 100
        shape = slide.Shapes.AddTextbox(
            Orientation=0x01,
            Left=left,
            Top=top,
            Width=width,
            Height=height,
        )
        shape.TextFrame.TextRange.Text = self.file_content

        # 保存演示文稿
        Presentation.SaveAs(self.output_file_path)

        # 关闭
        Presentation.Close()

    def write_xls_and_xlsx(self):
        Application = win32com.client.Dispatch("Excel.Application")

        # 使Excel应用程序可见（这样你可以看到它）
        Application.Visible = True

        # 创建一个新的工作簿
        Workbook = Application.Workbooks.Add()

        # 获取活动的工作表
        Worksheet = Workbook.ActiveSheet

        # 在工作表的A1单元格写入数据
        Worksheet.Range("A1").Value = self.file_content

        # 保存工作簿
        Workbook.SaveAs(self.output_file_path)

        # 关闭
        Workbook.Close()

    def write_pdf(self):
        output = pypandoc.convert_text(self.file_content, 'pdf', format='html', outputfile=self.output_file_path, extra_args=['--pdf-engine=E:\\miktex\\miktex\\bin\\x64\\xelatex.exe', "-V",  'mainfont:FangSong'])

    def write_wps(self):
        wps = win32com.client.Dispatch("Kwps.Application")
        wps.Visible = True
        doc = wps.Documents.Add()
        doc.Content.text = self.file_content
        doc.SaveAs(self.output_file_path)
        doc.Close()

    def write_et(self):
        et = win32com.client.Dispatch("Ket.Application")
        et.Visible = True
        doc = et.Workbooks.Add()
        sht = doc.Worksheets(1)
        sht.Cells(1, 1).Value = self.file_content
        doc.SaveAs(self.output_file_path)
        doc.Close()


if __name__ == "__main__":
    tt = "hello word"
    print(pypandoc.get_pandoc_formats()[1])
    FileWriter("中不中", "1.pdf", 'pdf', True)
    FileWriter(tt, '0.dps')
    FileWriter(tt, '0.et')
    FileWriter(tt, "0.wps")
    FileWriter("hello word", "1.pdf")
    FileWriter("hello word", '2.docx')
    FileWriter("hello word", '3.ppt')
    FileWriter(tt, '4.pptx')
    FileWriter(tt, '5.xls')
    FileWriter(tt, '6.xlsx')
    FileWriter(tt, '7.ofd')