
##Program Information
##Developer Information: Fly
##Developer Institution: JiangNan University (In WuXi, Jiangsu Provience, China)
##Contact me: COTOMO on GitHub
##Buy me a cup of coffee: Thanks, but no need currently


import docx
import os
import getpass
from os import system
import win32com.client

def save_doc_to_docx(path):  # doc转docx
    # 创建Word应用程序对象
    word_app = win32com.client.Dispatch("Word.Application")

    try:
        # 打开doc文件
        doc = word_app.Documents.Open(path)

        # 将doc文件另存为docx
        doc.SaveAs2(path + "x", FileFormat=16)  # 16 表示docx格式

        # 关闭doc文件
        doc.Close()

    except Exception as e:
        print("在将doc转换为docx时发生错误: {e}")

    finally:
        # 退出Word应用程序
        word_app.Quit()

class DocOperator:
    
    @staticmethod
    def main(doc):
        for paragraph in doc.paragraphs:
            text = paragraph.text
            start_spaces = len(text) - len(text.lstrip(' '))
            end_spaces = len(text) - len(text.rstrip(' '))
            total_length = len(text)

            if start_spaces > 0 and start_spaces < 6:  # 首行缩进
                DocOperator.Indent(paragraph)
            if start_spaces > 6 and (total_length - start_spaces - end_spaces) <= 15: # 居中排版
                DocOperator.Center(paragraph)    # 文件保存操作应该在外部进行

    @staticmethod
    def Indent(paragraph):
        paragraph.text = paragraph.text.lstrip(' ')
        # 设置首行缩进
        paragraph.paragraph_format.first_line_indent = docx.shared.Pt(24)  # 示例：24磅

    @staticmethod
    def Center(paragraph):
        paragraph.text = paragraph.text.strip()
        paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

class Informations:
    @staticmethod
    def Hello():
        print("\33[31m去你的空格排版法！\33[0m")
        usrnm = getpass.getuser()
        print("你好，\33[36m" + usrnm + "\33[0m!")
        print("就目前而言,此程序可以将您文档中的使用空格进行首行缩进和居中的排版按规范的方式重设--当然,它也只能做这么多。")
        print("此程序由\33[36mCOTOMO\33[0m开发和维护。")
        # 更多信息...

    @staticmethod
    def IsDocx(path):
        if path.lower().endswith(".docx"):
            return 0
        elif path.lower().endswith(".doc"):
            print("程序只能处理docx文件。您的doc文件会变成docx文件。")
            save_doc_to_docx(path)
            return 1
        else:
            print("这个程序只能处理docx文件和doc文件，这样是会出问题的。")
            return 2
        
class Configure:
    def developing():
        print("此功能正在开发中...")
    def main():
        CurrentSelection = 0
        Text = {"1.首行缩进\n", "2. 居中排版\n", "3. doc在处理完成后不再还原后缀\n", "4.保存为新文件\n", "5.退出\n"}
        Description = {"选择程序是否会处理错误的首行缩进","选择程序是否会处理错误的居中排版","选择程序是否会在处理完成后不再还原后缀","选择程序是否会保存为新文件","退出程序"}
        while(1):
            print("配置页面")
            print("可以自由地配置此程序的行为。我保证这无害。你可以使用W上移动光标，S下移动光标，以及使用C来更改此选项的配置")
            #列出所有的选择,并高亮当前选中的选项
            for i in range(len(Text)):
                if i == CurrentSelection:
                    print("\33[36m" + Text[i] + "\33[0m")
                else:
                    print(Text[i])
            print(' ')
            print(' ')
            print(Description[CurrentSelection])
            #获取用户输入
            key = input()
            #处理用户输入


#主要程序
Informations.Hello()
FilePath = input("请把需要操作的文件拖进来，然后按Enter键：").strip('"')
if FilePath == "":
# 调用配置类的主要方法
    Configure.developing
elif Informations.IsDocx(FilePath) == 1:
    FilePath = FilePath + "x"
elif Informations.IsDocx(FilePath) == 2:
    print("程序无法处理此文件。")
    system("pause")

doc = docx.Document(FilePath) # 打开文件
print(f"\33[31m成功地打开了文件 {FilePath} \33[0m此文件包含了 \33[37m{len(doc.paragraphs)}\33[0m 个Paragraph")
print("开始处理文件...")
DocOperator.main(doc)
# 保存文件
doc.save(FilePath.strip(".docx") + "_处理完成.docx")
print("处理完成！文件保存在了" + FilePath)
print("如果你输入了doc文件的话，还会有一个没有经过原始排版修改过的docx文件哦。")
print("可以关掉我啦")
input()
