from tkinter import *
import os
import shutil
from threading import Timer

def sync_folders(src_folder, dest_folder):
    # 遍历源文件夹中的所有文件和子文件夹
    for item in os.listdir(src_folder):
        src_path = os.path.join(src_folder, item)
        dest_path = os.path.join(dest_folder, item)

        # 如果是文件夹，则递归调用此函数
        if os.path.isdir(src_path):
            if not os.path.exists(dest_path):
                os.makedirs(dest_path)
            sync_folders(src_path, dest_path)
        else:
            # 如果是文件，则复制到目标文件夹中
            shutil.copy2(src_path, dest_path)

root = Tk()
root.geometry('460x240')
root.title('文件夹同步工具')

lb1 = Label(root, text='源地址')
lb1.place(relx=0.1, rely=0.1, relwidth=0.3, relheight=0.1)
lb2 = Label(root, text='目标地址')
lb2.place(relx=0.6, rely=0.1, relwidth=0.3, relheight=0.1)

inp1 = Entry(root)
inp1.place(relx=0.1, rely=0.3, relwidth=0.3, relheight=0.1)
inp2 = Entry(root)
inp2.place(relx=0.6, rely=0.3, relwidth=0.3, relheight=0.1)


btn2 = Button(root, text='开始同步', command=lambda: sync_folders(inp1.get(), inp2.get()))
btn2.place(relx=0.6, rely=0.6, relwidth=0.3, relheight=0.1)


root.mainloop()


