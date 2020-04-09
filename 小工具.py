import os
import tkinter as tk
import datetime


def init_tk():
    global list_text
    w = tk.Tk()
    w.title("小工具集合")
    w.config(bg="pink")
    w.geometry("400x400")
    # 创建Excel按钮
    photo_excel = tk.PhotoImage(file="Excel.png")  # file：t图片路径
    excel = tk.Button(text="Excel", image=photo_excel, command=close_execl, bd=5)
    excel.grid()
    # 创建Word按钮
    photo_word = tk.PhotoImage(file="word.png")  # file：t图片路径
    word = tk.Button(text="word", image=photo_word, command=close_word, bd=5)
    word.grid(row=0, column=1)
    # 创建PowerPoint按钮
    photo_powerpoint = tk.PhotoImage(file="point.png")  # file：t图片路径
    powerpoint = tk.Button(text="point", image=photo_powerpoint, bd=5, command=close_powerpoint)
    powerpoint.grid(row=0, column=2)
    # 创建列表框
    list_text = tk.Listbox(w, font=("微软雅黑", 10), width=27, height=3)
    list_text.grid(row=0, column=3)
    w.mainloop()


def close_execl():
    a = os.system("taskkill /im EXCEL.EXE")
    if a == 128:
        b = str(datetime.datetime.now())
        b = b.split(" ")[1]
        b = b.split(".")[0]
        list_text.insert("end", "Excel程序未运行!{}".format(b))
        list_text.see("end")
        print("Excel程序未运行")


def close_powerpoint():
    a = os.system("taskkill /im POWERPNT.EXE")
    if a == 128:
        b = str(datetime.datetime.now())
        b = b.split(" ")[1]
        b = b.split(".")[0]
        list_text.insert("end", "Point程序未运行!{}".format(b))
        list_text.see("end")
        print("Point程序未运行")


def close_word():
    a = os.system("taskkill /im WINWORD.EXE")
    if a == 128:
        b = str(datetime.datetime.now())
        b = b.split(" ")[1]
        b = b.split(".")[0]
        list_text.insert("end", "Word程序未运行!{}".format(b))
        list_text.see("end")
        print("Word程序未运行！")


def main():
    init_tk()


if __name__ == "__main__":
    main()
