import os
import tkinter as tk


def close_execl():
    try:
        a = os.system("taskkill /im EXCEL.EXE")
        print(a)
    except Exception:
        print("报错了！")


def close_popint():
    a = os.system("taskkill /im POWERPNT.EXE")

    if a == 128:
        print("错误")


def close_word():
    os.system("taskkill /im WINWORD.EXE")


def init_tk():
    w = tk.Tk()
    w.title("小工具集合")
    w.config(bg="pink")
    w.geometry("400x400")

    photo_excel = tk.PhotoImage(file="Excel.png")  # file：t图片路径
    excel = tk.Button(text="Excel", image=photo_excel, command=close_execl, bd=5)
    excel.place(x=0, y=0)

    photo_word = tk.PhotoImage(file="word.png")  # file：t图片路径
    word = tk.Button(text="word", image=photo_word, command=close_word, bd=5)
    word.place(x=60, y=0)

    photo_popint = tk.PhotoImage(file="popint.png")  # file：t图片路径
    popint = tk.Button(text="popint", image=photo_popint, bd=5, command=close_popint)
    popint.place(x=120, y=0)
    w.mainloop()


def main():
    init_tk()

if __name__ == "__main__":
    main()

