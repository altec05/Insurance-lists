from tkinter import messagebox


def error(title, text):
    mb = messagebox.showerror(title=title, message=text)


def info(title, text):
    mb = messagebox.showinfo(title=title, message=text)


def warning(title, text):
    mb = messagebox.showwarning(title=title, message=text)


def ask(title, text):
    mb = messagebox.askyesno(title=title, message=text, icon='warning')
    if mb == 'yes':
        return True
    else:
        return False
