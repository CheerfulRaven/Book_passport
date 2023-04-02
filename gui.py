"""Программа создает gui для ввода данных пользователем, с помощью функций new_isbn, new_order_number, new_edn и new_doi
присваивает коды, а затем заносит их в сводную таблицу 'Учет книг.xlsx'. Опционально можно отправить письма с кодами
через электронную почту"""

from tkinter import *
from tkinter.ttk import Checkbutton
from tkinter.messagebox import showinfo, askyesno
from ISBN import new_isbn
from order_number import new_order_number
from edn import new_edn
from DOI import new_doi
import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pathlib
from pathlib import Path

my_mail = 'raven-work-mail@yandex.ru'
password = 'vtfprepjmvpxmndg'
mail_1 = 'karodunum@gmail.com'
mail_2 = 'karodunum@gmail.com'

isbn = ''
book_number = ''
doi = ''
edn = ''
author = book_name = book_type = ''

#  вводим переменную s; если ISBN не присвоен s = 1, а EDN и DOI не будут присвоены
s = 0


# создаем функцию, вызываемую нажатием кнопки, для присвоения кодов
def clicked():
    # создаем переменные для номера заказа, EDN и DOI
    n = e = d = ''
    # Получаем ISBN
    i = new_isbn(txt1.get(), txt2.get(), txt3.get())
    # исключаем случай, когда ISBN закончились или документ 'ISBN_2023' открыт другой программой
    if i == "Добавьте ISBN" or i == "Закройте документ 'ISBN_2023'!":
        showinfo('Отчет', i)
        #  Изменяем переменную s, так как ISBN не присвоен
        global s
        s = 1
    else:
        s = 0
        # Получаем номер заказа
        n = new_order_number(txt1.get(), txt2.get(), txt3.get())
        # выводим окно подтверждения операции если документ 'Номера заказов_2023' открыт другой программой
        while True:
            if n == "Закройте документ 'Номера заказов_2023'!":
                result = askyesno(title="Ошибка", message="Документ 'Номера заказов_2023' открыт.\n"
                                                          "Закройте документ, а затем нажмите 'Да', "
                                                          "чтобы присвоить номер заказа. \nПри нажатии "
                                                          "'Нет' номер заказа не будет присвоен")
                if result:
                    n = new_order_number(txt1.get(), txt2.get(), txt3.get())
                else:
                    showinfo("Результат", "Операция отменена")
                    break
            else:
                break

    # Если edn_checkbutton включен, добавляем EDN
    if edn_value.get() == 1 and s != 1:
        e = new_edn(txt1.get(), txt2.get(), txt3.get())
        # выводим окно подтверждения операции если документ 'Номера заказов_2023' открыт другой программой
        # или EDN закончились.
        while True:
            if e == "Добавьте EDN!":
                result = askyesno(title="Ошибка", message="Закончились EDN!\n Пожалуйста, "
                                                          "добавьте в таблицу новые EDN")
                if result:
                    e = new_edn(txt1.get(), txt2.get(), txt3.get())
                else:
                    showinfo("Результат", "Операция отменена")

            if e == "Закройте документ 'EDN резерв'!":
                result = askyesno(title="Ошибка", message="Документ 'EDN резерв' открыт. Закройте документ, "
                                                          "а затем нажмите 'Да', чтобы присвоить EDN. "
                                                          "\nПри нажатии 'Нет' EDN не будет присвоен")
                if result:
                    e = new_edn(txt1.get(), txt2.get(), txt3.get())
                else:
                    showinfo("Результат", "Операция отменена")

            else:
                break

    # Если doi_checkbutton включен, добавляем DOI
    if doi_value.get() == 1 and s != 1:
        d = new_doi(i, txt1.get())
        # выводим окно подтверждения операции если документ 'DOI' открыт другой программой
        while True:

            if d == "Закройте документ 'doi'!":
                result = askyesno(title="Ошибка", message="Документ 'doi' открыт. Закройте документ, "
                                                          "а затем нажмите 'Да', чтобы присвоить doi. "
                                                          "\nПри нажатии 'Нет' DOI не будет присвоен")
                if result:
                    d = new_edn(txt1.get(), txt2.get(), txt3.get())
                else:
                    showinfo("Результат", "Операция отменена")

            else:
                break

    global author, book_name, book_type
    book_name = txt1.get()
    author = txt2.get()
    book_type = txt3.get()

    # Удаляем поля с кнопками и текстом
    if s != 1:
        lbl1.destroy()
        lbl2.destroy()
        lbl3.destroy()
        lbl4.destroy()
        txt1.destroy()
        txt2.destroy()
        txt3.destroy()
        btn.destroy()
        doi_checkbutton.destroy()
        edn_checkbutton.destroy()

        # выводим результат
        lbl5 = Label(window, text=f'ISBN: {i}\n\nНомер заказа: {n}\n\nEDN: {e}\n\nDOI: {d}',
                     bg="#fefbc6", font=("Times New Roman", 14), justify=LEFT)
        lbl5.place(relx=0.5, y=150, anchor=CENTER)

        email1_btn = Button(window, text="Отправить письмо редактору", command=clicked_email1)
        email1_btn.place(x=110, y=260)

        email2_btn = Button(window, text="Отправить письмо верстальщику", command=clicked_email2)
        email2_btn.place(x=320, y=260)

        finish_btn = Button(window, text="Занести данные в таблицу", command=clicked_2)
        finish_btn.place(x=210, y=320)
        global isbn, book_number, edn, doi
        isbn = i
        book_number = n
        edn = e
        doi = d


# создаем функцию, вызываемую нажатием кнопки, для занесения данных в таблицу "Учет книг.xlsx"
def clicked_2():
    wb = openpyxl.load_workbook('Учет книг.xlsx')
    sheet = wb.active
    data = [book_name, author, book_type, isbn, book_number, edn, doi]
    sheet.append(data)

    # следим за тем, чтобы документ не был открыт в другой программе; сохраняем документ
    try:
        wb.save("Учет книг.xlsx")
        result = askyesno(title='Отчет', message="Данные успешно занесены в таблицу.\nЗакрыть программу?")
        if result:
            # закрываем программу
            window.destroy()
        else:
            pass
        return result

    except PermissionError:
        return showinfo('Отчет', "Закройте документ 'Учет книг'!")


# создаем функции, вызываемую нажатием кнопки, для отправки почты
def clicked_email1():
    msg = MIMEMultipart()
    msg['From'] = my_mail
    msg['To'] = mail_1
    msg['Subject'] = f'Оформление {book_name}'
    # создаем текст письма
    message = f"{author}\n{book_name}\n{book_type}\nНомер заказа: {book_number}\nISBN: {isbn}\nEDN: {edn}\nDOI: {doi}"
    msg.attach(MIMEText(message))
    # формируем путь к файлу EDN
    if edn != '' and edn != "Добавьте EDN!" and edn != "Закройте документ 'EDN резерв'!":
        file_name = edn.replace('https://elibrary.ru/', '').upper()
        # указываем путь к файлу с EDN-кодом
        path = Path(pathlib.Path.cwd(), 'EDN', f"{file_name}.png")
        # добавляем вложение
        img = MIMEApplication(open(path, 'rb').read())
        img.add_header('Content-Disposition', 'attachment', filename=f'{file_name}.png')
        msg.attach(img)

    try:
        mailserver = smtplib.SMTP('smtp.yandex.ru', 587)
    # Определяем, поддерживает ли сервер TLS
        mailserver.ehlo()
    # Защищаем соединение с помощью шифрования tls
        mailserver.starttls()
    # Повторно идентифицируем себя как зашифрованное соединение перед аутентификацией.
        mailserver.ehlo()
        mailserver.login(my_mail, password)
        mailserver.sendmail(my_mail, mail_1, msg.as_string())
        mailserver.quit()
        showinfo('Отчет', 'Письмо успешно отправлено')
    except smtplib.SMTPException:
        showinfo('Отчет', 'Ошибка: Невозможно отправить сообщение')


def clicked_email2():
    msg = MIMEMultipart()
    msg['From'] = my_mail
    msg['To'] = mail_2
    msg['Subject'] = f'Оформление {book_name}'
    message = f"{author}\n{book_name}\n{book_type}\nНомер заказа: {book_number}\nISBN: {isbn}\nEDN: {edn}\nDOI: {doi}"
    msg.attach(MIMEText(message))
    if edn != '' and edn != "Добавьте EDN!" and edn != "Закройте документ 'EDN резерв'!":
        file_name = edn.replace('https://elibrary.ru/', '').upper()
        path = Path(pathlib.Path.cwd(), 'EDN', f"{file_name} + '.png'")
        img = MIMEApplication(open(path, 'rb').read())
        img.add_header('Content-Disposition', 'attachment', filename=f'{file_name}.png')
        msg.attach(img)
    try:
        mailserver = smtplib.SMTP('smtp.yandex.ru', 587)
        mailserver.ehlo()
        mailserver.starttls()
        mailserver.ehlo()
        mailserver.login(my_mail, password)
        mailserver.sendmail(my_mail, mail_2, msg.as_string())
        mailserver.quit()
        showinfo('Отчет', 'Письмо успешно отправлено')
    except smtplib.SMTPException:
        showinfo('Отчет', 'Ошибка: Невозможно отправить сообщение')


# создаем gui
window = Tk()
window.title("Book passport")
# задаем размер окна
window.geometry("600x400")
# задаем цвет окна
window['background'] = '#fefbc6'

# добавляем текст
lbl1 = Label(window, text="Название книги", bg="#fefbc6", font=("Arial", 10))
lbl1.place(x=45, y=50)

lbl2 = Label(window, text="Автор", bg="#fefbc6", font=("Arial", 10))
lbl2.place(x=75, y=100)

lbl3 = Label(window, text="Вид издания", bg="#fefbc6", font=("Arial", 10))
lbl3.place(x=55, y=150)

warning = "Внимание! Для присвоения DOI необходим перевод названия, аннотации, имени и \n фамилии автора на английский " \
          "язык. \n Для присвоения EDN необходимо, чтобы книга была научным или учебным изданием. \n У книги должен " \
          "быть рецензент либо книга должна быть сборником статей конференции."

lbl4 = Label(window, text=warning, bg="#fefbc6", font=("Arial", 8))
lbl4.place(x=48, y=195)

# создаем всплывающие окна для меню
x = y = 0


def popup_1(event):
    global x, y
    x = event.x
    y = event.y
    menu_1.post(event.x_root, event.y_root)


def popup_2(event):
    global x, y
    x = event.x
    y = event.y
    menu_2.post(event.x_root, event.y_root)


def popup_3(event):
    global x, y
    x = event.x
    y = event.y
    menu_3.post(event.x_root, event.y_root)


# создаем класс для меню
# класс содержит функции для копирования строки, вставки строки из буфера обмена и очистки строки
class ClickMenu:
    def __init__(self, e):
        self.e = e

    def paste(self):
        try:
            clipboard = self.e.clipboard_get()     # получаем строку из буфера обмена
            self.e.insert('end', clipboard)    # вставляем строку в поле для ввода текста
        except TclError:
            pass

    def copy(self):
        inp = self.e.get()    # получаем строку из поля для ввода текста
        self.e.clipboard_append(inp)    # добавляем строку в буфер обмена

    def delete(self):
        self.e.delete(0, END)
        self.e.insert(0, "")


# добавляем поля для ввода текста
txt1 = Entry(window, width=50, font=("Arial", 10))
txt1.place(x=150, y=50)
txt1.bind("<Button-3>", popup_1)

# добавляем меню, которое выводится нажатием правой кнопки мыши на текстовое поле
menu_1 = Menu(txt1, tearoff=0)
menu_1_click = ClickMenu(txt1)
menu_1.add_command(label="Вставить", command=menu_1_click.paste)
menu_1.add_command(label="Копировать", command=menu_1_click.copy)
menu_1.add_command(label="Удалить", command=menu_1_click.delete)

txt2 = Entry(window, width=50, font=("Arial", 10))
txt2.place(x=150, y=100)
txt2.bind("<Button-3>", popup_2)

menu_2 = Menu(txt2, tearoff=0)
menu_2_click = ClickMenu(txt2)
menu_2.add_command(label="Вставить", command=menu_2_click.paste)
menu_2.add_command(label="Копировать", command=menu_2_click.copy)
menu_2.add_command(label="Удалить", command=menu_2_click.delete)

txt3 = Entry(window, width=50, font=("Arial", 10))
txt3.place(x=150, y=150)
txt3.bind("<Button-3>", popup_3)

menu_3 = Menu(txt3, tearoff=0)
menu_3_click = ClickMenu(txt3)
menu_3.add_command(label="Вставить", command=menu_3_click.paste)
menu_3.add_command(label="Копировать", command=menu_3_click.copy)
menu_3.add_command(label="Удалить", command=menu_3_click.delete)


# Добавляем кнопку
btn = Button(window, text="Присвоить ISBN и номер заказа", command=clicked)
btn.place(x=210, y=320)

# Добавляем checkbutton
edn_value = IntVar()
edn_value.set(0)
edn_checkbutton = Checkbutton(text="Присвоить EDN", variable=edn_value, onvalue=1)
edn_checkbutton.place(x=170, y=275)
doi_value = IntVar()
doi_value.set(0)
doi_checkbutton = Checkbutton(text="Присвоить DOI", variable=doi_value, onvalue=1)
doi_checkbutton.place(x=350, y=275)

window.mainloop()