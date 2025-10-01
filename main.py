import camelot
import openpyxl 
import telebot
import requests
from datetime import datetime, timedelta 
from telebot import types
import os
import urllib.request

token='8347380655:AAE56FocrVCTzAY39vc4QOo9Oz0IsZttcBw'
bot=telebot.TeleBot(token)


#####

#####################







@bot.message_handler(content_types=['text'])
def aaa(message):
    if(message.text == "/start"):
        keyboard = telebot.types.ReplyKeyboardMarkup(resize_keyboard=True)
        button_support = telebot.types.KeyboardButton(text="Расписание")
        keyboard.add(button_support)
        bot.send_message(message.chat.id, "Привет", reply_markup=keyboard) 

    elif(message.text == "Расписание"):
        

        kb1 = telebot.types.InlineKeyboardMarkup()#Клавиатуа


        #Добавляем кнопки
        kb1b1= telebot.types.InlineKeyboardButton(text=f"{need_day(1)}",callback_data=f"{need_day(1)}")
        kb1b2= telebot.types.InlineKeyboardButton(text=f"{need_day(2)}",callback_data=f"{need_day(2)}")
        kb1b3= telebot.types.InlineKeyboardButton(text=f"{need_day(3)}",callback_data=f"{need_day(3)}")
        kb2b1= telebot.types.InlineKeyboardButton(text=f"{need_day(-1)}",callback_data=f"{need_day(-1)}")
        kb2b2= telebot.types.InlineKeyboardButton(text=f"{need_day(-2)}",callback_data=f"{need_day(-2)}")
        kb2b3= telebot.types.InlineKeyboardButton(text=f"{need_day(-3)}",callback_data=f"{need_day(-3)}")
        kb3b1= telebot.types.InlineKeyboardButton(text=f"<--",callback_data=f">")
        kb3b2= telebot.types.InlineKeyboardButton(text=f"Сегодня",callback_data=f"now")
        kb3b3= telebot.types.InlineKeyboardButton(text=f"-->",callback_data=f"<")

        #Вставляем в клавиатуру
        kb1.add(kb1b1,kb1b2,kb1b3,kb2b1,kb2b2,kb2b3,kb3b1,kb3b2,kb3b3)

        bot.send_message(message.chat.id, "Привет", reply_markup=kb1) #Выводим клавиатуру


@bot.callback_query_handler(func=lambda call: True)
def callback_query(call):

    if call.data == "now":
        url=f"https://gtnu.ru/wp-content/uploads/rasp/{datetime_full(0)}.pdf"
        urllib.request.urlretrieve(url, f'{datetime_full(0)}.pdf')
        



        desting = f'{datetime_full(0)}.xlsx'

        a = camelot.read_pdf(f'{datetime_full(0)}.pdf')
        os.remove(f'{datetime_full(0)}.pdf')
        a[1].df.to_excel(desting)
        
        #открываем фаил exel
        fff=openpyxl.load_workbook(desting) 
        os.remove(f'{datetime_full(0)}.xlsx') 
        f=fff.active
        g=[]

        #Ищем сколько всего пар
        for i in range(f.max_row-2):
            if f.cell(row=f.max_row-i,column=5).value is None:
                exit
            else:
                g.append([int(f.cell(row=f.max_row-i,column=5).value),[f.max_row-i,5]])
        g = g[::-1]    
        #Состовляем сообщение
        mes=""
        for i in range(len(g)):
            if None is f.cell(row=g[i][1][0],column=g[i][1][1]+1).value:
                mes = mes + f"\n{g[i][0]}          {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value}  -  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value}\n          {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value}\n"

            else:

                mes = mes + f"\n{g[i][0]}          {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[0]}  -  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines()[len(f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines())-1]}\n          {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[1]}\n"
        bot.send_message(call.from_user.id, f"Расписание на {need_day(-3)}\n"+mes)#Отправляем
    elif call.data == f"<":



        bot.send_message(call.from_user.id, "Пока не работает, терпи.")

    elif call.data == f">":
    


        bot.send_message(call.from_user.id, "Пока не работает, терпи.")

    elif call.data == f"{need_day(1)}":
        url=f"https://gtnu.ru/wp-content/uploads/rasp/{datetime_full(1)}.pdf"
        urllib.request.urlretrieve(url, f'{datetime_full(1)}.pdf')
        



        desting = f'{datetime_full(1)}.xlsx'

        a = camelot.read_pdf(f'{datetime_full(1)}.pdf')
        os.remove(f'{datetime_full(1)}.pdf')
        a[1].df.to_excel(desting)
        
        #открываем фаил exel
        fff=openpyxl.load_workbook(desting) 
        os.remove(f'{datetime_full(1)}.xlsx') 
        f=fff.active
        g=[]

        #Ищем сколько всего пар
        for i in range(f.max_row-2):
            if f.cell(row=f.max_row-i,column=5).value is None:
                exit
            else:
                g.append([int(f.cell(row=f.max_row-i,column=5).value),[f.max_row-i,5]])
        g = g[::-1]    

        #Состовляем сообщение
        mes=""
        for i in range(len(g)):
            if None is f.cell(row=g[i][1][0],column=g[i][1][1]+1).value:
                mes = mes + f"\n{g[i][0]}          {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value}  -  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value}\n          {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value}\n"

            else:

                mes = mes + f"\n{g[i][0]}          {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[0]}  -  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines()[len(f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines())-1]}\n          {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[1]}\n"
        bot.send_message(call.from_user.id, f"Расписание на {need_day(1)}\n"+mes)#Отправляем




    elif call.data == f"{need_day(2)}":
        url=f"https://gtnu.ru/wp-content/uploads/rasp/{datetime_full(2)}.pdf"
        urllib.request.urlretrieve(url, f'{datetime_full(2)}.pdf')
        



        desting = f'{datetime_full(2)}.xlsx'

        a = camelot.read_pdf(f'{datetime_full(2)}.pdf')
        os.remove(f'{datetime_full(2)}.pdf')
        a[1].df.to_excel(desting)
        
        #открываем фаил exel
        fff=openpyxl.load_workbook(desting) 
        os.remove(f'{datetime_full(2)}.xlsx') 
        f=fff.active
        g=[]

        #Ищем сколько всего пар
        for i in range(f.max_row-2):
            if f.cell(row=f.max_row-i,column=5).value is None:
                exit
            else:
                g.append([int(f.cell(row=f.max_row-i,column=5).value),[f.max_row-i,5]])
        g = g[::-1]    
        #Состовляем сообщение
        mes=""
        for i in range(len(g)):
            if None is f.cell(row=g[i][1][0],column=g[i][1][1]+1).value:
                mes = mes + f"\n{g[i][0]}          {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value}  -  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value}\n          {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value}\n"

            else:

                mes = mes + f"\n{g[i][0]}          {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[0]}  -  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines()[len(f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines())-1]}\n          {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[1]}\n"
        bot.send_message(call.from_user.id, f"Расписание на {need_day(2)}\n"+mes)#Отправляем




    elif call.data == f"{need_day(3)}":
        url=f"https://gtnu.ru/wp-content/uploads/rasp/{datetime_full(3)}.pdf"
        urllib.request.urlretrieve(url, f'{datetime_full(3)}.pdf')
        



        desting = f'{datetime_full(3)}.xlsx'

        a = camelot.read_pdf(f'{datetime_full(3)}.pdf')
        os.remove(f'{datetime_full(3)}.pdf')
        a[1].df.to_excel(desting)
        
        #открываем фаил exel
        fff=openpyxl.load_workbook(desting) 
        os.remove(f'{datetime_full(3)}.xlsx') 
        f=fff.active
        g=[]

        #Ищем сколько всего пар
        for i in range(f.max_row-2):
            if f.cell(row=f.max_row-i,column=5).value is None:
                exit
            else:
                g.append([int(f.cell(row=f.max_row-i,column=5).value),[f.max_row-i,5]])
        g = g[::-1]    
        #Состовляем сообщение
        mes=""
        for i in range(len(g)):
            if None is f.cell(row=g[i][1][0],column=g[i][1][1]+1).value:
                mes = mes + f"\n{g[i][0]}          {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value}  -  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value}\n          {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value}\n"

            else:

                mes = mes + f"\n{g[i][0]}          {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[0]}  -  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines()[len(f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines())-1]}\n          {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[1]}\n"
        bot.send_message(call.from_user.id, f"Расписание на {need_day(3)}\n"+mes)#Отправляем





    elif call.data == f"{need_day(-1)}":
        url=f"https://gtnu.ru/wp-content/uploads/rasp/{datetime_full(-1)}.pdf"
        urllib.request.urlretrieve(url, f'{datetime_full(-1)}.pdf')
        



        desting = f'{datetime_full(-1)}.xlsx'

        a = camelot.read_pdf(f'{datetime_full(-1)}.pdf')
        os.remove(f'{datetime_full(-1)}.pdf')
        a[1].df.to_excel(desting)
        
        #открываем фаил exel
        fff=openpyxl.load_workbook(desting) 
        os.remove(f'{datetime_full(-1)}.xlsx') 
        f=fff.active
        g=[]

        #Ищем сколько всего пар
        for i in range(f.max_row-2):
            if f.cell(row=f.max_row-i,column=5).value is None:
                exit
            else:
                g.append([int(f.cell(row=f.max_row-i,column=5).value),[f.max_row-i,5]])
        g = g[::-1]    
        #Состовляем сообщение
        mes=""
        for i in range(len(g)):
            if None is f.cell(row=g[i][1][0],column=g[i][1][1]+1).value:
                mes = mes + f"\n{g[i][0]}          {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value}  -  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value}\n          {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value}\n"

            else:

                mes = mes + f"\n{g[i][0]}          {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[0]}  -  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines()[len(f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines())-1]}\n          {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[1]}\n"
        bot.send_message(call.from_user.id, f"Расписание на {need_day(-1)}\n"+mes)#Отправляем












    elif call.data == f"{need_day(-2)}":
        url=f"https://gtnu.ru/wp-content/uploads/rasp/{datetime_full(-2)}.pdf"
        urllib.request.urlretrieve(url, f'{datetime_full(-2)}.pdf')
        



        desting = f'{datetime_full(-2)}.xlsx'

        a = camelot.read_pdf(f'{datetime_full(-2)}.pdf')
        os.remove(f'{datetime_full(-2)}.pdf')
        a[1].df.to_excel(desting)
        
        #открываем фаил exel
        fff=openpyxl.load_workbook(desting) 
        os.remove(f'{datetime_full(-2)}.xlsx') 
        f=fff.active
        g=[]

        #Ищем сколько всего пар
        for i in range(f.max_row-2):
            if f.cell(row=f.max_row-i,column=5).value is None:
                exit
            else:
                g.append([int(f.cell(row=f.max_row-i,column=5).value),[f.max_row-i,5]])
        g = g[::-1]    
        #Состовляем сообщение
        mes=""
        for i in range(len(g)):
            if None is f.cell(row=g[i][1][0],column=g[i][1][1]+1).value:
                mes = mes + f"\n{g[i][0]}          {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value}  -  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value}\n          {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value}\n"

            else:

                mes = mes + f"\n{g[i][0]}          {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[0]}  -  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines()[len(f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines())-1]}\n          {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[1]}\n"
        bot.send_message(call.from_user.id, f"Расписание на {need_day(-2)}\n"+mes)#Отправляем










    elif call.data == f"{need_day(-3)}":
    

        url=f"https://gtnu.ru/wp-content/uploads/rasp/{datetime_full(-3)}.pdf"
        urllib.request.urlretrieve(url, f'{datetime_full(-3)}.pdf')
        



        desting = f'{datetime_full(-3)}.xlsx'

        a = camelot.read_pdf(f'{datetime_full(-3)}.pdf')
        os.remove(f'{datetime_full(-3)}.pdf')
        a[1].df.to_excel(desting)
        
        #открываем фаил exel
        fff=openpyxl.load_workbook(desting) 
        os.remove(f'{datetime_full(-3)}.xlsx') 
        f=fff.active
        g=[]

        #Ищем сколько всего пар
        for i in range(f.max_row-2):
            if f.cell(row=f.max_row-i,column=5).value is None:
                exit
            else:
                g.append([int(f.cell(row=f.max_row-i,column=5).value),[f.max_row-i,5]])
        g = g[::-1]    
        #Состовляем сообщение
        mes=""
        for i in range(len(g)):
            if None is f.cell(row=g[i][1][0],column=g[i][1][1]+1).value:
                mes = mes + f"\n{g[i][0]}          {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value}  -  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value}\n          {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value}\n"

            else:

                mes = mes + f"\n{g[i][0]}          {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[0]}  -  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines()[len(f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines())-1]}\n          {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[1]}\n"
        bot.send_message(call.from_user.id, f"Расписание на {need_day(-3)}\n"+mes)#Отправляем










######### Доп функции там всякие


def need_day(n):
    now = datetime.now()  
    tomorrow = now + timedelta(days=n)
    return tomorrow.strftime("%d.%m")

def datetime_full(n):
    now = datetime.now()  
    tomorrow = now + timedelta(days=n)
    return tomorrow.strftime("%Y%m%d")


bot.polling()