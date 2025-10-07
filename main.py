import camelot
import openpyxl 
import telebot
from datetime import datetime, timedelta 
import json
import os
import urllib.request

token='8269738099:AAETqsa8WwNzhBfVH2zLay7_svsH_DLQDTc'
#8347380655:AAE56FocrVCTzAY39vc4QOo9Oz0IsZttcBw ориг
#8269738099:AAETqsa8WwNzhBfVH2zLay7_svsH_DLQDTc тест
bot=telebot.TeleBot(token)



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
        



        b1= telebot.types.InlineKeyboardButton(text=f"<--",callback_data=f">")
        b2= telebot.types.InlineKeyboardButton(text=f"Сегодня",callback_data=f"{datetime_full(0)}")
        b3= telebot.types.InlineKeyboardButton(text=f"-->",callback_data=f"<")

        b6= telebot.types.InlineKeyboardButton(text=f"{need_day(1)}",callback_data=f"{datetime_full(1)}")
        b7= telebot.types.InlineKeyboardButton(text=f"{need_day(2)}",callback_data=f"{datetime_full(2)}")
        b8= telebot.types.InlineKeyboardButton(text=f"{need_day(3)}",callback_data=f"{datetime_full(3)}")
        b9= telebot.types.InlineKeyboardButton(text=f"{need_day(-1)}",callback_data=f"{datetime_full(-1)}")
        b10= telebot.types.InlineKeyboardButton(text=f"{need_day(-2)}",callback_data=f"{datetime_full(-2)}")
        b11= telebot.types.InlineKeyboardButton(text=f"{need_day(-3)}",callback_data=f"{datetime_full(-3)}")
        
        #Вставляем в клавиатуру
        kb1.add(b6,b7,b8,b9,b10,b11,b1,b2,b3)

        bot.send_message(message.chat.id, "Привет", reply_markup=kb1) #Выводим клавиатуру


@bot.callback_query_handler(func=lambda call: True)
def callback_query(call):
    if call.data == f"<":

        bot.send_message(call.from_user.id, "Пока не работает, терпи.")

    elif call.data == f">":

        b6= telebot.types.InlineKeyboardButton(text=f"{need_day(1)}",callback_data=f"{datetime_full(1)}")







        bot.send_message(call.from_user.id, "Пока не работает, терпи.")

    else:
        send_mes(call)


#elif call.data == f"{need_day(-3)}":
    

def send_mes(call):
    url=f"https://gtnu.ru/wp-content/uploads/rasp/{call.data}.pdf"
    pdf=f'{call.data}.pdf'
    try:
        urllib.request.urlretrieve(url, pdf)
    except urllib.error.HTTPError as err:

        bot.send_message(call.from_user.id, f"Расписания нет. Отдыхай, разрешаю👍")
        return None

        
    

    desting = f'{call.data}.xlsx'

    a = camelot.read_pdf(pdf)
    os.remove(pdf)
    a[1].df.to_excel(desting)
    
    #открываем фаил exel
    fff=openpyxl.load_workbook(desting) 
    os.remove(f'{call.data}.xlsx') 
    f=fff.active

    #Ищем сколько всего пар
    g=[]
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
            None
            #mes = mes + f"\n{g[i][0]}          {f.cell(row=g[i][1][0],column=2).value}\n        ———\n"




        elif len(f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines()) == 3:

            mes = mes + f"\n{g[i][0]}       {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[0]}  —  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines()[len(f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines())-1]}\n          {f.cell(row=g[i][1][0]+2,column=g[i][1][1]+1).value.splitlines()[1]}  —  {f.cell(row=g[i][1][0]+2,column=g[i][1][1]+1).value.splitlines()[0]}\n          {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines()[1]}  —  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines()[0]}\n       {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[1]}\n"





        else:#              1                                 08:30 - 09:55                                             Тех.мех.                                                        134                                                                                                                                                                 Лекции                                                                                                                                                                                      

            mes = mes + f"\n{g[i][0]}       {f.cell(row=g[i][1][0],column=2).value}\n    {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[0]}  —  {f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines()[len(f.cell(row=g[i][1][0]+1,column=g[i][1][1]+1).value.splitlines())-1]}\n       {f.cell(row=g[i][1][0],column=g[i][1][1]+1).value.splitlines()[1]}\n"
    

    bot.send_message(call.from_user.id, f"Расписание на  {call.data[6:8:]}.{call.data[4:6:]}\n" + mes)#Отправляем
    













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