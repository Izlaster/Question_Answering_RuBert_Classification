import config
import logging
import telebot
import os
import sys
import json
import time
from telebot import types
from requests.exceptions import ConnectionError, ReadTimeout
from openpyxl import load_workbook

# log
logging.basicConfig(level=logging.INFO)

# bot init
bot = telebot.TeleBot(config.TOKEN)

# making user class


class Session():
    def __init__(self, id, cell):
        self.id = id
        self.cell = cell
        self.time = time.perf_counter()

    def getId(self):
        return self.id


# import data from JSON
with open("position.json", "r") as read_it:
    data = json.load(read_it)

# Global variables
maxCell = int(data['Position']['maxCell'])
missedCells = data['Position']['missedCells']
allSessions = []

# keryboard init
categories = {"знакомство": 1,
              "МИCиC": 2,
              "поступление - перевод": 3,
              "общежитие": 4,
              "учебная деятельность": 5,
              "внеучебная деятельность": 6,
              "документы": 7,
              "заказ услуг": 8,
              "сайт": 9,
              "Мoсква": 10,
              "другие города": 11,
              "я": 12,
              "ИИ": 13,
              "работа": 14,
              "обязанности студента": 15,
              "финансы": 16,
              "фигня, удалить": 17,
              "сохранить результат": 18}
keyboard = []
for category in categories:
    keyboard.append(types.KeyboardButton(category))

# load excel file
workbook = load_workbook(filename="output.xlsx")

# open workbook
sheet = workbook.active


@bot.message_handler(commands=['start'])
def welcome(message):

    # keyboard
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    item1 = types.KeyboardButton("Начнем?)")

    markup.add(item1)

    bot.send_message(message.chat.id, "Добро пожаловать, {0.first_name}!\nЯ - <b>{1.first_name}</b>, бот созданный чтобы создать классификацию для лучшего* датасета)).".format(message.from_user, bot.get_me()),
                     parse_mode='html', reply_markup=markup)


@bot.message_handler(content_types=['text'])
def lalala(message):

    global maxCell

    if message.text == 'Начнем?)':

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Погнали!", reply_markup=markup)

        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'знакомство':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "1"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'МИCиC':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "2"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'поступление - перевод':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "3"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'общежитие':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "4"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'учебная деятельность':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "5"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'внеучебная деятельность':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "6"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'документы':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "7"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'заказ услуг':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "8"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'сайт':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "9"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'Мoсква':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "10"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'другие города':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "11"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'я':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "12"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'ИИ':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "13"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'работа':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "14"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'обязанности студента':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "15"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'финансы':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "16"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'фигня, удалить':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                if (time.perf_counter() - allSessions[i].time) <= 900:
                    sheet["B" + str(allSessions[i].cell)] = "!"
                    allSessions.pop(i)
                    break
                else:
                    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                    item1 = types.KeyboardButton("Начнем?)")
                    markup.add(item1)
                    bot.send_message(
                        message.chat.id, text="Ответ не засчитан", reply_markup=markup)
                    missedCells.append(allSessions[i].cell)
                    allSessions.pop(i)
                    break

        # keyboard
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(*keyboard)
        bot.send_message(message.chat.id, text="Следующее",
                         reply_markup=markup)
        if (len(missedCells) == 0):
            bot.send_message(
                message.chat.id, text=sheet["A" + str(maxCell)].value)
            allSessions.append(Session(message.from_user.id, maxCell))
            maxCell = maxCell + 1
        else:
            bot.send_message(
                message.chat.id, text=sheet["A" + str(missedCells[0])].value)
            allSessions.append(Session(message.from_user.id, missedCells[0]))
            missedCells.pop(0)

    elif message.text == 'сохранить результат':
        for i in range(0, len(allSessions)):
            if (allSessions[i].getId() == message.from_user.id):
                workbook.save('output.xlsx')
                markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                item1 = types.KeyboardButton("Начнем?)")
                markup.add(item1)
                bot.send_message(
                    message.chat.id, text="Сохраненно!", reply_markup=markup)
                missedCells.append(allSessions[i].cell)
                allSessions.pop(i)
                break

# RUN
try:
    bot.infinity_polling(timeout=10, long_polling_timeout=5)
except (ConnectionError, ReadTimeout) as e:
    data['Position']['maxCell'] = maxCell
    data['Position']['missedCells'] = missedCells
    with open('position.json', 'w') as file:
        json.dump(data, file)
    sys.stdout.flush()
    os.execv(sys.argv[0], sys.argv)
else:
    bot.infinity_polling(timeout=10, long_polling_timeout=5)
    data['Position']['maxCell'] = maxCell
    data['Position']['missedCells'] = missedCells
    with open('position.json', 'w') as file:
        json.dump(data, file)