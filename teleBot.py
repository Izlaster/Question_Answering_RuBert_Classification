import config
import logging
import telebot

import difflib

import pandas as pd
import re
from transformers import BertTokenizer

from pymorphy2 import MorphAnalyzer
from nltk.corpus import stopwords
import random

from modeling import *

# log
logging.basicConfig(level=logging.INFO)

# bot init
bot = telebot.TeleBot(config.TOKEN)

# import main df
df = pd.read_excel('make_dataset/10042023_dataset_sort_change_tokenizer_fix_fix.xlsx')

# tokenizer + model
tokenizer = BertTokenizer.from_pretrained(pre_trained_model_ckpt)
class_names = ['поступление - перевод', 'общежитие', 'учебная деятельность', 'внеучебная деятельность', 'документы', 'работа', 'финансы']
myModel = SentimentClassifier(len(class_names))
myModel.load_state_dict(torch.load('model/best_model_state.bin'))
myModel = myModel.to(device)

# make match
def mySort(s1, s2):
  matcher = difflib.SequenceMatcher(None, s1, s2)
  return matcher.ratio()

patterns = "[A-Za-z0-9!#$%&'()*+,./:;<=>?@[\]^_`{|}~—\"\-]+"
stopwords_ru = stopwords.words("russian")
morph = MorphAnalyzer()

def lemmatize(doc):
    doc = re.sub(patterns, ' ', doc)
    tokens = []
    for token in doc.split():
        if token and token not in stopwords_ru:
            token = token.strip()
            token = morph.normal_forms(token)[0]
            
            tokens.append(token)
    return tokenizer.encode(" ".join(tokens), add_special_tokens=False)


@bot.message_handler(commands=['start'])
def welcome(message):
    bot.send_message(message.chat.id, "Добро пожаловать, {0.first_name}!\nЯ - <b>{1.first_name}</b>, бот создан в качестве дипломной работы. Напиши свой вопрос".format(message.from_user, bot.get_me()),
                     parse_mode='html')


@bot.message_handler(content_types=['text'])
def lalala(message):
    text = message.text
    encoded_review = tokenizer.encode_plus(text, max_length=512, add_special_tokens=True, return_token_type_ids=False, pad_to_max_length=True, return_attention_mask=True,
                                        truncation=True, return_tensors='pt')
    input_ids = encoded_review['input_ids'].to(device)
    attention_mask=encoded_review['attention_mask'].to(device)
    output = myModel(input_ids, attention_mask)
    _,prediction = torch.max(output, dim=1)

    tToken = lemmatize(text)
    tq = []
    index = 0
    probability = 0.00
    textQ = ''
    textA = ''
    questions = []

    for i in range(0, len(df)):
        if (probability <= mySort(tToken, eval(str(df['tq_fix'][i]))) and df['label'][i] == int(prediction)):
            tq = eval(str(df['tq_fix'][i]))
            index = i
            textQ = df['q'][i]
            textA = df['a'][i]
            probability = mySort(tToken, eval(str(df['tq_fix'][i])))
            questions.append([tq, index, textQ, textA, probability])

    max_last_elem = max([que[-1] for que in questions])
    max_lists = [lst for lst in questions if lst[-1] == max_last_elem]
    rand = random.choice(max_lists)

    bot.send_message(message.chat.id, f'{text} - Текст основного вопроса' + '\n' +
                     f'{class_names[prediction]} - ИИ класс' + '\n' + 
                     f'{tToken} - Токенизированный текст основного вопроса' + '\n' +
                     f'{rand[4]*100}% - Процент сходства с найденным вопросом' + '\n' +
                     f'{rand[1]} - Номер найденного вопроса' + '\n' +
                     f'{rand[2]} - Текст найденного вопроса' + '\n' +
                     f'{rand[0]} - Токенизированный найденный вопрос' + '\n' +
                     f'{rand[3]} - Ответ к найденному вопросу')
    
# RUN
bot.infinity_polling(timeout=10, long_polling_timeout=5)
