import config
import logging
import telebot

import difflib

import pandas as pd
from conllu import parse
from nltk import DependencyGraph
from ufal.udpipe import Model, Pipeline
import re
from transformers import AutoTokenizer
import numpy as np
import random

from modeling import *

# log
logging.basicConfig(level=logging.INFO)

# bot init
bot = telebot.TeleBot(config.TOKEN)

# import main df
df = pd.read_excel('make_dataset/08042023_dataset_sort_change_tokenizer_fix.xlsx')

# tokenizer + model
tokenizer = BertTokenizer.from_pretrained(pre_trained_model_ckpt)
model = Model.load('model/russian-ud-2.0-170801.udpipe')
class_names = ['поступление - перевод', 'общежитие', 'учебная деятельность', 'внеучебная деятельность', 'документы', 'работа', 'финансы']
myModel = SentimentClassifier(len(class_names))
myModel.load_state_dict(torch.load('model/best_model_state.bin'))
myModel = myModel.to(device)

# make match
def mySort(s1, s2):
  matcher = difflib.SequenceMatcher(None, s1, s2)
  return matcher.ratio()

# subject object verb
def get_sov(sent):
    graph = DependencyGraph(tree_str=sent)
    sov = {}
    for triple in graph.triples():
        if triple:
            if triple[0][1] == 'VERB':
                sov[triple[0][0]] = {'subj':'','obj':''}
    for triple in graph.triples():
        if triple:
            if triple[1] == 'nsubj':
                if triple[0][1] == 'VERB':
                    sov[triple[0][0]]['subj']  = triple[2][0]
            if 'obj' in triple[1]:
                if triple[0][1] == 'VERB':
                    sov[triple[0][0]]['obj'] = triple[2][0]
    return sov

# token function
def tokenText(text):
    arrParsed = []
    longParsed = ''

    sent = str(text)
    sent = re.sub("\s\s+", ' ', sent)
    sent = sent.lower()
    sent = sent.strip()
    sent = sent.replace(u'\xa0', u' ')
    sentS = re.split(";|! |\?|\. ", sent)
    sentS = list(filter(None, sentS))

    for nSentS in sentS:
        nSentS = re.sub(r'[^\w\s]','', nSentS) 

        pipeline = Pipeline(model, 'tokenize', Pipeline.DEFAULT, Pipeline.DEFAULT, Pipeline.DEFAULT)
        parsed = pipeline.process(nSentS)

        # костыли для dependency graph
        parsed = '\n'.join([line for line in parsed.split('\n') if not line.startswith('#')])
        parsed = parsed.replace('\troot\t', '\tROOT\t')

        if (len(longParsed) < len(parsed)):
            longParsed = parsed

        arrParsed.append(parsed)
    arrParsed = list(filter(None, arrParsed))

    arrSov = []
    for nArrParsed in arrParsed:
        sov = get_sov(nArrParsed)
        arrSov.append(sov)

    bigSov = arrSov[0]
    if (len(arrSov) > 1):
        for i in range(1, len(arrSov)):
            bigSov = { **bigSov, ** arrSov[i] }

    tokenizedA = []
    if (bigSov == {}):
        arrGraph = []
        graph = DependencyGraph(tree_str=longParsed)
        for i in range(0, len(list(graph.triples()))):
            for j in range(0, len(list(graph.triples())[0])):
                if (list(graph.triples())[i][j][1] == 'NOUN'):
                    arrGraph.append(list(graph.triples())[i][j][0])
        arrGraph = set(arrGraph)
        if (len(arrGraph) == 0):
            tokenizedA.append('!');
        else:
            for nArrGraph in arrGraph:
                tokenizedA.append(tokenizer.encode(nArrGraph, add_special_tokens=False))
            tokenizedA = np.concatenate(tokenizedA, axis=0, out=None, dtype=None, casting="same_kind")
    else:
        for k, v in bigSov.items(): 
            tokenizedA.append(tokenizer.encode(k, add_special_tokens=False))
            tokenizedA.append(tokenizer.encode(v['subj'], add_special_tokens=False))
            tokenizedA.append(tokenizer.encode(v['obj'], add_special_tokens=False))
        tokenizedA = np.concatenate(tokenizedA, axis=0, out=None, dtype=None, casting="same_kind")

    return tokenizedA


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

    tToken = tokenText(text)
    tq = []
    index = 0
    probability = 0.00
    textQ = ''
    textA = ''
    questions = []

    for i in range(0, len(df)):
        if (probability <= mySort(tToken, eval(df['tq'][i])) and df['label'][i] == int(prediction)):
            tq = eval(df['tq'][i])
            index = i
            textQ = df['q'][i]
            textA = df['a'][i]
            probability = mySort(tToken, eval(df['tq'][i]))
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