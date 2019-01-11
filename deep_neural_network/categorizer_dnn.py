import nltk
from nltk.stem.lancaster import LancasterStemmer
stemmer = LancasterStemmer()
import numpy as np
import tflearn
import tensorflow as tf
import random
import pandas as pd
import win32com
from win32com import client
import datetime as dt
import traceback
import pickle
import string
import os
import glob
from functions_and_classes import *
import re
from dateutil.parser import parse
from collections import Counter
from IPython.display import display_html
pd.set_option('display.max_colwidth', -1)
pd.set_option('display.max_rows', 500)
pd.options.mode.chained_assignment = None
path = "path/to/tensorflow/"

#load data from training
data = pickle.load(open(path+"outlook_categorizer_dnn_training_data","rb"))
words = data['words']
classes = data['classes']
train_x = data['train_x']
train_y = data['train_y']

# initialize dnn shell based on original trained model
tf.reset_default_graph()
net = tflearn.input_data(shape=[None, len(train_x[0])])
net = tflearn.fully_connected(net,8)
net = tflearn.fully_connected(net,8)
net = tflearn.fully_connected(net,len(train_y[0]), activation='softmax')
net = tflearn.regression(net)
dnn = tflearn.DNN(net, tensorboard_dir=path+'tflearn_logs/loaded_outlook_categorizer')

#load trained model into dnn shell
dnn.load(path+"outlook_categorizer_dnn_model.tflearn")

# initialize outlook for the dnn to work with
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
main_folder = outlook.Folders[2]
inbox=main_folder.Folders["Inbox"]
date = dt.datetime.now().strftime("%m/%d/%Y")
exclude=['File','Needs Attention','TO BE POSTED']
for email in list(main_folder.Folders["Inbox"].Items.Restrict("[SentOn] < '{} 7:00 AM'".format(date))):
    if not email.Categories:
        e = email_obj(email)
        e.clean_attributes()
        e.get_features()
        p_input = bow(e.features, words, show_details=False)
        predict = dnn.predict([p_input])
        p_output = classes[predict[0].argmax()]
        email.Categories = p_output
        email.Save()
        print(c_ai_text.format(e=e.subject,p=p_output,))
