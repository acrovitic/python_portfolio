import nltk
from nltk.stem.lancaster import LancasterStemmer
stemmer = LancasterStemmer()
import numpy as np
import tflearn
import tensorflow as tf
import win32com
from win32com import client
import datetime as dt
import traceback
import pickle
import sys
import string
import os
import glob
from functions_and_classes import *
import re
from dateutil.parser import parse
from collections import Counter
from IPython.display import display_html

path = "path/to/tensorflow/"

# load filer training data
data = pickle.load(open(path+"outlook_filer_dnn_training_data","rb"))
words = data['words_f']
classes = data['classes_f']
train_x = data['trainf_x']
train_y = data['trainf_y']

# initialize dnn shell based on originally trained dnn
tf.reset_default_graph()
net = tflearn.input_data(shape=[None, len(train_x[0])])
net = tflearn.fully_connected(net,8)
net = tflearn.fully_connected(net,8)
net = tflearn.fully_connected(net,len(train_y[0]), activation='softmax')
net = tflearn.regression(net)
dnn = tflearn.DNN(net, tensorboard_dir=path+'tflearn_logs/loaded_outlook_filer')

# load trained dnn into shell
dnn.load(path+"outlook_filer_dnn_model.tflearn")

# initialize outlook for the dnn to work in
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
main_folder = outlook.Folders[2]
inbox=main_folder.Folders["Inbox"]
date = dt.datetime.now().strftime("%m/%d/%Y")
exclude=['File','Needs Attention','TO BE POSTED','In Process']

for email in list(main_folder.Folders["Inbox"].Items.Restrict("[SentOn] < '{} 7:00 AM'".format(date))):
    if not email.Categories:
        continue
    else:
        e = email_obj(email)
        e.clean_attributes()
        e.get_features()
        if not any([i in email.Categories for i in exclude]):
            p_input = bow(e.features, words, show_details=False)
            predict = dnn.predict([p_input])
            p_output = classes[predict[0].argmax()]
            formatted_p_output = format_outlook_path(build_outlook_path(p_output),e.received)
            if "SARFs" in formatted_p_output:
                if "*" not in formatted_p_output.rsplit(".",1)[1]:
                    formatted_p_output = formatted_p_output.replace("['2018']","['*2018']")
            try:
                email.Move(eval(formatted_p_output))
                print(f_ai_text.format(e=e.subject,f=formatted_p_output))
            except:
                print("[F_AI]: <<ERROR>> Could not move {e}.\n======\n".format(e=e.subject),
                      f_ai_text.format(e=e.subject,f=formatted_p_output))
                continue
        if e.category == "File": #attachments that need to be saved and msg moved to M
                if e.attcount < 1:
                    atp = 'no attachment'
                else:
                    atp = e.attpath
                    try:
                        for att in email.Attachments:
                            if not os.path.exists(e.attpath):
                                os.makedirs(e.attpath)
                            att.SaveAsFile(e.attpath+att.FileName)
                    except:
                        print("[F_AI]: <<ERROR>> Could not file attachment(s) from {e}.\n------".format(e=e.subject))
                        traceback.print_exc()
                        continue
                try:
                    frmtd_loc = build_outlook_path(e.outlook) # outlook folder
                    m_loc = e.mdrive # mdrive email folder
                    email.SaveAs(m_loc+e.subject.replace(":","_")+".msg") # save to mdrive email folder
                    email.Move(eval(frmtd_loc)) # move to outlook folder
                    print(f_ai_text2.format(e=e.subject,f=frmtd_loc,m=m_loc,c='1.0',a=atp))
                except:
                    print("[F_AI]: <<ERROR>> Could not save to M; {e}.\n======\n".format(e=e.subject),f_ai_text2.format(e=e.subject,f=frmtd_loc,m=m_loc,c='1.0',a=atp))
                    traceback.print_exc()
                    continue
