#****Importing Libraries****#

import numpy as np
import pandas as pd 
import string
import warnings
warnings.filterwarnings('ignore')
import re

#from textblob import TextBlob

import nltk
#nltk.download('punkt')
#nltk.download('wordnet')

import spacy

from nltk.corpus import stopwords

#Stemming
#from nltk.stem import PorterStemmer
#stemmer = PorterStemmer()

#Lemmetization
from nltk.stem import WordNetLemmatizer
lemmetizer = WordNetLemmatizer()

from flask import Flask, redirect, url_for, request,render_template

#Initializing spacy model
nlp = spacy.load('en_core_web_sm')

#Words that we do not require after tokenization in the search term and the title column as well
unwanted_words = ['Email','Sample','Letter','Template','email','sample','letter','template','to','To']

#Punctuations
#punctuations = '''()-[]{};:'"\,<>./?@#$%^&*_`~'''

#Stopwords
stop_words = stopwords.words('english')


#****Importing Dataset and data Preprocessing****#

#Reading the data set and preprocessing
data = pd.read_excel('templates.xls')
data = data.dropna()
data = data.reset_index(drop=True)

#Replacing 'Letter' with 'Email' || Replacing '-' with ' ' in the cat column
for i in range (0,len(data)):
    data.iloc[i,0] = data.iloc[i,0].replace('Letter','Email')
    data.iloc[i,0] = data.iloc[i,0].replace('letter','Email')
    data.iloc[i,1] = data.iloc[i,1].replace('Letter','Email')
    data.iloc[i,1] = data.iloc[i,1].replace('letter','Email')
    data.iloc[i,0] = data.iloc[i,0].replace('-',' ')

#Adding Search column which will contain tokens of title and arranging the columns in order    
data['Search_column'] = ''
data = data[["Category","Search_column", "Title", "Body"]]

#Removing the 'Download Related Samples...' from footer
for i in range (0,len(data)):
    txt = data.iloc[i,3]
    txt0 = txt.split('\nDownload ')[0]
    data.iloc[i,3] = txt0

#Removing address fields
for i in range (0,len(data)):
    list0 = data.Body[i].split('\n')
    list1 = []
    for string in list0:
        if len(string)> 50:
            list1.append(string)
    body = 'To: [Email id]\n\nSubject: [Subject]\n\n[Salutation]\n\n'
    for j in list1:
        body = body + j + '\n\n'
    body = body + '[Closure]\n\n'+'[Name/Signature]\n'+'[Designation]\n'+'[Organization]\n'
    body = body.replace('Letter','Email')
    body = body.replace('letter','email')
    #print(body)
    data.iloc[i,3] = body

#Preprocessing the title column
for i in range (0,len(data)):
    tex = str(data.Title[i])
    tex_p = ''
    #for char in tex:
    #    if (char not in punctuations):
    #        tex_p = tex_p+char
    tex_p = re.sub(r'[^\w\s]','',tex)
    tex_p = tex_p.strip()
    tex_a = nltk.word_tokenize(tex_p)
    tex_s=[]
    for word in tex_a:
        if (word not in stop_words):
            tex_s.append(word)
    tex_l = []
    for word in tex_s:
        tex_l.append(lemmetizer.lemmatize(word))
    tex_f = []
    for word in tex_l:
        if word not in unwanted_words:
            tex_f.append(word)
        else:
            pass
    data.iloc[i,1] = tex_f

#Named Entity Recognition (Replacing named entities with their tags)
for i in range(0,len(data)):
    s= data.Body[i]
    doc = nlp(s)
    newString = s
    for e in reversed(doc.ents): #reversed to not modify the offsets of other entities when substituting
        start = e.start_char
        end = start + len(e.text)
        newString = newString[:start] + '[' + e.label_ + ']' + newString[end:]
    data.Body[i] = newString
    

#****For future reference****#

df = pd.DataFrame(data.Category.value_counts())
df = df.reset_index()
df = df.rename(columns={'index':'Category','Category':'N_Letters'})


#****Deployment****#

app = Flask(__name__)

#app.static_folder = 'static'


@app.route('/')
def home_page():
    cats = df.Category.to_list()
    return render_template('index.html',cat_list = cats)



@app.route('/category',methods = ['POST','GET'])
def category():
    cats = df.Category.to_list()
    if request.method == 'POST':
        s = str(request.form['cat'])
        b = data[data.Category == s]
        title = b.Title.to_list()
    return render_template('cat_title.html',title_list=title,cat_list = cats)



@app.route('/category/title',methods = ['POST','GET'])
def cat_title():
    if request.method == 'POST':
        s = str(request.form['title'])
        b = data[data.Title == s]
        b = b.reset_index(drop=True)
        letter = 'Category: '+str(b.Category[0])+'\nTitle: '+str(b.Title[0])+'\n____________________________________________________________________________________________________________\n'+str(b.Body[0])
        letter = letter.split('\n')
        cats = df.Category.to_list()
        c = data[data.Category == b.Category[0]]
        title = c.Title.to_list()
    return render_template('print.html',title_list = title,cat_list = cats,letter=letter)


        
@app.route('/search_results',methods = ['POST','GET'])
def search():
    if request.method == 'POST':
        s = str(request.form['keywords'])
        text=''
        text = re.sub(r'[^\w\s]','',s)
        text = text.strip()
        text = text.title()
        text_a = nltk.word_tokenize(text)
        text_s=[]
        for word in text_a:
            if (word not in stop_words):
                text_s.append(word)
        text_l = []
        for word in text_s:
            if word == 'boss':
                text_l.append(word)
            else: 
                text_l.append(lemmetizer.lemmatize(word))
        text_f = []
        for word in text_l:
            if word not in unwanted_words:
                text_f.append(word)
            else:
                pass
        #Searching for keywords in the search column
        result_in = []
        for word in text_f:
            for i in range (0,len(data)):
                for j in range (0,len(data.iloc[i,1])):
                    if data.iloc[i,1][j] not in text_f:
                        pass
                    else:
                        result_in.append(int(i))
        result_in.sort(reverse=False)
        x = data.iloc[result_in]
        x['priority'] = np.NaN
        p = x.index.value_counts()
        x = x.reset_index()
        p = p.reset_index()
        p = p.rename(columns={0:'priority'},inplace=False)
        for i in range(0,len(x)):
            ind = x['index'][i] 
            for j in range(0,len(p)):
                if p['index'][j] == ind:
                    pr = p['priority'][j]
                else:
                    pass
            x['priority'][i] = pr
        x = x.sort_values(by=['priority'],ascending=False)
        x = x.drop_duplicates(subset=['index'])
        x = x.drop(['index','Search_column','priority'],axis=1)
        x = x.reset_index(drop=True)
        #letter = 'Category: '+str(x.Category[0])+'\nTitle: '+str(x.Title[0])+'\n____________________________________________________________________________________________________________\n'+str(x.Body[0])
        #letter = letter.split('\n')
        cats = df.Category.to_list()
        title_list = x.Title.to_list()
    return render_template('keywordsearch.html',cat_list = cats,title_list=title_list)



@app.route('/search_results/print',methods = ['POST','GET'])
def print_search_results():
    if request.method == 'POST':
        s = str(request.form['title'])
        b = data[data.Title == s]
        b = b.reset_index(drop=True)
        letter = 'Category: '+str(b.Category[0])+'\nTitle: '+str(b.Title[0])+'\n____________________________________________________________________________________________________________\n'+str(b.Body[0])
        letter = letter.split('\n')
        cats = df.Category.to_list()
        c = data[data.Category == b.Category[0]]
        title = c.Title.to_list()
    return render_template('searchresultprint.html',letter = letter,cat_list = cats,title_list=title)




if __name__ == '__main__':
    app.run()