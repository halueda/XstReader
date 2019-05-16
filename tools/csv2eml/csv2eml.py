#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os


# In[2]:


os.chdir('//172.17.129.234/d/gdrive/ueda-note/Outlookuファイル/recovered')


# In[3]:


#os.getcwd()


# In[26]:


import glob
targets = glob.glob('*.csv')


# In[5]:


#import pandas
#
#
## In[6]:
#
#
#import pandas as pd
#
#
## In[7]:
#
#
#file='01-family.csv'
#dir=file.replace('.csv','')
##(file,dir)
#
#
## In[8]:
#
#
#data = pd.read_csv(file)
#data
#
#
## In[9]:
#
#
#col=data.columns
#(col[col.str.contains('User')],col[col.str.contains('Header')],col[col.str.contains('To')],col[col.str.contains('Subject')],col[col.str.contains('Date')],col[col.str.contains('Time')],col[col.str.contains('to')])
#
#
## In[10]:
#
#
## NaN check
#data[data['TransportMessageHeaders'].isnull() ]
#
#
## In[11]:
#
#
#pd.set_option("display.max_rows", 200)
#pd.set_option('max_colwidth',3000)
#noheader0=data.loc[13,:].T
#noheader0
#
#
## In[12]:
#
#
#pd.set_option('max_colwidth',1000)
#noheader0[noheader0.str.contains('travelance').fillna(False)]
#
#
## In[13]:
#
#
#data.loc[10:10,('OriginalSubject', 'ClientSubmitTime','DisplayTo', 'DeferredDeliveryTime','InReplyToId')]
#
#
## In[14]:
#
#
#pd.set_option('max_colwidth',1000)
#data.loc[10:10,('SentRepresentingEmailAddress')]
#
#
## In[15]:
#
#
#data['alt_header'] = 'Subject: ' + data['Subject'].fillna('') + '; Date: ' + data['ClientSubmitTime'].fillna('') + '; To: ' + data['DisplayTo'].fillna('')
#data.loc[:,'alt_header']
#
#
## In[16]:
#
#
#pd.reset_option("display.max_rows")
#pd.reset_option("display.max_colwidth")
#
#
## In[17]:
#
#
#data0=data.fillna({'TransportMessageHeaders': data['alt_header']})
#
#
## In[18]:
#
#
#data1 = data0.loc[:,('TransportMessageHeaders','UserEntryId')]
#data1
#
#
## In[19]:
#
#
#os.makedirs(dir, exist_ok=True)
#
#
## In[20]:
#
#
#header = data.loc[0,'TransportMessageHeaders'].split(';')
#header
#
#
## In[21]:
#
#
import re
#header1 = []
#cur = header[0]
#for l in header[1:]:
#    if re.match(r" [a-zA-Z][-_a-zA-Z0-9]*:", l) :
#        header1.append( cur )
#        cur = l[1:]
#    elif re.match(r" [ \t]", l) :
#        header1.append( cur )
#        cur = l[1:]
#    else :
#        cur += ";"+l
#header1.append( cur )
#header1
#
#
## In[22]:
#
#
def reformat_header(header):
    header1 = []
    cur = header[0]
    for l in header[1:]:
        if re.match(r" [a-zA-Z][-_a-zA-Z0-9]*:", l) :
            header1.append( cur )
            cur = l[1:]
        elif re.match(r" [ \t]", l) :
            header1.append( cur )
            cur = l[1:]
        else :
            cur += ";"+l
    header1.append( cur )
    return header1
#
#reformat_header(header)
#
#
## In[23]:
#
#
#body = data1.loc[0,'UserEntryId'].split(';')
#body
#
#
## In[24]:
#
#
#index=0
#
#with open(os.path.join(dir, ('%03d.eml'%index)), 'wt', encoding='iso2022-jp') as fout:
#    print( "\n".join(reformat_header(data1.loc[index,'TransportMessageHeaders'].split(';'))), file=fout)
#    print( '', file=fout)
#    print( "\n".join(data1.loc[index,'UserEntryId'].split(';')), file=fout)
#
#
## In[25]:
#
#
#os.linesep="\n"
#for index in range(0, len(data1)):
#    with open(os.path.join(dir, ('%03d.eml'%index)), 'wt',encoding='iso2022-jp',errors='replace') as fout:
#        print( "\n".join(reformat_header(data1.loc[index,'TransportMessageHeaders'].split(';'))), file=fout)
#        print( '', file=fout)
#        print( "\n".join(data1.fillna({'UserEntryId':''}).loc[index,'UserEntryId'].split(';')), file=fout)
#
#
## # メモ
## - EntryIdの長さが短すぎる。contentではないようだ。
#
## In[ ]:
#
#
#
#
#
## In[ ]:
#
#
#
#
#
## In[ ]:
#
#
#
#
#
## In[ ]:


import pandas as pd


def csv2eml( f ):
    dir=f.replace('.csv','')

    data = pd.read_csv(f)

    data['alt_header'] = 'Subject: ' + data['Subject'].fillna('') + '; Date: ' + data['ClientSubmitTime'].fillna('') + '; To: ' + data['DisplayTo'].fillna('')

    data0=data.fillna({'TransportMessageHeaders': data['alt_header']})

    data1 = data0.loc[:,('TransportMessageHeaders','UserEntryId')]

    os.makedirs(dir, exist_ok=True)

    os.linesep="\n"
    for index in range(0, len(data1)):
        with open(os.path.join(dir, ('%04d.eml'%index)), 'wt',encoding='iso2022-jp',errors='replace') as fout:
            print( "\n".join(reformat_header(data1.loc[index,'TransportMessageHeaders'].split(';'))), file=fout)
            print( '', file=fout)
            print( "\n".join(data1.fillna({'UserEntryId':''}).loc[index,'UserEntryId'].split(';')), file=fout)

for f in targets:
    csv2eml(f)
