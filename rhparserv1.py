# -*- coding: utf-8 -*-
"""
RH Parser release 1.98

Created on Wed Jan 15 14:10:41 2020
Last Update Mar 11 2020

@author: 122188 Shivaji Basu, TCS Tech BU Labs
(C) Tata Consultancy Services Ltd.

Reverse Heuristic Parser

Extracts text and matches with mater data attrbutes to identify relevant element

The algorithm does not need positional mapping of OCR or stuctured data
"""
import os
import json
import codecs
import time
import random
import datetime
import sys
import csv
#from get_tokens import get_tokens
#from logprint import logprint as logwrite
import os
import json

try:
    from pdfminer.pdfinterp import PDFResourceManager,PDFPageInterpreter
    from pdfminer.converter import TextConverter
    from pdfminer.layout import LAParams
    from pdfminer.pdfpage import PDFPage
    from io import StringIO
    #from io import BytesIO
    
except:
    print("Error: 140003 probably pdfminer not installed")
#Loading Configuration

logfile=""

##############################################################################
#Fetching config from config.json

try:
    currentdir=os.path.dirname(os.path.realpath(__file__))
    with open(currentdir+'\\config.json',encoding='utf-8') as txt:
        sr=txt.read()
        config = json.loads(sr)
        txt.close()
except:
    #raise
    print("Error 11001: Configurations could not be loaded. Check for config.json file.")

###############################################################################
##Check license

from entitlement import check_entitlement
if not check_entitlement(file=config['license_key_path']):
    sys.exit()

################################################################################

try:
    logfile=config["logfile"]
except:
    print("Error 140001: logfile configuration not found")

try:
    size=os.path.getsize(logfile)
    if size>10000000:
        os.rremove(logfile)
except:
    print("Log Size check failed")

##########################################################################
#This module is used to log issues in place of print

def logwrite(message,annotation='',prmlist=None, debugmode=False):
    global logfile
    try:
        
        timestamp=None
        params=None
        
        timestamp="'time':'"+str(time.strftime("%H:%M:%S"))+"'" 
        if not prmlist is None:
            params="'prms':'"+prmlist+"'"
        else:
            params=""
        
        log="{"+",".join(["'message':'"+str(message)+"'",timestamp, params])+"},"
        
        file=codecs.open(logfile,'a+','utf-16')
        file.write(log+'\r\n')
        print(message, annotation)
        file.close()
        if debugmode:        
                raise
    except:
        #raise
        print(message)
        file.close()
##########################################################################

try:
    
    import pandas as pd 
    import collections
    from itertools import combinations    
    import re
    import numpy as np

except:
    #raise
    print("Error: 140003 probably pandas not installed")

try:
    #from statistics import median
    import win32com.client
    import win32com
except:
    logwrite("Error: 140004 splash screen library missing")

"""try:
    from PIL import Image
    #pip install wand
    #http://docs.wand-py.org/en/latest/guide/install.html#install-imagemagick-on-windows
    #pip install MagickWand
    #pip install pillow
    #pip install pytesseract
    from wand.image import Image as Imgpdf
    #conda install -c brianjmcguirk pyocr
    '''
    import pyocr
    import pyocr.builders
    '''
except:
    
    logwrite("Error 140005: One or more of the libraries PIL, wand, pyocr, pytesseract missing. If not please check if tesseract executable exist in local folder")"""

#Splash screen
try:
     
    ps=0.5
    appname="Reverse Heuristics Parser"
    time.sleep(ps)
    
    Error=False
    
    print(""  ) 
    print(""  )
    #logwrite("*****************************************************************"  )
    #print(cfg.renderText(appname))

    logwrite("Reverse Heuristics Parser release 1.2 Beta" )
    logwrite("Copyright Tata Consultancy Services Ltd. 2020"  )
    logwrite("*****************************************************************"  )
    print(""  )
except:

    print("Reverse Heuristics Parser release 1.2 Beta" )
    print("Copyright Tata Consultancy Services Ltd. 2020"  )
    print("*"*45)

#########fetching file argument################################################
try:
    files=[]
    if len(sys.argv)>1:
        file=sys.argv[1]
        files.append(file)

except:
    print("Error: 140006 File not found.")
    
###############################################################################
#SINGLETONS
    
valuecounter={}#keeps occurance of same numeric values
rhstopwords=[]#stopwords. Can be forther enhanced at runtime in add_to_stopword function
assigned=[]#keeps tokens already asigned
prospects=[]#values that are counterparts in product match
table_start=0
table_end=99999
#INITIALIZE SINGLETONS
def initialize_case():    
    global valuecounter,rhstopwords,assigned,prospects,table_start,table_end
    try:
        valuecounter={}#keeps occurance of same numeric values
        rhstopwords=config["stopwords"] #stopwords. Can be forther enhanced at runtime in add_to_stopword function
        assigned=[]#keeps tokens already asigned
        prospects=[]#values that are counterparts in product match
        #indices for table start and end to be ppopulated later, would help us match better
        table_start=9999
        table_end=0
    except:
        logwrite("Error 140007 COuld not initialize new case")

###############################################################################
#stopwords. Can be forther enhanced at runtime in add_to_stopword function
try:
    rhstopwords=config["stopwords"]    
    
except:
    logwrite("Error 140008: problem loading stopwords from config.json")
    
try:

    from nltk.corpus import stopwords
    from nltk.corpus import wordnet
    from nltk import pos_tag
    nltk_stpwds=list(stopwords.words(config["language"]))
    stopwords=list(set(rhstopwords+nltk_stpwds))
except:
    
    logwrite("Error 140009: problem loading stopwords from NLTK. Ignored")
    

#############################################################################
#Delealing with comma and dot separators
def apply_culture(token_string,culture_dict={}):
    
    try:
        placeholders=[]
        d=json.loads(culture_dict)
            
    except:
        
        logwrite("Alert 140050: Culture format incorrect or missing.")
        return token_string
    
    try:
        s=token_string
        
            
        while len(placeholders)<len(d):
            placeholders.append("$"+str(random.randint(100000,999999))+"$")
            placeholders=list(set(placeholders))
        
        i=0
        for item in d:
            s=s.replace(item,placeholders[i])
            i+=1
        i=0
        for item in d:
            s=s.replace(placeholders[i],d[item])
            i+=1
        

        return s
    except:
        #raise
        logwrite("Alert  140055: Culture fuction error")
###############################################################################
#dealing with content in mail body instead of pdf

def getmsgs(folder, senderemail):
    try:
        messages = folder.Items
        a=len(messages)
        if a>0:
            for msg in messages:
                try:
                    sender = msg.SenderEmailAddress
                    if sender == senderemail:
                        if msg.Unread==True:
                            txt = str(msg.Subject) +"\n "+ msg.Body
                            return txt     
                except:
                    logwrite("Error 140011: Could ot extract mail content")
                    return False
    except:
        logwrite("Error 140012 Could ot extract mail content")
        return False
        
###################################################################################
#getting content from mail body instead of df or excel
def getUnreadMsg(foldername, sendername):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts;
    except:
        logwrite("Error 140013: Outlook client not open. Ingoring mail check." )
        return False

    try:
        for account in accounts:
            global inbox
            inbox = outlook.Folders(account.DeliveryStore.DisplayName)
            folders = inbox.Folders
        
        
            for folder in folders:
                if str(folder)==foldername:
                    s=getmsgs(folder, sendername)
                    return s    
    
    except:
        logwrite("Error 140014 Could not fatch mail from outlook.")
        return False

###############################################################################
#any stopword logic. y can be a function that returns a list
def add_to_stopwords(x):
    try:
        global rhstopwords
        if type(x)==type([]):
            rhstopwords=list(set(rhstopwords+x))
        
        
    except:
        logwrite("Error 140025 Adding to stopwords failed")
            
##############################################################################

#fetch numeric tokens        toke

def get_values(tokens,culture_dict=None):
    #print("temp",tokens)
    try:
        alltokens=[]
        lsamt=[]
        for i in range(len(tokens)):
                ls=[]
                ls.append((tokens[i][0],tokens[i][1]))
                m=apply_culture(tokens[i][0],culture_dict)
                ls.append((m,tokens[i][1]))
                m=tokens[i][0].replace(",","")
                ls.append((m,tokens[i][1]))
                m=tokens[i][0].replace(",",".")
                ls.append((m,tokens[i][1]))
                ls.append((m,tokens[i][1]))
                m=tokens[i][0].replace(".","")
                ls.append((m,tokens[i][1]))
                m=tokens[i][0].replace(".",",")
                ls.append((m,tokens[i][1]))
                ls=list(set(ls))
                for j in ls:
                    alltokens.append(j)
                    
    except:
        #raise
        logwrite("Error 140037 Culture processing failed")
        alltokens=tokens
        
    try:
        for i in alltokens:        
                try:
                    
                    x=float(i[0])
                    if x-int(x)<0.000001:
                        lsamt.append((int(x),i[1]))
                        
                    else:
                        lsamt.append((x,i[1]))
                except:
                    z="float conversion does  not apply"
        #global valuecounter            
        ###############Add valucounter , which is to be used in range_validator 
        global valuecounter
        lsvc=[]
        for i in lsamt:
            lsvc.append(i[0])
        valuecounter=collections.Counter(lsvc)
        
        return list(lsamt)
                
    except:
        logwrite("Error 140038: Converting to numeric values failed.")
###############################################################################

def get_keyrecords_dp(tokens, filepath, matchcol, match_based_on='CPN', match_type='Strict',ignorecol=None,connection_type='csv',connection_details=None):
    # This function retrives the first key value and attribute from the excel where the key resides in excel. Typically used to identify company name using Vat no. as key
    try:
        lines = []
        if connection_type=='csv':
            if match_based_on == 'CPN':
                df = pd.read_csv(filepath)
            elif match_based_on == 'MPN':
                import mpn_matches
                df = mpn_matches.get_MPN_dataset(filepath)
        elif connection_type=='sql':
            import pyodbc
            conn_string="Driver={ODBC Driver 17 for SQL Server};Server=01HW1810791\\SQLEXPRESS;Database=TestDB;Uid=sa;Pwd=Tbusxo@123;"
            query="SELECT [PARTNO],[PRICEMIN],[PRICEMAX],[QUANTITYMIN],[QUANTITYMAX],[DELTA_BACL],[DELTA_FORWARD],[TI_PART_NO] FROM [TestDB].[dbo].[line]"
            #conn_string=connection_details['conn_string']
            #query=connection_detials['conn_string']
            cnxn = pyodbc.connect(conn_string)
            cursor = cnxn.cursor()
            df = pd.read_sql(query, cnxn)
            print(df)
           
        tokens = sorted(tokens, key=lambda x: x[1])
        #print(match_type)
        
        if match_type == 'Strict' or match_type == 'Broken':
            df_tokens=pd.DataFrame(tokens)#,columns=["PARTNO","Index"])
            #df_tokens = df_tokens.astype(str)
            df_tokens['Tokens_Cleaned']=df_tokens[df_tokens.columns[0]].str.replace('[^0-9a-zA-Z.]+', '', regex=True)
            new_df=df.copy()
            new_df = new_df.astype(str)
            new_df['PARTNO_Cleaned']=new_df[new_df.columns[matchcol]].str.replace('[^0-9a-zA-Z.]+', '', regex=True)
            match_records=pd.merge(df_tokens,new_df,left_on="Tokens_Cleaned",right_on="PARTNO_Cleaned",how="inner")
            for record in match_records.values.tolist():
                dict = {}
                dict["key"] = list(record)[df_tokens.shape[1]+matchcol]
                dict["token"] = list(record)[0]
                dict["index"] = list(record)[1]
                dict["range"] = list(record)[df_tokens.shape[1]:df_tokens.shape[1]+matchcol] + list(record)[df_tokens.shape[1]+matchcol+1:-1]
                tok_index = next((index for (index, d) in enumerate(lines) if d["index"] == list(record)[1]), None)
                if tok_index is not None and dict["index"] == lines[tok_index]["index"]:
                    if (len(list(record)[df_tokens.shape[1]+matchcol]) > len(lines[tok_index]["key"])):
                        print("I am here")
                        lines.pop(tok_index)
                        lines.append(dict)
                else:
                    lines.append(dict)
                    
        if match_type == 'Broken':
            for token in tokens:
                dict = {}
                record = None
                # removing special characters in match
                tk = re.sub('[^0-9a-zA-Z.]+', '', token[0])
                #print("df2")
                if tk.strip() != '':
                    if df[(df[df.columns[matchcol]].str.replace('[^0-9a-zA-Z.]+', '', regex=True).str.startswith(tk)) & (df[df.columns[matchcol]].str.replace('[^0-9a-zA-Z.]+', '', regex=True).str.len() > len(tk))].shape[0] > 0:
                        match_record = df[(df[df.columns[matchcol]].str.replace('[^0-9a-zA-Z.]+', '', regex=True).str.startswith(tk)) & (df[df.columns[matchcol]].str.replace('[^0-9a-zA-Z.]+', '', regex=True).str.len() > len(tk))]
                        for eachrecord in range(0, match_record.shape[0]):
                            print(eachrecord)
                            remaining_token = re.sub('[^0-9a-zA-Z.]+', '', match_record.iloc[eachrecord].tolist()[matchcol])[len(tk):]
                            if remaining_token != '':
                                newtokens = tokens[token[1]:token[1]+20]
                                for newtoken in newtokens:
                                    print(newtoken)
                                    if re.sub('[^0-9a-zA-Z.]+', '', remaining_token) == re.sub('[^0-9a-zA-Z.]+', '', newtoken[0]):
                                        record = match_record.iloc[eachrecord].tolist()
                                        break
                            if record != None:
                                break
                    if record != None:
                        dict["key"] = list(record)[matchcol]
                        dict["token"] = list(token)[0]
                        dict["index"] = list(token)[1]
                        dict["range"] = list(record)[0:matchcol] +  list(record)[matchcol+1:]
                        # lines.append(dict)
                        tok_index = next((index for (index, d) in enumerate(lines) if d["index"] == list(token)[1]), None)
                        if tok_index is not None and dict["index"] == lines[tok_index]["index"]:
                            if (len(list(record)[0]) > len(lines[tok_index]["key"])):
                                print("I am inside")
                                lines.pop(tok_index)
                                lines.append(dict)
                        else:
                            lines.append(dict)
        
        if ignorecol is not None:
            ignore_list=[]
            index_to_be_poped=[]
            for count,item in enumerate(lines):
                if ignorecol > matchcol:
                    ignore_data=item["range"][ignorecol-1]
                else:
                    ignore_data=item["range"][ignorecol]
                if ignore_data.strip() != "":
                    if item["key"] not in ignore_data.split("|"):
                        ignore_list.extend(ignore_data.split("|"))
            print(ignore_list)    
            index_to_be_poped=[count for count,item in enumerate(lines) if item["key"] in ignore_list]
            print(index_to_be_poped)
            index_to_be_poped.sort(reverse=True)
            for ind in index_to_be_poped:
                lines.pop(ind)
        
        #print(lines)
        return(lines)
    except Exception as e:
        print(str(e))
        #raise
        print("Error 140039A: file record reading failed")
        # file.close

def get_keyrecords(tokens, filepath, matchcol, match_based_on='CPN', match_type='Strict',ignorecol=None,connection_type='csv',connection_details=None):
    # This function retrives the first key value and attribute from the excel where the key resides in excel. Typically used to identify company name using Vat no. as key
    try:
        lines = []
        
        ##checking the connection type and loading the content of line to a dataframe
        if connection_type=='sql':
            try:
                import pyodbc
                conn_string=connection_details['conn_string']
                query_type = match_based_on.lower()+"_query"
                query=connection_details[query_type]
                cnxn = pyodbc.connect(conn_string)
                cursor = cnxn.cursor()
                df = pd.read_sql(query, cnxn)
                cursor.close()
                cnxn.close()
            except:
                print("Error 140039A.1: Probably pyodbc is not installed or check the connection details for data query")
        elif connection_type=='csv':
            #print(filepath)
            #print(match_based_on)
            df = pd.read_csv(filepath[match_based_on.lower()])
            
        
        ##Cleaning Dataframe
        df = df.fillna(value=np.nan)
        #print(df)
        df.replace(np.nan,"", inplace=True)
                
        '''
        ###checking whether it is an MPN match
        if match_based_on == 'MPN':
            import mpn_matches
            df = mpn_matches.get_MPN_dataset(df)'''
            
           
        tokens = sorted(tokens, key=lambda x: x[1])
        #print(match_type)
        
        if match_type == 'Strict' or match_type == 'Broken':
            df_tokens=pd.DataFrame(tokens)#,columns=["PARTNO","Index"])
            #df_tokens = df_tokens.astype(str)
            df_tokens['Tokens_Cleaned']=df_tokens[df_tokens.columns[0]].str.replace('[^0-9a-zA-Z.]+', '', regex=True)
            new_df=df.copy()
            new_df = new_df.astype(str)
            new_df['PARTNO_Cleaned']=new_df[new_df.columns[matchcol]].str.replace('[^0-9a-zA-Z.]+', '', regex=True)
            match_records=pd.merge(df_tokens,new_df,left_on="Tokens_Cleaned",right_on="PARTNO_Cleaned",how="inner")
            for record in match_records.values.tolist():
                dict = {}
                if list(record)[df_tokens.shape[1]+matchcol].strip() == "":
                    continue
                dict["key"] = list(record)[df_tokens.shape[1]+matchcol]
                dict["token"] = list(record)[0]
                dict["index"] = list(record)[1]
                dict["range"] = list(record)[df_tokens.shape[1]:df_tokens.shape[1]+matchcol] + list(record)[df_tokens.shape[1]+matchcol+1:-1]
                tok_index = next((index for (index, d) in enumerate(lines) if d["index"] == list(record)[1]), None)
                if tok_index is not None and dict["index"] == lines[tok_index]["index"]:
                    if (len(list(record)[df_tokens.shape[1]+matchcol]) > len(lines[tok_index]["key"])):
                        lines.pop(tok_index)
                        lines.append(dict)
                else:
                    lines.append(dict)
                    
        if match_type == 'Broken':
            for token in tokens:
                dict = {}
                record = None
                # removing special characters in match
                tk = re.sub('[^0-9a-zA-Z.]+', '', token[0])
                #print("df2")
                if tk.strip() != '':
                    if df[(df[df.columns[matchcol]].str.replace('[^0-9a-zA-Z.]+', '', regex=True).str.startswith(tk)) & (df[df.columns[matchcol]].str.replace('[^0-9a-zA-Z.]+', '', regex=True).str.len() > len(tk))].shape[0] > 0:
                        match_record = df[(df[df.columns[matchcol]].str.replace('[^0-9a-zA-Z.]+', '', regex=True).str.startswith(tk)) & (df[df.columns[matchcol]].str.replace('[^0-9a-zA-Z.]+', '', regex=True).str.len() > len(tk))]
                        for eachrecord in range(0, match_record.shape[0]):
                            #print(eachrecord)
                            remaining_token = re.sub('[^0-9a-zA-Z.]+', '', match_record.iloc[eachrecord].tolist()[matchcol])[len(tk):]
                            if remaining_token != '':
                                newtokens = tokens[token[1]:token[1]+20]
                                for newtoken in newtokens:
                                    #print(newtoken)
                                    if re.sub('[^0-9a-zA-Z.]+', '', remaining_token) == re.sub('[^0-9a-zA-Z.]+', '', newtoken[0]):
                                        record = match_record.iloc[eachrecord].tolist()
                                        break
                            if record != None:
                                break
                    if record != None:
                        dict["key"] = list(record)[matchcol]
                        dict["token"] = list(token)[0]
                        dict["index"] = list(token)[1]
                        dict["range"] = list(record)[0:matchcol] +  list(record)[matchcol+1:]
                        # lines.append(dict)
                        tok_index = next((index for (index, d) in enumerate(lines) if d["index"] == list(token)[1]), None)
                        if tok_index is not None and dict["index"] == lines[tok_index]["index"]:
                            if (len(list(record)[0]) > len(lines[tok_index]["key"])):
                                lines.pop(tok_index)
                                lines.append(dict)
                        else:
                            lines.append(dict)
        
        if ignorecol is not None:
            ignore_list=[]
            index_to_be_poped=[]
            for count,item in enumerate(lines):
                if ignorecol > matchcol:
                    ignore_data=item["range"][ignorecol-1]
                else:
                    ignore_data=item["range"][ignorecol]
                if ignore_data.strip() != "":
                    if item["key"] not in ignore_data.split("|"):
                        ignore_list.extend(ignore_data.split("|"))
            #print(ignore_list)    
            index_to_be_poped=[count for count,item in enumerate(lines) if item["key"] in ignore_list]
            #print(index_to_be_poped)
            index_to_be_poped.sort(reverse=True)
            #print(index_to_be_poped)
            for ind in index_to_be_poped:
                lines.pop(ind)
        
        #print(lines)
        return(lines)
    except Exception as e:
        print(str(e))
        #raise
        print("Error 140039A: file record reading failed")
        # file.close

###############################################################################

def get_token_atIndexRange(tokens,startindex, endindex):
    try:
        value_list=[]
        value_list=(i for i in tokens if startindex <=i[1]<=endindex)
        return(value_list)
    except:
        logwrite("Error 16800: Cannot fetch tokens in specified range")

###############################################################################
def get_value_atIndexRange(tokens,startindex, endindex, culture=None,rounding=6,unit_price_incrementer=None):
    try:
        
        tokens=get_token_atIndexRange(tokens,startindex, endindex)
        #unit_price_incrementer=2
        value_list=[]
        tokens_sliced=[]
        for i in tokens:
            #print(i)
            
            s1=i[0].split("/")
            
            #if len(s1)>2:
            #    continue
            
            s1 = [x for x in s1 if x != '']
            #print(s1)
            try:
                if len(s1)==2 and (int(float((s1[1]))) in [1,10,100,1000,10000]):
                    unit_price_incrementer=0
                    for s in s1:        
                        tokens_sliced.append((s,i[1]))
                else:
                    tokens_sliced.append(i)
            except:
                tokens_sliced.append(i)
                continue

        for i in tokens_sliced:
            
            if not culture is None:
                s=apply_culture(i[0], culture)
                #print(s)
                s=s.split(" ")
                for l in s:
                    if startindex <=i[1]<=endindex:
                        try:
                            v= (round(float(l),rounding),i[1])
                            #print(v)
                            value_list.append(v)
                        except:
                            continue
            else:
                s=i[0]
                if startindex <=i[1]<=endindex:
                    try:
                        
                        v= (round(float(s),rounding),i[1])
                        #print(v)
                        value_list.append(v)
    
                    except:
                        continue
    except Exception as e:
        #print(str(e))
        logwrite("Error 16801: Extracting value between indices failed. Check culture record in header.")
    #handling quantitites expresses as rates say 60.00/1000
    try:
        #unit_price_incrementer=5
        
        if unit_price_incrementer is None:
            unit_price_incrementer=0
        
        derived_list=[]
        for v in value_list:
            for k in value_list:
                if k!=v:
                    #if k[1]==v[1]:
                    if k[1]==v[1]+unit_price_incrementer:
                        if k[0] in [1,10,100,1000,10000]:
                            derived_list.append((round(v[0]/k[0],rounding),v[1]))
                
                
        value_list=value_list+derived_list
   
    except Exception as e:
        print(str(e))
        logwrite("Error 16802: Extracting value between indices failed. Check culture record in header.")
    
    return(value_list)
        
###############################################################################

def get_keyrecord(tokens, filepath, matchcol):

    # This function retrives the first key value and attribute from the excel where the key resides in excel. Typically used to identify company name using Vat no. as key

    try:
        with open(filepath, 'r', encoding='utf-8') as file:

            records = list(csv.reader(file, dialect='excel'))

            for record in records[1:]:
                hrecord = re.sub('[^0-9a-zA-Z|]+', '', record[matchcol])
                hnames = hrecord.split('|')
                #print(hnames)
                for token in tokens:
                    tk = re.sub('[^0-9a-zA-Z]+', '', token[0])
                    #tk=token[0]
                    if tk in hnames:
                        #print(tk)
                        hnames.remove(tk)
                    if len(hnames) == 0:
                        file.close
                        return (record[matchcol], record, token[1])

    except Exception as e:
        print(str(e))
        logwrite("Error 140039B: Record file reading failed")
        file.close

##########################################################################
def crossvalidation(values,hint):
    return None

##########################################################################
def get_seedmatch(tokens,lowerbound, upperbound, rounding=5, hint=1,function=crossvalidation,stoplist=[],allowzero=True):
    pass
    
    #This function retrives the all values that fall between the range
    #About hints
#############################################################################
def get_token_by_index(tokens, idx):
    try:
        for i in tokens:
            if i[1]==idx:
                return i
        return None

    except:
        logwrite("Error 16009: Error in parameters")

###############################################################################
             
def morestops(y):
    try:
        sl=[]
        for i in range(1,y+1):
            for j in range(1,i):
                sl.append(str(j) +' of '+str(i))
        return sl       
    except:
        return []

###############################################################################

def pdftotext(path):
    try:
        manager = PDFResourceManager()
        retstr = StringIO() #BytesIO()
        #layout = LAParams(all_texts=True)
        layout = LAParams(line_margin=0.01,all_texts=True)
        device = TextConverter(manager, retstr,laparams=layout)
        filepath = open(path, 'rb') 
        interpreter = PDFPageInterpreter(manager, device)
        
        
        for page in PDFPage.get_pages(filepath, check_extractable=True):
            #print(page,type(page))
            interpreter.process_page(page)
        text = retstr.getvalue()
        filepath.close()
        device.close()
        retstr.close()
        
        #RE = re.compile(u'([^\x00-\x7F])+', re.UNICODE)
        #text = RE.sub('', text)
        #print(nochinese)
        ext= str(text)
        return ext
    except Exception as e:
        #raise
        #print(str(e))
        logwrite("Error 140016: Probably pdf file does not exist")

################################################

def exceltotext(path):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path,data_only=True)
        sheet = wb.active
        tokens=list(sheet.values)
        wb.close
        return tokens
    except:


        logwrite("Error 140017: Probably Excel file does not exist")

######################################################

def htmltotext(path):
    try:
        import email
        from bs4 import BeautifulSoup
    except:
        logwrite("Error 140018A: Probably BeautifulSoup is not installed use pip insatll beautifulsoup4")
    try:
        with open(path,'r') as fp:
            message = email.message_from_file(fp)
            for part in message.walk():
                if (part.get_content_type() == "text/html" or part.get_content_type() == "text/plain"):
                    soup = BeautifulSoup(part.get_payload(decode=True),'lxml')
                    #print(soup.find("body").get_text('\n'))
                    return soup.find("body").get_text('\n')
                    #return (soup.body()[0])
        fp.close()
    except UnicodeDecodeError as e:
        with open(path,encoding='utf-8') as fp:
            message = email.message_from_file(fp)
            for part in message.walk():
                if (part.get_content_type() == "text/html" or part.get_content_type() == "text/plain"):
                    soup = BeautifulSoup(part.get_payload(decode=True),'lxml')
                    #print(str(soup.find("body")).encode().decode('unicode_escape'))
                    return str(soup.find("body").get_text('\n')).encode().decode('unicode_escape')
                    #return (soup.body()[0])
        fp.close()
    except Exception as e:

        #print(str(e))
        logwrite("Error 140018B: Probably html file does not exist")

############################################################################

def Texttotext(path):
    try:
        #import string
        try:
            f=open(path,'r',encoding="utf-8")
            s=f.read()
        except:
            f.close()
            f=open(path,'r',encoding="big5")
            s=f.read()
        #print(s)
        #logwrite(path, s)
        f.close()
        #printable = set(string.printable)
        #''.join(filter(lambda x: x in printable, s))
        return(s)
        
    
    except Exception as e:
        #raise
        print(str(e))
        logwrite("Error 140018: Probably mail file does not exist")

############################################################################

def Image_to_text(Imagefile, lang='eng', numeric_only=False,blob=True, convert_to_blackwhite=True):
    img=None
    return []
    #pyocr deprepcated
    #removing temp image file
    #logwrite(Imagefile)
    try:
        os.remove('tmp.png')
    except:
        logwrite("")
    #Checking Tesseract and Loanguage tools
    try:
        tools = pyocr.get_available_tools()
        tool = tools[0]
        langs = tool.get_available_languages()
        lang1 = langs[0]
        #logwrite(langs)
    
    except:
        
        logwrite("Error 140019: Language or tesseract not installed")
        return False
    #If the image is in PDF
    try:
        if Imagefile[-4:].upper()=='.pdf'.upper():
            #"This software may use ImageMagic if same is available in installed lication and does not distrubute th same otherwise. The library is attributed to ImageMagick Studio LLC.")
            with Imgpdf(filename=Imagefile, resolution=300) as img:
                img.compression_quality = 99
                
                img.save(filename='tmp.png')
                img=Image.open('tmp.png')
                
    except:
        
        logwrite("Error 140020: Could not convert PDF image")
        return False

    #Coverting Image into Black and while for better OCR
    try:
        if convert_to_blackwhite:    
            if img==None:
                img=Image.open(Imagefile)
            img=img.convert('L')
            
    except:
        
        logwrite("Error 140021: Image file not found or may be open")
        return False    
    #to get the digits only
    try:
        if numeric_only:
            digits = tool.image_to_string(img,lang=lang,builder=pyocr.tesseract.DigitBuilder())
            
            return digits
    except:
     
        logwrite("Error 140022: Could not read numbers")
        return False
   
    if blob:
        #to get the whole text using pyOCR
        try:
            txt = tool.image_to_string(img,lang=lang,builder=pyocr.builders.TextBuilder())
            return txt
        except:
            logwrite("Error 140023: Image could to be converted to text blob")
            return False            

    else:
        #to list of lines
        try:
            
            line_and_word_boxes = tool.image_to_string(img, lang=lang,builder=pyocr.builders.LineBoxBuilder())
            line=[]
            for i in line_and_word_boxes:
                line.append(i.content)          
            return line
  
        except:
            
            logwrite("Error 140024: Image could to be converted list of lines")
            return False

#######################################################################################################

def pdftotext_pdfplumber(path):
    try:
        import pdfplumber
        with pdfplumber.open(path) as p:
            page=p.pages
            q=len(page)
            fulltext=""
            for i in range(q):
                r=page[i]
                textdata=r.extract_text()
                #print(textdata)
                fulltext=fulltext+"\n"+str(textdata)

        return fulltext
    except Exception as e:
        # raise
        print(str(e))
        logwrite("Error 140026: Probably pdf file does not exist")

#######################################################################################################

def get_between_tokens(tokens, start_index, last_index):
    tokenlist = []
    for token in tokens:
        if token[1] >= start_index and token[1] < last_index:
            tokenlist.append(token[0])
    return tokenlist
    
#######################################################################################################
#Coverts any file to tokens list

def get_tokens(path,process_as_image=False,pdf_reader=None):
    global config
    global rhstopwords
    ls=[]
    ext=''
    #rhstopwords=config['stopwords']
    #extracting mail body
    try:
        logwrite('='*45)
        logwrite('Fetching config')
        stoplist=config["stoplist"]
        #stoplist.extend(morestops(1,10))
        
        replacelist=config["replacelist"]
        #replacelist.extend=len(morestops(1,10))*['']

        logwrite('Processing '+ path)
    except:
        logwrite("Error 140026: Error fetching stoplist. Check config and path")
        
    #extracting mail body
    try:
        if path[-4:].upper()=='.txt'.upper():
            logwrite('Processing txt file')
            ext =Texttotext(path)
            
    except:
        
        logwrite("Error 140027: mail text could not be parsed")

    #extracting pdf
    try:
        #print(path)
        if path[-4:].upper()=='.pdf'.upper():
            try:
                logwrite('Processing PDF in native text.')
                #print("pdf")
                if pdf_reader is None or pdf_reader=="pdfplumber":
                    #print("in pdfplumber")
                    ext = pdftotext_pdfplumber(path)   
                elif pdf_reader=="pdfminer":
                    ext = pdftotext(path)
                elif pdf_reader=="camelot":
                    try:
                        import camelot_extraction as ce
                        #if sequence_tokens is None:
                        sequence_tokens = get_tokens(path,pdf_reader="pdfplumber")
                        #print(sequence_tokens)
                        ext = ce.camelot_extraction(path,sequence_tokens)
                    except Exception as e:
                        logwrite(str(e))
                        logwrite("Error 140028.a : pdf could not be parsed as camelot")
                #print(ext)
                length=len(ext)
                #print(length)
            
            except:
                #raise
                logwrite("Error 140028: pdf data could not be parsed as text. Will treat it as image")
        
        #extracting pdf as image
            if length<15:
                try:
                    logwrite('Processing PDF as image')
                    #changed >>ext = ext+' '+ Image_to_text(Imagefile=path,lang=config['language'][0:3],numeric_only=False,blob=True,convert_to_blackwhite=True)
                    ext = Image_to_text(Imagefile=path,lang=config['language'][0:3],numeric_only=False,blob=True,convert_to_blackwhite=True)
                    #ext = ext+' '+ Image_to_text(Imagefile=path,lang=config['language'][0:3],numeric_only=True,blob=True,convert_to_blackwhite=True)
                    #logwrite(ext)
                except:
                    logwrite("Error 140029: pdf image could not be parsed")
                
    except:
        #raise
        logwrite("Error 140030: PDF could not be found")

    #extracting image files
    try:
        if path[-4:].upper() in ['.PNG','.JPG','.GIF','JPEG','TIFF']:
            try:
                logwrite('Processing image file')
                ext = ext+' '+ Image_to_text(Imagefile=path,lang=config['language'][0:3],numeric_only=False,blob=True,convert_to_blackwhite=True)
                #change>>ext = ext+' '+ Image_to_text(Imagefile=path,lang=config['language'][0:3],numeric_only=True,blob=True,convert_to_blackwhite=True)
            
            except:
                
                logwrite("Error 140031: Image could not be parsed. should be of type "+",".join(['.PNG','.JPG','.GIF','JPEG','TIFF']))
    except:
        logwrite("Error 140032: PDF could not be found")

    #extracting excel
    try:
        if path[-5:].upper()=='.xlsx'.upper():
            logwrite('Processing excel file')
            tokens =exceltotext(path)
            tokensraw=[]
            for i in tokens:
                for element in i:
                    
                    if not element is None:
                        tokensraw.append(str(element))
                        
            excelext=' '.join(tokensraw)  
            ext=ext+' '+excelext
            
    except:
        
        logwrite("Error 140033: Excel could not be parsed")
    
    try:
        if path[-4:].upper()=='.mht'.upper() or path[-4:].upper()=='.htm'.upper() or path[-5:].upper()=='.html'.upper():
            logwrite('Processing html file')
            ext =str(htmltotext(path))
            #print(ext)
    except:
        
        logwrite("Error 140035: html could not be parsed")
        #replacing stoplist

    try:
        
        ext=" ".join(ext.split())
        #print(ext)
        for i in morestops(10):
            
            ext.replace(i,' ')
        #print(ext)
        
        for i in range(len(stoplist)):
            #print(stoplist[i])
            ext=ext.replace(stoplist[i],replacelist[i])
        #ext=ext.replace("："," ")    
            
        
        ls=list(ext.split(' '))
    except Exception as e:
        print(str(e))
        logwrite("Error 140034 Stoplist format error.Check config file")
        
    #removing stopwords
    for i in rhstopwords:
        try:
            ls.remove(i)
        except:
            continue #If stopword alreadr removed, it would continue to next
    #print(ls)    
    #cleaning number format 1,200.00 vs 1.200,00
    try:
        ls2,ls3=[],[]
        for i in ls:
            if not i=='':
                ls2.append(i)
    except:
        logwrite("Error 140035: space elimination failed.")
        #cleaning the problem of comma used as decimels and vice versa

    #building index
    LsWithOrder=[]
    
    ls3=ls2 #disabling commawork
    
    for i in range(len(ls3)):
        try:
            LsWithOrder.append((ls3[i],i))
        except:
            logwrite("Error 140036A: value "+ str(ls3[i])+"  could not be processed")
            continue
    #print(LsWithOrder, "temp")    
    
    
    return LsWithOrder #ls3 #Not to be converted to distints

############################################################################
    #date identifier
def get_date(tokens,start_index,end_index, delta_back=365,delta_forward=90,format="y/m/d"):
    
    try:
        
        tokens2=[]
        now=datetime.date.today()
        d1=now - datetime.timedelta(days=delta_back)
        d2=now + datetime.timedelta(days=delta_forward)
        
        #print(d1)
        #print(d2)
        for i in tokens:
            if start_index<=i[1]<=end_index:tokens2.append(i)
        
        #print(tokens2)
        f=format.split("/")
        #print(f)
        
        if f.index('m')<f.index('d'):
            monthfirst=True
        else:
            monthfirst=False
            
        #not used so far
        monthindex={"JAN":1,"FEB":2,"MAR":3,"APR":4,"MAY":5,"JUN":6,"JUL":7,"AUG":8,"SEP":9,"OCT":10,"NOV":11,"DEC":12}
        #"MAY":5
        splitters=['.','-','/','\\','−']
        
        dates=[]
        for count,t in enumerate(tokens2):
            #print(t)
            parts=[]
            #if len(str(t[0]))!=10:logwrite(f'Alert: data format for {t} not acceptable. Do padding to pass arguments')
            if str(t[0]).upper() in monthindex.keys():
                parts=[tokens2[count-1][0],tokens2[count][0],tokens2[count+1][0]]
            elif len(str(t[0]))<=11:
                #print(str(t[0]))
                for s in splitters:
                    parts=str(t[0]).split(s)
                    if len(parts)==3:
                        break
            if len(parts)==3:
                #print("###############################")
                #print(t)
                #print("###############################")
                yidx=None
                for p in parts:
                    if len(p)==4:
                        yidx=parts.index(p)
                    elif len(p)==3 and p.isalpha() and (p.upper() in monthindex.keys()):
                        parts[parts.index(p)]=monthindex[p.upper()]
                #print(yidx)
                if yidx is None:
                    yidx=f.index('y')
                    parts[yidx]="20"+parts[yidx]
                    #print(parts[yidx])
                idx=[0,1,2]
                idx.remove(yidx)
                if monthfirst:
                    try:
                        #print(int(parts[idx[0]]),int(parts[idx[1]]),int(parts[yidx]))
                        dt=datetime.date(year=int(parts[yidx]),month=int(parts[idx[0]]),day=int(parts[idx[1]]))
                        
                        #print(dt)
                        if (d1<=dt<=d2):
                            #print(dt)
                            dates.append((str(dt),t[1])) 
                            #print("dates:")
                            #print(dates)
                    except:
                        pass
                else:                                
                    try:
                        #print(int(parts[idx[0]]),int(parts[idx[1]]),int(parts[yidx]))
                        dt=datetime.date(year=int(parts[yidx]),month=int(parts[idx[1]]),day=int(parts[idx[0]]))
                        #print(dt)
                        if (d1<=dt<=d2):dates.append((str(dt),t[1])) 
                        
                    except:
                        #raise
                        pass
            
        seen=set()
        dates = [(a, b) for a, b in dates 
             if not (a in seen or seen.add(a))]
        if len(dates)==1:
            return (dates[0],1)
    
        if len(dates)==0:
            return None
        if len(dates)>1:
            #print(dates)
            return (dates[0],round(1/len(dates),2))
    
    except Exception as e :
        #raise
        print(str(e))
        logwrite("Error 16989 Date resolution failed")
        return None

##########################################################################
def get_rangematch(tokens,lowerbound, upperbound, rounding=6, hint=1,function=crossvalidation,stoplist=[],allowzero=True, rangekey=None, culture=None,ignore_values=[]):
    logwrite(f"performing range match between values {lowerbound}  and {upperbound} with hint {hint} and rangekey {rangekey}")
    #This function retrives the all values that fall between the range
    #About hints
    #hint 0 is range match
    #hint 1 is specical stats
    #hint <10 is product match test, will be tried in all cases
    #hint 4 is summation match test , ill be done if product match fails
    global assigned, valuecounter, prospects
    #print("tbr",tokens)
    matched_tokens=[]
    product_matches=[]
    accuracy=-1
    try:
        
        tokens=list(filter(lambda x: x[0] not in stoplist,tokens)) #removing stoplist, Do not use Set operator here

    except:
        logwrite("Error 16001 Could not remove stoplist. check config file")
        
    if allowzero==False:
        tokens=list(filter(lambda x: x[0] !=0,list(tokens))) 
    '''
    tokens2=[]
    for i in tokens:
        s=apply_culture(i[0],culture)
        tokens2.append((s,i[1]))
    tokens=tokens2
    '''
    for token in tokens:
        #print(token)
        try:
            if hint==0:
                #print(round(float(lowerbound),rounding),float(token),round(float(upperbound),rounding))
                if round(float(lowerbound),rounding)<=float(token[0])<=round(float(upperbound),rounding):
                    #print(token)
                    matched_tokens.append(token)
                    
                          
            if hint==1:
                #Define a statistical variable
                testvariable =float(token[0])
                matched_tokens.append(token)
            
        except Exception as e:
            #print(str(e))
            logwrite("Error 140040: "+str(token) +"has been wrongly classified as value")

    try:
        #This part ensures values allready identified for other fields and not used as candidates
        #values=list(set(matched_tokens)-set(assigned))#
        #print(matched_tokens)
        values=list(filter(lambda x: (x[0] not in assigned) and (x[0] not in ignore_values),matched_tokens)) 
        #print(values)
        seen=set()  
        #seen=set()  
        values = [(a, b) for a, b in values 
             if not (a in seen or seen.add(a))]
        #print(values[0],"rangekey")
        
        if len(values)>1:
            #print(values)
            d=len(values)
            return (values[0],round(1/d,2))
        else:
            return (values[0],1)
        
    except Exception as e:
        #raise
        #print(str(e))
        value=None
        accuracy=-1
        logwrite("Error 140041: Not Values values exist")
    
    return (value,accuracy)
    
################################################################################
def find_tok_btwwords(identity,name_attribute,tokens):  
    """
    This function will read identity list and then
    list will have dictionaries as elements
    now each dictionary will have iden1 iden2 and element
    element is the search word and our function 
    will search for the element if it is present between iden1 and iden2 
    it returns the header name and value that needs to be assigned to header name
    """
    try:
        attr_name_val=[]                                                                     
        for ide in range(len(identity)):                                  
            iden=identity[ide]                         
            newname_attribute=name_attribute[:-1]+str(ide+1) 
            searchword=(iden["elements"]).split("|")                                    
            shiptoadd=get_betweentoken(tokens,iden["ident1"],iden["ident2"])
            if shiptoadd[0] is not None:
                for sw in searchword:
                    if sw.lower() in shiptoadd[0][0].lower():
                        acuity_shipto_Identifier=sw
                        break
                attr_name_val.append((newname_attribute,acuity_shipto_Identifier))            
            elif shiptoadd[0] is None:
                attr_name_val.append((newname_attribute,None))
    except Exception as e:
        raise
        print(str(e))
        logwrite("Error 07012021 find tok between words failed")#change 
    return attr_name_val

def putout(lines,path):
        
    try:
        np.savetxt(path, lines, delimiter=",", fmt='%s')
        logwrite("\nSaved to File: "+path+"\n")

        df=pd.DataFrame(lines)
        return df
    except:
        logwrite("ERROR 140045: Failed to wite to results file. Check if output.csv is already open")
        
################################################################################
def add_data(line,value):
    if not value in None:
        line.append(value)
    else:
        line.append(None)
    
    return line

################################################################################
def product_match(value_tokens,lowerbound,upperbound, rangekey, decimals=False, value_assigned=[],rounding=5,minimal_qty=0):
    
    logwrite(f"Validating hint 6 - Product Match with  lowerbound {lowerbound} upperbound {upperbound}  , decimals {decimals} , rangekey {rangekey[0]}")
    try:
        remaining_tokens=[]
        product_match=[]
        valuesets=[]#should be the product value 
        for i in value_tokens:
                if i[0] not in value_assigned and i[0]!=0 and i[1]!=rangekey[1]:
                    remaining_tokens.append(i)
        
        for i in remaining_tokens:
            valuesets.append(i[0])

        #print(valuesets)
        
        for i in remaining_tokens:
            ls=valuesets
            if not (decimals == False and float(i[0]).is_integer()==False):
                try:
                    
                    if round(float(lowerbound),rounding)<=i[0]<=round(float(upperbound),rounding) and i[0]>minimal_qty:
                        
                        ls.remove(i[0])
                        #print(i, rangekey[0], "range")
                        
                        #print(ls)
                        #print(product)
                        #print(type(i[0]),type(rangekey[0]), "token")
                        
                        product=i[0]*rangekey[0]
                        #print(product)
                        #rest_tokens=remaining_tokens.remove(i)

                        product_token=get_identitytoken(remaining_tokens,round(product,rounding))
                        #print(product_token)
                        if product_token is not None:
                            if product_token[0] in ls and product_token[1]!=rangekey[1]:
                                product_match.append(i)

                except Exception as e:
                    #print(str(e))
                    #raise
                    logwrite("Error 167000 Produc Match failed. Check config file value formats.")
                    
        seen=set()
        product_match = [(a, b) for a, b in product_match 
             if not (a in seen or seen.add(a))]            
        if len(product_match)==0:
            return (None,-1)
        
        if len(product_match)==1:
            return (product_match[0],1)
        
        if len(product_match)>1:
            #print(product_match)
            return (product_match[0],round(1/len(product_match),2))
        
    except Exception as e:
        #raise
        #print(str(e))
        logwrite("Error 16020 Product Match failed.")
            
#############################################################################

def get_identitytoken(tokens,identity):
    try:
        tokens=list(tokens)
        try:
            ids=identity.split(" ")
        except:
            ids=[identity]
        count=len(ids)
        #print(count)
        for item in tokens:
            if type(item[0])==type("str"):
                v=' '.join(i[0] for i in tokens[tokens.index(item):tokens.index(item)+len(ids)])
                #print(v)
                if v==identity:
                    #print(tokens[item[1]+len(ids)])
                    #print(tokens[tokens.index(item)+len(ids)-1])
                    return tokens[tokens.index(item)+len(ids)-1]
            else:
                #print(identity)
                if item[0]==ids[0]:
                    return item
        return None
    except Exception as e:
        print(str(e))
        logwrite("Error 16021 Cannot find the identity token")
##############################################################################
    
        
def get_neighbourhoodmatch(identity,tokens,startindex,endindex,regex=None,nei_type=None):
    try:
        #print("inside function")
        #print(identity)
        if identity !=  '':
            identity_token=get_identitytoken(tokens,identity)
        else:
            return (None,None)
        #print(identity_token)
        
        values=[]
        if identity_token is not None:
            items=get_token_atIndexRange(tokens,identity_token[1]+startindex,identity_token[1]+endindex)
            
            #if hint=="a":
            i=0
            
            value=''
            index=[]
            Flag=False
            inistr=''
            #print(regex)
            if nei_type=="regex":
                return get_regexmatch(items,None,None,regex)
            else:
                r=re.compile(regex)
                
                for item in items:
                    #print(item)
                    if r.match(item[0]):
                        #print(item)
                        if nei_type=="tel":
                            if len(value)+len(item[0])<20:
                                value=value+inistr+item[0]
                                index.append(item[1])
                                inistr=' '
                                Flag=True
                            else:
                                if len(value)>=5:
                                    values.append((value,index))
                                
                                value=''
                                index=[]
                                break
                        else:
                            #if r.match(item[0]):
                            value=value+inistr+item[0]
                            index.append(item[1])
                            inistr=' '
                            Flag=True
                    elif Flag:
                        if (nei_type=="tel" and len(value)>=5) or nei_type!="tel":
                            values.append((value,index))
                        
                        value=''
                        index=[]
                        break
                    else:
                        continue        
                try:
                    if (nei_type=="tel" and len(value)>=5) or nei_type!="tel" and value[0]!='':    
                        values.append((value,index))
                except:
                    pass
        if len(values)==0:
            return (None,-1)
        elif len(values)==1:
            return (values[0],1)
        else:
            #print(values)
            return (values[0],round((1/len(values)),2))
            
    except Exception as e:
        print(str(e))
        logwrite("Error 16022: Failed to fetch neighbourhood match")

#############################################################################

def get_regexmatch(tokens,startindex,endindex,regex,re_list=None,avoid_list=[]):
    try:
        #print(startindex)
        #print(endindex)
        #print(regex)
        if str(regex) is None:
            return (None,-1)
        elif str(regex).strip()=='':
            return (None,-1)
        if endindex is not None and startindex is not None:
            tokens=get_token_atIndexRange(tokens,startindex,endindex)
        #else:
        #    tokens=tokens
        #avoid_list=["500000"]
        #print(avoid_list)
        tokens=[token for token in tokens if token[0] not in avoid_list]
        #print(l)    
        #print(tokens)
        final_match=[]
        #print(regex)
        #print(list)
        match_tokens=[]
        r=re.compile(str(regex))
        for item in tokens:
            for i in r.findall(item[0]):
                match_tokens.append((i,item[1]))
        #match_tokens = [item for item in tokens if r.match(item[0])]
        #print(match_tokens)
        if re_list is not None:
            final_match=[match for match in match_tokens if match[0] in re_list]
        else:
            final_match=match_tokens
        #print(final_match)
        #remove duplicates
        seen=set()
        final_match = [(a, b) for a, b in final_match 
             if not (a in seen or seen.add(a))] 
        if len(final_match)==0:
            return (None,-1)
        elif len(final_match)==1:
            return (final_match[0],1)
        elif len(final_match)>1:
            #print(final_match)
            return (final_match[0],round((1/len(final_match)),2))
    except Exception as e:
        print(str(e))
        logwrite("Error 16023:Regex Macth failed")

###########################################################################

def get_distancetoken(tokens,identifier_token,distance):
    try:
        #print(identifier_token)
        tokens=list(tokens)
        if identifier_token is None:
            return (None,-1)
        elif identifier_token.strip()=='':
            return (None,-1)
        else:
            #print("#####")
            #print(get_identitytoken(tokens,identifier_token))
            distok=get_token_by_index(tokens,get_identitytoken(tokens,identifier_token)[1]+distance)
            #print(distok)
            if distok is not None:
                return (distok,1)
            else:
                return (None,-1)
            
    except Exception as e:
        print(str(e))
        logwrite("Error 16024:distancetoken Macth failed")
        return (None,-1)
        
###########################################################################

def get_betweentoken(tokens,identifier_token1,identifier_token2):
    try:
        print(identifier_token1)
        print(identifier_token2)
        print(get_identitytoken(tokens,identifier_token1))
        print(get_identitytoken(tokens,identifier_token2))
        #get_identitytoken(tokens,identifier_token2)
        if (identifier_token1 is None) or (identifier_token2 is None):
            return (None,-1)
        elif identifier_token1.strip()=='' or identifier_token2.strip()=='':
            return (None,-1)
        elif (get_identitytoken(tokens,identifier_token1) is None) or (get_identitytoken(tokens,identifier_token2) is None):
            #print("both r none")
            return(None,-1)
        elif get_identitytoken(tokens,identifier_token1)[1]+1==get_identitytoken(tokens,identifier_token2)[1]:
            return(None,1) 
        else:
            items=get_token_atIndexRange(tokens,get_identitytoken(tokens,identifier_token1)[1]+1,get_identitytoken(tokens,identifier_token2)[1]-1)
            #print(items)
            values=""
            index=[]
            for item in items:
                values=values+" "+item[0]
                index.append(item[1])
            #distok=get_token_by_index(tokens,get_identitytoken(tokens,identifier_token)[1]+distance)
            #print(distok)
            if values.strip() !='':
                return ((values,index),1)
            else:
                return (None,-1)
        
    except Exception as e:
        logwrite(str(e))
        logwrite("Error 16025:Betweentoken Macth failed")
        return (None,-1)
               
def get_brokentokens(tokens,startindex,endindex,match_token):
    try:
        if endindex is not None and startindex is not None:
            tokens=get_token_atIndexRange(tokens,startindex,endindex)
        #else:
        #    tokens=tokens
        #print(tokens)
        match_tokens=str(match_token).split("|")
        final_match=[]
        for count,t in enumerate(tokens):
            for match_tok in match_tokens:
                if t[0]==match_tok:
                    final_match.append(t)
                    continue
                if t[0]==match_tok[0:len(t[0])]:
                    remining_token=match_tok[len(t[0]):]
                    #print(remining_token)
                    match=[(match_tok,t[1]) for right_t in tokens if (right_t[0]==remining_token and right_t[1]>t[1])]
                    #print(match)
                    if len(match)>0:
                        final_match.append(match[0])
                    match=[]
        seen=set()
        final_match = [(a, b) for a, b in final_match 
             if not (a in seen or seen.add(a))] 
        #print(final_match)
        if len(final_match)==0:
            return (None,-1)
        elif len(final_match)==1:
            return (final_match[0],1)
        elif len(final_match)>1:
            #print(final_match)
            return (final_match[0],round((1/len(final_match)),2))
                    
    except Exception as e:
        logwrite(str(e))
        logwrite("Error 16026:Broken token Macth failed")
        return (None,-1)
        
#####################################################################################  
def itemno_check(itemno, adhoc_value, index, tokens):
    try:
        if itemno is not None:
            return itemno
        else:
            res = get_token_by_index(tokens, index-1)[0]
            return res
            # j[q["name"]] = rh.get_token_by_index(
            #     tok, dict['lines'][k]['index']-1)[0]
    except Exception as e:
        print(str(e))
        print("Error 17006")

def get_header_configuration_data(record):
    try:
        global config
        config_dict={}
        config_dict["record"]=record
        config_dict["culture"]=record[1][config["header"]["culture_column"]]
        config_dict["presicion"]=record[1][config["header"]["Amount_Precision"]]
        config_dict["date_format"]=record[1][config.get("header", None).get("date_format_column",None)]
        config_dict["start_date_delta"]=record[1][config["header"]["start_date_delta_column"]]
        config_dict["end_date_delta"]=record[1][config["header"]["end_date_delta_column"]]
        config_dict["linestart_buffer"]=record[1][config.get("header", None).get("linestart_buffer",None)]
        config_dict["unit_price_incrementer"]=record[1][config.get("header", None).get("unit_price_divider_position",None)]
        config_dict["Part_match_on"]=record[1][config.get("header", None).get("Part_match_on",'CPN')]
        if config_dict["Part_match_on"].strip() == '':
            config_dict["Part_match_on"]='CPN'
        config_dict["Part_match_type"]=record[1][config.get("header", None).get("Part_match_type",'Strict')]
        if config_dict["Part_match_type"].strip() == '':
            config_dict["Part_match_type"]='Strict'
        
        try:
            if config_dict["unit_price_incrementer"] == '':
                config_dict["unit_price_incrementer"]=None
            else:
                config_dict["unit_price_incrementer"]=int(config_dict["unit_price_incrementer"])
        except:
            logwrite("Could not capture unit_price_incrementer,Probably number was not provided")
           
        #print(record)
        try:
            config_dict["pdf_reader"]=json.loads(record[1][config.get("header", None).get("adhoc_config",None)])["pdf_reader"]  
            #pdf_reader="pdfminer"
            #print(pdf_reader)
        except Exception as e:
            #print(str(e))
            config_dict["pdf_reader"]=None
               
        if config_dict["linestart_buffer"] == '':
            config_dict["linestart_buffer"]=0
        config_dict["lineend_buffer"]=record[1][config["header"]["lineend_buffer"]]
        if config_dict["lineend_buffer"] == '':
            config_dict["lineend_buffer"]=0
    except Exception as e:
        logwrite("Error 16027: Unable to get configuration data")
    return config_dict

#####################################################################################  
def get_line_index_table(line_index_table):
    try:
        line_index_table=sorted(line_index_table, key = lambda x: x[1])
        
        if len(line_index_table)>1:
            curr_diff=line_index_table[1][1]-line_index_table[0][1]
        else:
            curr_diff=40 # max how many numeric tokens expected in a line
        
        diff=[]
        for i in range(len(line_index_table)):
            if i+1<len(line_index_table) and i<2:
                diff.append(line_index_table[i+1][1]-line_index_table[i][1])
                line_index_table[i].append(line_index_table[i+1][1]-1) 
            elif i+1<len(line_index_table):
                
                curr_diff=    line_index_table[i+1][1]-line_index_table[i][1]
                #print(curr_diff)
                if curr_diff-max(diff)>30:
                    line_index_table[i].append(line_index_table[i][1]+max(diff)+7)
                else:
                    line_index_table[i].append(line_index_table[i+1][1]-1)
                    diff.append(curr_diff)
            else:
                #line_index_table[i].append(line_index_table[i][1]+curr_diff)
                #print("last line")
                if len(diff)>0:
                    if max(diff)>30:
                        line_index_table[i].append(line_index_table[i][1]+max(diff)+7) 
                    else:
                        #print(line_index_table[i][1]+40)
                        line_index_table[i].append(line_index_table[i][1]+40)
                else:
                    line_index_table[i].append(line_index_table[i][1]+40)
    except:
        logwrite("Error in line_index_table mapping")
    return line_index_table
    
#####################################################################################
def header_extraction(tokens,config_dict,ls,header):
    try:
        global config,table_start,table_end
        linelist=[]
        for i in ls:
            if i.get("type","header")=="line":
                linelist.append(i)
                continue
            if i["range"] is not None:
                if i["range"]["hint"]==7:
                    try:
                        index_start=0
                        end_index=int(table_start)
                        lowerbound_delta =int(float(config_dict["record"][1][i["range"]["start_date_delta_column"]]))
                        upperbound_delta=int(float(config_dict["record"][1][i["range"]["end_date_delta_column"]]))
                        date_tokens=tokens
                        if i["range"]["identity_column"] is not None:
                            identity=config_dict["record"][1][i["range"]["identity_column"]]
                            if identity.strip()!='':
                                #print("identity")
                                identity_token=get_identitytoken(tokens,identity)
                                if identity_token is not None:
                                    date_tokens=get_token_atIndexRange(tokens,identity_token[1]-1,identity_token[1]+10)
                        date=get_date(date_tokens,start_index=index_start,end_index=end_index,delta_back =lowerbound_delta, delta_forward=upperbound_delta,format=config_dict["date_format"])
                        
                        if date!= None:
                            header[i["name"]]=date
                            #value_assigned.append(j[k["name"]][0])
                    except Exception as e:
                        #raise
                        print(str(e))
                        logwrite("Error 17901 Error in Data extraction")
                        
                if i["range"]["hint"]==8:
                    try:
                        
                        index_start=i["range"]["search_range_from"]
                        end_index=i["range"]["search_range_upto"]
                        identity=config_dict["record"][1][i["range"]["identity_column"]]
                        nei_type=i["range"]["type"]
                        if i["range"]["regex_column"] is not None:
                            regex=config_dict["record"][1][i["range"]["regex_column"]]
                        else:
                            regex=i["range"]["regex"]
                        
                        header[i["name"]]=get_neighbourhoodmatch(identity,tokens,index_start,end_index,regex,nei_type)
                        #print(header[i["name"]])
                        
                        #if header[i["name"]]!= None:
                            #value_assigned.append(header[i["name"]][0])
                    except Exception as e:
                        #raise
                        #print(str(e))
                        logwrite("Error 17902 Error in Neighorhood extraction")
                
                if i["range"]["hint"]==9:
                    try:
                        index_start=0
                        if i["range"]["search_range"]=="total":
                            end_index=None
                        elif i["range"]["search_range"]=="header":
                            end_index=int(table_start)-1
                        regex=config_dict["record"][1][i["range"]["regex_column"]]
                        re_list=i["range"]["list"]
                        header[i["name"]]=get_regexmatch(tokens,index_start,end_index,regex,re_list=re_list)
                    except:
                        #raise
                        #print(str(e))
                        logwrite("Error 17902 Error in Regex extraction")
                        
                if i["range"]["hint"]==10:
                    try:
                        iden=config_dict["record"][1][i["range"]["identity_column"]]
                        #print(config_dict["record"][1][i["range"]["distance"]])
                        distance=config_dict["record"][1][i["range"]["distance"]]
                        if distance.strip()=='':
                            distance=0
                        else:
                            distance=int(distance.strip())
                        header[i["name"]]=get_distancetoken(tokens,iden,distance)
                    except Exception as e:
                        #raise
                        #print(str(e))
                        logwrite("Error 17903 Error in Distance token extraction")
                        
                if i["range"]["hint"]==11:
                    try:
                        identity1=str(config_dict["record"][1][i["range"]["identity_column1"]])
                        #print(config_dict["record"][1][i["range"]["distance"]])
                        identity2=str(config_dict["record"][1][i["range"]["identity_column2"]])
                        header[i["name"]]=get_betweentoken(tokens,identity1,identity2)
                    except Exception as e:
                        #raise
                        print(str(e))
                        logwrite("Error 17904 Error in Between tokens extraction")
                if i["range"]["hint"]==12:
                    try:
                        identity=config_dict["record"][1][i["range"]["identity_column"]]  
                        name_attribute=i["name"]  #renmae z,nn  #done 
                        if identity!="":
                            identity = json.loads(identity)
                            nameatt_value=find_tok_btwwords(identity,name_attribute,tokens)
                            for nv in nameatt_value:
                                header[nv[0]]=nv[1]            
                    except Exception as e:
                        #raise#comment later raise
                        print(str(e))
                        logwrite("Error 17904 Error in search word extraction")     
       
            else:
                header[i["name"]]=config_dict["record"][1][i["col"]]
    except:
        logwrite("Error in header extraction")
    return header,linelist
    
#####################################################################################    
def line_extraction(tokens,config_dict,line_index_table,items,linelist):
    lines=[]
    itemno = None
    global config
    try:
        for i in line_index_table:
            for j in items:
                if i[0]==j["key"] and i[1]==j["index"]:
                    j["start index"]=i[1]+int(config_dict["linestart_buffer"])
                    #if line_index_table.index(i)==len(line_index_table)-1:
                    j["end index"]=i[2]+int(config_dict["lineend_buffer"])
                    #else:
                    #    j["end index"]=i[2]+int(config_dict["lineend_buffer"])
                    j["value tokens"]=get_value_atIndexRange(tokens,j["start index"],j["end index"],culture=config_dict["culture"],unit_price_incrementer=config_dict["unit_price_incrementer"])
                    j["extended values"]=get_value_atIndexRange(tokens,j["start index"],j["end index"],culture=config_dict["culture"],unit_price_incrementer=config_dict["unit_price_incrementer"])
                    #print(j["extended values"],"****************")
                    line_attributes=config["line"]["attributes"]
                    
                    rangekey=None 
                    value_assigned=[] #used to keep the tokens consumed in hint 6
                    range_config={}
                    for q in linelist:
                        if q["range"]["hint"]==9:
                            try:
                                if q["range"]["search_range"] is not None:
                                    index_start=j["start index"]+q["range"]["search_range"]
                                    end_index=j["end index"]+q["range"]["search_range"]
                                else:
                                    index_start=j["start index"]
                                    end_index=j["end index"]
                                regex=config_dict["record"][1][q["range"]["regex_column"]]
                                #print("regex is:",regex)
                                #regex=j["range"][q["range"]["regex_column"]-1]
                                re_list=q["range"]["list"]
                                avoid_list=[]
                                #print(k["range"]["ignore"])
                                #print(j.keys())
                                avoid_list.append(j["key"])
                                avoid_list.append(j["token"])
                                if q["range"]["ignore"] is not None:
                                    for ig in q["range"]["ignore"]:
                                        try:
                                            avoid_list.append(get_token_by_index(tokens,j[ig][0][1])[0])
                                        except:
                                            avoid_list.append('')
                                #print("##########################")
                                #print(avoid_list)
                                j[q["name"]]=get_regexmatch(tokens,index_start,end_index,regex,re_list=re_list,avoid_list=avoid_list)
                            except Exception as e:
                                #raise
                                print(str(e))
                                logwrite("Error 17902 Error in Regex extraction")
                        if q["range"]["hint"] == 13:
                            try:
                                adhoc_value = json.loads(
                                    config_dict["record"][1][q["range"]["identity_column"]])
                                if itemno is None:
                                    itemno = adhoc_value["start_value"]
                                index = j["index"]
                                j[q["name"]] = itemno_check(
                                    itemno, adhoc_value, index, tokens)
                                #print(j[q["name"]],"value")
                                if adhoc_value["incrementer"] is not None:
                                    itemno += adhoc_value["incrementer"]
                            except Exception as e:
                                print(str(e))
                                logwrite("Error 17005 in Item No Mapping")

                        
                    for k in line_attributes:
                        
                        if k["range"]["hint"]==0:
                            lowerbound =j["range"][k["range"]["colstart"]-1]
                            upperbound=j["range"][k["range"]["colend"]-1]
                            ignore_values=[j[element][0][0] if isinstance(element,tuple) else j[element] for element in k["range"]["ingnore_items"]]
                            range_config["lowerbound"]=lowerbound
                            range_config["upperbound"]=upperbound
                            range_config["ignore_values"]=ignore_values
                            j[k["name"]]=get_rangematch(j["extended values"],lowerbound,upperbound,hint=0, stoplist=config["stoplist"],culture=config_dict["culture"],ignore_values=ignore_values)
                            
                            if rangekey is None:
                                rangekey=j[k["name"]][0]
                                   
                        #Avoid "else" as there could be more breakdowns later
                        if k["range"]["hint"]==6:
                            price_recheck=False
                            if not rangekey is None:
                                                                
                                lowerbound =j["range"][k["range"]["colstart"]-1]
                                upperbound=j["range"][k["range"]["colend"]-1]
                                
                                extended_values=j["extended values"]
                                decimals=k["range"]["decimals"]
                                
                                #acc_price=rangekey
                                
                                if config_dict["unique_token"] is not None:
                                    if config_dict["unique_token"].lower() in config["minimal_qty_cm"]:
                                        quantity_match=product_match(extended_values,lowerbound=lowerbound,upperbound=upperbound, rangekey=rangekey, decimals=decimals, value_assigned=[],rounding=int(config_dict["presicion"]))
                                    else:
                                        quantity_match=product_match(extended_values,lowerbound=lowerbound,upperbound=upperbound, rangekey=rangekey, decimals=decimals, value_assigned=[],rounding=int(config_dict["presicion"]),minimal_qty=10)
                                
                                else:
                                    quantity_match=product_match(extended_values,lowerbound=lowerbound,upperbound=upperbound, rangekey=rangekey, decimals=decimals, value_assigned=[],rounding=int(config_dict["presicion"]),minimal_qty=10)
                                
                                if (quantity_match[0]==None) and (rangekey is not None) and (j["Price"][1]<1):
                                    price_recheck=True
                                    new_range=rangekey
                                    
                                while price_recheck:
                                    range_config["ignore_values"].append(new_range[0])
                                    #print(range_config["ignore_values"])
                                    new_range_value=get_rangematch(j["extended values"],range_config["lowerbound"],range_config["upperbound"],hint=0, stoplist=config["stoplist"],culture=config_dict["culture"],ignore_values=range_config["ignore_values"])
                                    new_range=new_range_value[0]
                                    if new_range!=None:
                                        if config_dict["unique_token"] is not None:
                                            if config_dict["unique_token"].lower() in config["minimal_qty_cm"]:
                                                quantity_match=product_match(extended_values,lowerbound=lowerbound,upperbound=upperbound, rangekey=new_range, decimals=decimals, value_assigned=[],rounding=int(config_dict["presicion"]))
                                            else:
                                                quantity_match=product_match(extended_values,lowerbound=lowerbound,upperbound=upperbound, rangekey=new_range, decimals=decimals, value_assigned=[],rounding=int(config_dict["presicion"]),minimal_qty=10)
                                        else:
                                            quantity_match=product_match(extended_values,lowerbound=lowerbound,upperbound=upperbound, rangekey=new_range, decimals=decimals, value_assigned=[],rounding=int(config_dict["presicion"]),minimal_qty=10)
                                        if quantity_match[1]==1:
                                            j["Price"]=new_range_value
                                            price_recheck=False
                                    else:
                                        price_recheck=False
                                    if new_range_value[1]==1:
                                        price_recheck=False
                                            
                                
                                j[k["name"]]=quantity_match
                                                                
                                #j[k["name"]]=product_match(extended_values,lowerbound=lowerbound,upperbound=upperbound, rangekey=rangekey, decimals=decimals, value_assigned=[],rounding=int(presicion))
                                if j[k["name"]][0]!= None:
                                    value_assigned.append(j[k["name"]][0])
                            else:
                                logwrite("Error 16009: Field with hint 0 should be the first item in range. Check config file")
                        
                        #date Match
                        if k["range"]["hint"]==7:
                            try:
                                
                                index_start=j["start index"]
                                end_index=j["end index"]
                                lowerbound_delta =int(float(j["range"][k["range"]["start_date_delta_column"]-1]))
                                upperbound_delta=int(float(j["range"][k["range"]["end_date_delta_column"]-1]))
                                
                                extended_values=j["extended values"]
                                date_tokens=get_token_atIndexRange(tokens,index_start,end_index)
                                if (k["range"]["identity_column"]) is not None:
                                    identity=str(j["range"][k["range"]["identity_column"]-1])
                                    if identity.strip()!='':
                                        #print(identity)
                                        identity_token=get_identitytoken(date_tokens,identity)
                                        if identity_token is not None:
                                            index_start=identity_token[1]
                                            end_index=identity_token[1]+5
                                            #date_tokens=get_token_atIndexRange(date_tokens,,)
                                j[k["name"]]=get_date(tokens,start_index=index_start,end_index=end_index,delta_back =lowerbound_delta, delta_forward=upperbound_delta,format=config_dict["date_format"])
                                
                                if j[k["name"]]!= None:
                                    value_assigned.append(j[k["name"]][0])
                            except Exception as e:
                                #raise
                                print(str(e))
                                logwrite("Error 16901 Error in Date extraction")
                                
                        if k["range"]["hint"]==9:
                            try:
                                if k["range"]["search_range"] is not None:
                                    index_start=j["start index"]+k["range"]["search_range"]
                                    end_index=j["end index"]+k["range"]["search_range"]
                                else:
                                    index_start=j["start index"]
                                    end_index=j["end index"]
                                regex=j["range"][k["range"]["regex_column"]-1]
                                re_list=k["range"]["list"]
                                avoid_list=[]
                                avoid_list.append(j["key"])
                                avoid_list.append(j["token"])
                                #print(k["range"]["ignore"])
                                if k["range"]["ignore"] is not None:
                                    for ig in k["range"]["ignore"]:
                                        try:
                                            avoid_list.append(get_token_by_index(tokens,j[ig][0][1])[0])
                                        except:
                                            avoid_list.append('')
                                #print(avoid_list)
                                j[k["name"]]=get_regexmatch(tokens,index_start,end_index,regex,re_list=re_list,avoid_list=avoid_list)
                                #print(header[i["name"]])
                                
                                #if header[i["name"]]!= None:
                                #value_assigned.append(header[i["name"]][0])
                            except Exception as e:
                                #raise
                                print(str(e))
                                logwrite("Error 16999 Error in Regex extraction")
                        if k["range"]["hint"]==12:
                            try:
                                index_start=j["start index"]
                                end_index=j["end index"]
                                if (j["range"][k["range"]["identity_column"]-1]) is not None:
                                    #print(j["range"][k["range"]["identity_column"]-1])
                                    j[k["name"]]=get_brokentokens(tokens,index_start,end_index,j["range"][k["range"]["identity_column"]-1])
                            except:
                                logwrite("Error 17000 Error in broken token extraction")
                    #print(j)
                    lines.append(j)
                    
    except Exception as e:
        logwrite(str(e))
        logwrite("Error in line extraction")
    
    return lines
    
####Main function####################################################################

def get_extract(file,unique_token=None,generate_tokenfile=False):
    global config
    global table_start, table_end
    dict={}
    try:
        dict={}
        initialize_case()
        tokens=None
        
        ########## invoke get tokens to get all the tokens ###########################################
        tokens=get_tokens(file)
        
        ########## appending unique_token identifier to the list on tokens ############################
        if unique_token is not None:
            tokens.append((unique_token,-1))
        
        ########## Writting tokens to a file when generate_tokenfile flag is set true #################
        if generate_tokenfile:
            f= open(file+".txt","w+",encoding='utf-8')
            for t in tokens:
                try:
                   f.write(' '.join(str(s) for s in t) + '\n')
                except:
                    continue
            f.close() 
        
        ######### Getting entity records ##############################################################
        record=get_keyrecord(tokens,config["header"]["path"],config["header"]["key"]["col"])
        #print(record)
        
        
        ########################  loading configuration  ##############################################
        header={}
        header["key"]=record[1][config["header"]["key"]["col"]]
        header["entity"]=record[1][config["header"]["attributes"][0]["col"]]
        ls=config["header"]["attributes"]
        config_dict={}
        config_dict=get_header_configuration_data(record)
        config_dict["unique_token"]=unique_token
        if config_dict["pdf_reader"] is not None:
            if config_dict["pdf_reader"]!="pdfplumber" and file[-3:]=="pdf":
                #print(pdf_reader)
                tokens=get_tokens(file,pdf_reader=config_dict["pdf_reader"])
                #print(tokens)
        
        dict["header"]=header

    except Exception as e:
        #raise
        print(str(e))
        logwrite("Error 140146 Header parsing error")

    ##########################     LINE       #######################################################################################
    #Getting item matches
    lines={}
    
    try:
        logwrite("fetching items")
        
        ###################### Fetching matching key attributes ####################################################################
        
        items=get_keyrecords(tokens,config["line"]["path"],config["line"]["key"]["col"],match_based_on=config_dict["Part_match_on"],match_type=config_dict["Part_match_type"],ignorecol=config["line"]["key"]["ignore_col"],connection_type=config["line"]["connection_type"],connection_details=config["line"]["connection_details"])
        #print(items,"temp")
        
        ###################### getting range of tokens for the item using indxes betweeen two consecutive part numbers #############
        line_index_table=[]
        for ln in items:
            #print(ln, "dict")
            line_index_table.append([ln["key"],ln["index"]])
            if table_start>ln["index"]:
                
                table_start=ln["index"]
            if table_end<ln["index"]:
                table_end=ln["index"]+10
        

        line_index_table=get_line_index_table(line_index_table)
        
        
        ##############################   Fetch Header elements based on the rules    #############################################
        header,linelist=header_extraction(tokens,config_dict,ls,header)
        
        lines=[]
       
        ##############################   Fetch Line elements based on the rules    ##############################################
        lines=line_extraction(tokens,config_dict,line_index_table,items,linelist)
                    
        ##############################  checking for duplicate records in the final dataframe  ###################################
        try:
            if len(lines)>0:
                records=pd.DataFrame(lines)
                keys_with_price=set(records[records['Price']!=(None,-1)]['key'].tolist())
                df_keys_with_price=records[records['Price']!=(None,-1)]
                df_keys_with_price['only_price_value']=df_keys_with_price['Price'].apply(lambda x:x[0][0])
                groupby_keys_with_price_n_key=df_keys_with_price.groupby(['key','only_price_value'])
                groupby_keys_with_price=df_keys_with_price.groupby('key')
                cross_check_for_none_quantity=[]
                for key,value in groupby_keys_with_price.groups.items():
                    for k,v in groupby_keys_with_price_n_key.groups.items():
                        if key in k:
                            if (v.tolist() != value.tolist()) and len(list(set(value.tolist())-set(v.tolist())))!=len(value.tolist()):
                                sample_df=records.loc[value,'Quantity'][records.loc[value,'Quantity']==(None,-1)].index.tolist()
                                #sample_df=records.loc[value,records.Quantity == (None,-1)].index.tolist()
                                cross_check_for_none_quantity.extend(sample_df)
                            
                cross_check_for_none=records[(records['Price']==(None,-1)) & (records['key'].isin(keys_with_price))].index.tolist()
                cross_check_for_none.extend(cross_check_for_none_quantity)
                cross_check_for_none_tolist=list(set(cross_check_for_none))
                cross_check_for_none_sorted=sorted(cross_check_for_none_tolist)
                for i in cross_check_for_none_sorted[::-1]:
                    lines.pop(i)
        except Exception as e:
            logwrite(str(e))
            logwrite("Error in cross_checking the key,price and quntity")
                
        dict["lines"]=lines
        return dict
            
    except Exception as e:
        logwrite(str(e))
        logwrite("Error 140150 Parsing line error")
        #raise