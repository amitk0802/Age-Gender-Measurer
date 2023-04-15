# -*- coding: utf-8 -*-
"""
Created on Jan 19 23:45:34 2022

@author: Amit
"""

import pandas as pd
import en_core_web_sm
nlp = en_core_web_sm.load()
import tkinter as tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
import time
import warnings
warnings.filterwarnings("ignore")
import logging

if __name__ == '__main__':
    try:
        # ------------------------- READING STATIC FILE ----------------------
        # AGE
        data_age_df = pd.read_csv(r'E:\RA\Age_Gender_Lexica\emnlp14age.csv',skiprows = [1])
        data_age_df.set_index("term", drop=True, inplace=True)
        db_age_dictionary = data_age_df.to_dict(orient="index")
        intercept_age = 23.2188604687
        
        # GENDER
        data_gender_df = pd.read_csv(r'E:\RA\Age_Gender_Lexica\emnlp14gender.csv',skiprows = [1])
        data_gender_df.set_index("term", drop=True, inplace=True)
        db_gender_dictionary = data_gender_df.to_dict(orient="index")
        intercept_gender = -0.06724152
        
        tk.Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
        input_file = askopenfilename()
        ext = input_file.split(".")[-1]
        print("Extension : {}".format(ext))
        present_date = time.strftime('%Y%m%d') 
        input_file_name = str(input_file.rsplit('/', 1)[1])
        output_file_name = str(input_file_name).replace('.' + str(ext),' - Output from AgeAndGenderMeasurer.' + str(ext))
        output_file = str(input_file.rsplit('/', 1)[0]) + '/' + str(present_date) + '-' + output_file_name
        
        print("Input File : {}".format(input_file))
        if ext == 'xlsx':
            data_df = pd.read_excel(input_file, sheet_name = 0, engine = 'openpyxl')
            print('File Read Success!')
        
        elif ext == 'xls':
            data_df = pd.read_excel(input_file, sheet_name = 0, engine = 'xlrd')
            print('File Read Success!')
        
        elif ext == 'csv':
            data_df  = pd.read_csv(input_file)
            print('File Read Success!')
        
        else:
            print('FILE FORMAT NOT SUPPORTED!')
        
        col1 = data_df.columns[0]
        col2 = data_df.columns[1]
        data_df.rename(columns={ data_df.columns[0]: "id",  data_df.columns[1]: "para"}, inplace = True)
        
        data_df['Age'] = ''
        data_df['Gender'] = ''
        
        punc = '''!()-[]{};:'"\,<>./?@#$%^&*_~'''
        
        for i in range(len(data_df)):
            a = data_df['para'][i]
            for ele in a:
                if ele in punc:
                    a = a.replace(ele, " ")
            getDocFile = nlp(a)
            getAllWords = [token for token in getDocFile]
            
            total_word = len(getAllWords)
            weight_age_list = []
            for j in range(len(getAllWords)):
                check_age_word = str(getAllWords[j])
                if check_age_word in db_age_dictionary:
                    weight_age_list.append(db_age_dictionary[check_age_word]['weight']/total_word)
            age = sum(weight_age_list) + intercept_age
            
            total_word = len(getAllWords)
            weight_gender_list = []
            for k in range(len(getAllWords)):
                check_gender_word = str(getAllWords[k])
                if check_gender_word in db_gender_dictionary:
                    weight_gender_list.append(db_gender_dictionary[check_gender_word]['weight']/total_word)
            gender = sum(weight_gender_list) + intercept_gender
            
            data_df['Gender'][i] = gender
            data_df['Age'][i] = age
        
        data_df.rename(columns={ "id" : col1, "para" : col2}, inplace = True)
        data_df.reset_index(drop=True, inplace = True)
        
        # ---------------------- WRITING OUTPUT ---------------------- #
        if (ext == 'xlsx') or (ext == 'xls'):
            data_df.to_excel(output_file, sheet_name = 'Sheet1', index = False)
        elif ext == 'csv':
            data_df.to_csv(output_file, index = False)
        
        print('\nCOMPLETED SUCCESSFULLY !!')
        
    except Exception as e:
        logfile = str(input_file.rsplit('/', 1)[0]) + '/' + str(present_date) + ' - AgeAndGenderMeasurer.log' 
        logging.basicConfig(filename=logfile, 
                            level=logging.DEBUG)
        logger=logging.getLogger(__name__)
        logger.error(e)
        logger.info(input_file)