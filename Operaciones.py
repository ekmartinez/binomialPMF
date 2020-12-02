# -*- coding: utf-8 -*-
"""
Created on Sat May 16 17:59:45 2020

@author: Erick Martinez Torres
"""
import matplotlib.pyplot as plt
import openpyxl
import pandas as pd
from openpyxl import load_workbook
from scipy.stats import binom


class  Distribucion:
    ''''Displays the binomial probability distribution'''
    def __init__(self, tests, prb):
        self.tests = tests + 1
        self.prb = prb

    def Binomial(self):
        #Create a range object
        lst = []
        for x in range(self.tests):
            lst.append(x)
            
        #Data Frame - Binomial computation.
        df = pd.DataFrame(lst, columns = ['Tests'])
        df['Binomial'] = binom.pmf(df['Tests'], self.tests, self.prb)
        pd.options.display.float_format = '{:.8f}'.format
        df.set_index('Tests', inplace = True)
        print(df)

        #Matplotlib Barchart
        plt.bar(df.index, df['Binomial'], color='slategray')
        plt.title('Probability Mass Function')
        plt.ylabel('Probability')
        plt.xlabel('Number of Tests')
        plt.savefig('pmf.png', dpi=100)
        plt.show()
        
        #Export to Excel.
        fname = 'Binomial_distribution.xlsx'
        df.to_excel(fname)
        wb = load_workbook(fname)
        ws = wb.active
        img = openpyxl.drawing.image.Image('pmf.png')
        img.anchor = 'E2'
        ws.add_image(img)
        wb.save(fname) 
     
       
       
        
        
