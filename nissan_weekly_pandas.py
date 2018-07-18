"""
.. module:: nissan_weekly
   :platform: Windows
   :synopsis: A useful module indeed.

.. moduleauthor:: Xolani Mazibuko <xolani@ddi.media>

"""
import pandas as pd
import requests
from openpyxl import load_workbook
import os
import datetime

subjectList = ['GT-R', 'NV 350 Impendulo', 'Qashqai', 'NV 200 Combi',
               'Tiida', 'Leaf', 'NAAMSA', 'K-line', 'Juke', 'X-Trail',
               'Navara', 'Micra', 'NP 200', 'Patrol', 'Sentra',
               'Nissan Festival of Motoring',
               'Nissan drives for further growth',
               'ICC Sponsorship', 'NP 300', 'Nissan Trailseeker',
               'Nissan Corporate',
               'Nissan Taxi Call Centre oppurtunity',
               'Nissn long distance driving tips',
               'Nissan Holiday Checklist', 'Pathfinder', 'NP 300 hardbody',
               'NV 300', 'Workhorse', 'Almera', '350z', 'Hardbody',
               'Simola Hillclimb', '370z', 'Murano', 'Nissan',
               'Nissan Kicks', 'Nissan Primera', 'Bladeglider',
               'Nissan ProPilot', 'Festival of Motoring', 'NV 350 Panel Van',
               'Patrol Wagon', 'Livina', 'Nissan Sponsorship/sport',
               'Nissan Motorsport','Nissan Spokespeople',
               'General Industry related']

productList = ['NP 300', 'GT-R', 'Qashqai', 'NV 200 Combi', 'Tiida',
               'NV 300',  'Leaf', 'K-line', 'Juke', 'X-trail', 'NP 300',
               '350z', 'Navara', 'Micra', 'NP 200', 'Patrol', 'Sentra' ,
               '370z', 'Murano', 'Pathfinder', 'NP 300 hardbody', 'NV 300',
               'Workhorse', 'Almera', 'NV 350 Panel Van', 'Livina',
               'Patrol Wagon', 'Nissan Primera', 'Nissan ProPilot',
               'Bladeglider']

corporateList = ['NAAMSA', 'Nissan Corporate', 'Nissan', 'Nissan Motorsport',
                 'Nissan Spokespeople', 'General Industry related',
                 'Nissan Sponsorship/sport', 'Festival of Motoring',
                 'Nissn long distance driving tips',
                 'Nissan Holiday Checklist', 'Nissan Taxi Call Centre oppurtunity',
                 'ICC Sponsorship', 'Nissan Festival of Motoring',
                 'Nissan drives for further growth', 'Nissan Trailseeker']


def inputfile(excelFile):
    """

    :param excelFile:
    :return:
    """
    df = pd.read_excel(excelFile)
    return df


def inputfilePrint(excelFile):
    """Takes in an excel file name and returns a DataFrame without the
        first two rows.

    Args:
       excelFile (str):  The excel files name.

    Returns:
       class DataFrame. 
    """
    df = pd.read_excel(excelFile, skiprows=2)
    return df


def online_fix(df):
    #  Change Sentiment from number to String
    for index in range(len(df.Sentiment)):
        df.loc[[index], 'Sentiment']


def print_fix(df):
    """
    Changes values for each row.
    """
    df['Category'] = ''


    #  Decides whether the category is Product or Corporate
    category = []
    for index in range(len(df.Subjects)):
        category = []
        subList = df.Subjects[index]
        for sub in subList.split(', '):
            try:
                if productList.index(sub) != -1:
                    cat = 'Product'
                    break
            except ValueError:
                cat = 'Corporate'
        category.append(cat)
        df.loc[[index], 'Category'] = category[0]
    
    #  Change Sentiment from number to String
    for index in range(len(df.Sentiment)):
        if df.Sentiment[index] == 1:
            df.loc[[index], 'Sentiment'] = "Positive"
        elif df.Sentiment[index] == 0:
            df.loc[[index], 'Sentiment'] = "Neutral"
        elif df.Sentiment[index] == -1:
            df.loc[[index], 'Sentiment'] = "Negative"
    
    #df = df.rename(columns = {'Referred Date' : 'Date'})

    #  Converts the values to datetime values
    df['Referred Date'] = pd.to_datetime(df['Referred Date'])
    #  Gets rid of the time
    df['Referred Date'] = df['Referred Date'].dt.date
                              
    return df


def broadcast_fix(df):
    df.Date  = pd.to_datetime(df.Date)
    df.Date = df.Date.dt.date
    ind = os.getcwd().rfind('\\')
    folderName = os.getcwd()[ind+1:]
    date = folderName[8:15]

    

def rankAVEData(df):
    df.sort_values(['Category', 'Ave'], ascending=False, inplace=True)

    df = df[['Title', 'Media', 'Scanned Date', 'Sentiment', 'Readership',
             'Ave', 'Text']]

    df.rename(columns = {'Scanned Date' : 'Date',
                         'Readership' : 'Read',
                         'Ave' : 'AVE'})

    ##  Product
    df = df.drop_duplicates(subset=['Title'])

    ##  Corporate
    df = df.reset_index()
    #https://stackoverflow.com/questions/41255215/pandas-find-first-occurence?utm_medium=organic&utm_source=google_rich_qa&utm_campaign=google_rich_qa
    index = dftest2.Category.eq('Corporate').idxmax()

    dfProduct = df[:5]
    dfCorporate = df[index:index+5]

    return dfProduct, dfCorporate

def copysnapshotData(excelFile):
    wb = load_workbook(excelFile)

    


##excelFile = 'Print RAW fix.xlsx'
##Sheet = 'Sheet1'
##df, df2 = inputfile(excelFile, Sheet)
##outputsheet(df2)


##outputfile(df, df.columns, 'Powerpoint Data1 - Copy.xlsx')
#df = inputfilePrint('Print RAW.xlsx')
