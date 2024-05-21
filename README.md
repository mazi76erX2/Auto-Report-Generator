# Auto-Report-Generator

A Project I worked on for a company in 2017 (Used with permission)

Considering refactoring the code to show what I've learnt over the years

## Description

Generates a excel and powerpoint report using [pandas](https://pandas.pydata.org), [openpyxl](https://openpyxl.readthedocs.io/en/stable) and [python-pptx](https://openpyxl.readthedocs.io/en/stable) using data exported from [DataScouting](https://datascouting.com/) and [Mention](https://mention.com/en/).

## Usage

1. Place the following files in the same directory and rename them as follows:

* DataScouting File --> print.xlsx
* Mention --> onlineandsocial.xlsx
* Google sheets --> broadcast.xlsx
* Mention --> onlineSSA.xlsx
* Mention --> socialSSA.xlsx
* delete_list.csv
* AVE Template.xlsx

![directory](./data-files/directory.PNG "directory")

2. run [nissan_weekly_pandas.py](https://github.com/mazi76erx2/nissan_weekly_pandas.py). This will generate the following files:

* AVE Data.xlsx
* AVE Report (start date) - (end date).xlsx for example AVE Report 09.08.2018 - 15.08.2018.xlsx 
* SSA digital data for weekly report (start date) - (end date).xlsx 

The output excel file looks like this:
![ave-data](./data-files/ave-data.PNG "AVE Data")

3. Select the articles the articles that you would like to add to the powerpoint presentation as follows:

* In the Print and Online tab highlight 5 product articles in yellow (#FFFFFF00) and 5 corporate articles in green (#FF92D050).

![alt text](./data-files/excel-data.PNG "AVE Data")

* In the Broadcast tab highlight 2 broadcasts in red (#FFFF0000).

![alt text](./data-files/excel-data2.PNG "AVE Data")

* In the OnlineSSA and SocialSSA highlight 5 articles in orange (#FFFFC000) and 3 tweets in light blue (#FF00B0F0).

![alt text](./data-files/excel-data3.PNG "AVE Data")
![alt text](./data-files/excel-data4.PNG "AVE Data")

4. run [powerpoint_nissan.py](https://github.com/mazi76erx2/powerpoint_nissan.py). This will generate the following file:

* Nissan Weekly Report (start date) - (end date).pptx for example Nissan Weekly Report  09 August - 15 August 2018

![alt text](./data-files/powerpoint-data.PNG "AVE Data")
