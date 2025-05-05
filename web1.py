from urllib.request import urlopen 
from bs4 import BeautifulSoup
a='https://www.w3schools.com/html/'

file=urlopen(a)
html=file.read ()




sp = BeautifulSoup(html,'html.parser')
print (sp.title)
print (sp.title.string)

for script in sp  (['script,style']):
    script.extract ()

text= sp.get_text()
lines = (line.strip () for line in text.splitlines())

L1=[]

for line in lines :
    for line in lines:
        words = line.split()
        L1.extend(words)
print(L1)
import re

# using regular expression to remove symbol and number
patt = re.compile (r'(?<!\/S)([A-Za-z])+(?!\S)')
L2=[]
for word in L1:
    mt=patt.match (word)
    if (mt):
                       
        L2.append (mt.group().lower())
                                  
#print(L2)
ignore_list = [] #blank list
#populating ignore list from file ignore.txt

with open ('C:\\Users\\Yuvraj pradhan\\New folder\\project1\\ignore.txt','r') as file2:
    data = file2.readlines()
    for line in data:
        for x in line.split():
            ignore_list.append(x.lower())

L2.sort() #sorting words and their frequency
Dict_word={} # blank dictonary for sorting word and their freaquency
for word in L2:

    cnt = L2.count (word)
    Dict_word[word]=cnt

#print(Dict_word)

for word in ignore_list:
    if word in Dict_word:
        del Dict_word[word]

print ('*'*50)

#print (Dict_word)
 
 # database coding 
 #starting database connection 

import sqlite3

conn= sqlite3.connect('project.db')
print ('open database success')

 #creating a new table 
conn.execute ('''create table if not exists word (word text not null,count int not null);''') 

print("table created successfully")

#inserting record in table

for word,count in Dict_word.items():
    count = str (count)
    c ="INSERT INTO WORD VALUES ('"+word+"',"+count+");"

    conn.execute(c)
conn.commit()
print('recod created successfully')

cur = conn.execute("select * from word ")

for row in cur:
    print (row [0],':',row [1])

conn.close()

# FIND the 10 most frequently used  words 

a= {}
import operator
a = dict (sorted (Dict_word.items(),key = operator.itemgetter(1),reverse=True)[:10])
print(a) 


import xlsxwriter

workbook = xlsxwriter.Workbook ('C:\\Users\\Yuvraj pradhan\\New folder\\project1\\word.xlsx')
worksheet = workbook.add_worksheet ()
worksheet.set_column ('A:A',30)
bold= workbook.add_format({'bold':True})
worksheet.write ('A1','WORDS',bold)
worksheet.write ('B1','COUNT',bold)

row = 1
col = 0

#writing the most used in excel sheet

for word, count in a.items():
    worksheet.write (row,col,   word)
    worksheet.write (row,col+ 1, count )
    row += 1

#create a new chart object 

chart = workbook.add_chart ({'type':'pie'})

# setting the axis attributes 

chart.set_x_axis({
    'name':'words',
    'name_font': {'size': 14, 'bold': True},
    'num_font': {'italic': True},
  })

chart.set_y_axis ({
    'name': 'freaquency',
    'name_font': {'size': 14, 'bold':True},
    'num_font':  {'italic': True},    
})

#configure the chart . in simplest case we add one or more data series.

chart.add_series ({'values': '=Sheet1!$B$2:$B$11',
                   'categories': '=Sheet1!$A$2:$A$11'})

#insert the chart into the worksheet.
worksheet.insert_chart ('D2',chart)

workbook.close()