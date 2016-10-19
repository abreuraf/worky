#!/usr/bin/python3
import xlrd
import sys
import pymysql
import time, os
import configparser

# List of data extract from excel workbook
list1 = []

# Open database connection
config = configparser.RawConfigParser()
config.read("database.conf")
db = pymysql.connect(config.get('database','hostname'),config.get('database','username'),config.get('database','password'),config.get('database','database'))

# prepare a cursor object using cursor() method
cursor = db.cursor()

# Load workbook
book = xlrd.open_workbook("planilha.xlsx")

# Select the first sheet in workbook
sh = book.sheet_by_index(0)

# value to search/add
v = input("P/N: ")

# search for value and get all data of the part with the PN indicated 
for ry in range(sh.nrows):
    if sh.cell_value(rowx=ry, colx=0) == v:
        for rx in range(17):
            list1.append(sh.cell_value(rowx=ry,colx=rx))
               
print("WORKBOOK: ",list1)

# execute SQL query
cursor.execute("SELECT * FROM Part WHERE name=%s",v)

# Fetch a single row 
data = cursor.fetchone()

# Verify in the database if doesn't exist and add
if data is not None:
    cursor.execute("SELECT id FROM Part WHERE name=%s",v)
    data = cursor.fetchone()
    query = "INSERT INTO purchases_purchase (name, ps, quantity_buy, quantity_loss, package, price_unit, price_loss, price_checkout, part_id) VALUES (%s,%s,%i,%i,%s,%d,%d,%d,%s)"
    cursor.execute(query,())
else:
    time = time.strftime('%Y-%m-%d %H:%M:%S')
    query = ("""INSERT INTO Part (category_id, footprint_id, name, description, comment, stockLevel, minStockLevel, averagePrice,status, needsReview, partCondition, createDate, internalPartNumber, removals, lowStock, partUnit_id, storageLocation_id) VALUES (1,NULL,'%s','%s','%s',%i,0,%d,'',0,'','%s','%s',0,0,1,1)""")%(list1[0],list1[9],list1[12],int(list1[3]),round(list1[6],2),time,list1[4])
    affected_count = cursor.execute(query)
    db.commit()
    print("Number of inserted rows: ")
    print(affected_count)

# disconnect from server
db.close()



		
