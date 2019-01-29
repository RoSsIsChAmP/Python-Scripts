#!/usr/bin/env python
#Modules required
import xlrd
import urllib
import imgkit
import datetime
import whois

#Written using Python 2.7
#This program is designed to read an excel document and perform an online request of the domain to verify if it exists and if it does it will screenshot the page.
#Created by Mitchell Blystone-Pemberton


#excel doc variable
print ("__        __   _         _ _         _                _                ")
print ("\ \      / /__| |__  ___(_) |_ ___  | |    ___   ___ | | ___   _ _ __  ")
print (" \ \ /\ / / _ \ '_ \/ __| | __/ _ \ | |   / _ \ / _ \| |/ / | | | '_ \ ")
print ("  \ V  V /  __/ |_) \__ \ | ||  __/ | |__| (_) | (_) |   <| |_| | |_) |")
print ("   \_/\_/ \___|_.__/|___/_|\__\___| |_____\___/ \___/|_|\_ \__,_| .__/ ")
print ("                                                                |_|    ")
print ("                                                 By: Mitchell Pemberton")
print ("-----------------------------------------------------------------------")
print ("                              Main Menu                                ")
print ("                   1. Lookup using spreadsheet                         ")
print ("                   2. Lookup using URL                                 ")
print ("-----------------------------------------------------------------------")
user_ans = raw_input("Please pick a number: ")

if user_ans == "1":
     path2doc = raw_input("Please enter the path to the excel doc you wish to use: ")

     exceldoc = path2doc

     #Opens the excel doc
     wb = xlrd.open_workbook(exceldoc)

     sheet = wb.sheet_by_index(0)
     #adjust values to read from different rows and columns
     sheet.cell_value(0, 0)

     #Performs the loop to scan all the entries in the excel book
     for i in range(sheet.nrows):
          try:
               now = datetime.datetime.now()
               print ("------------------------")
               print (sheet.cell_value(i, 0)),"HTTP Code:"
               print urllib.urlopen("http://"+sheet.cell_value(i, 0)).getcode()
               print whois.whois(sheet.cell_value(i, 0))
               imgkit.from_url(sheet.cell_value(i, 0), str(now) +".jpg")
          except Exception as e:
               print e

if user_ans == "2":
     lookup = raw_input("Please enter the URL you wish to lookup I.E. www.example.com: ")
     try:
          now = datetime.datetime.now()
          print("------------------------")
          print lookup, "HTTP Code:"
          print urllib.urlopen("http://"+lookup).getcode()
          print whois.whois(lookup)
          imgkit.from_url(lookup, str(now) + ".jpg")
          print("------------------------")
     except Exception as e:
          print e
