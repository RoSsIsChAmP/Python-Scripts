#Modules required
import xlrd
import urllib
import imgkit
import datetime

#This program is designed to read an excel document and perform an online request of the domain to verify if it exists and if it does it will screenshot the page.
#Created by Mitchell Blystone-Pemberton


#excel doc variable
print (" ___ ____    ____                        _         _                _                ")
print ("|_ _|  _ \  |  _ \  ___  _ __ ___   __ _(_)_ __   | |    ___   ___ | | ___   _ _ __  ")
print (" | || |_) | | | | |/ _ \| '_ ` _ \ / _` | | '_ \  | |   / _ \ / _ \| |/ / | | | '_ \ ")
print (" | ||  _ <  | |_| | (_) | | | | | | (_| | | | | | | |__| (_) | (_) |   <| |_| | |_) |")
print ("|___|_| \_\ |____/ \___/|_| |_| |_|\__,_|_|_| |_| |_____\___/ \___/|_|\_\\__,_| .__/ ")
print ("                                                                              |_|    ")
print ("                                                               By: Mitchell Pemberton")
print ("Please enter the path to the excel doc you wish to use:")
path2doc = raw_input()

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
          print(sheet.cell_value(i, 0)),"HTTP Code:"
          print urllib.urlopen(sheet.cell_value(i, 0)).getcode()
          imgkit.from_url(sheet.cell_value(i, 0), str(now) +".jpg")
          print("------------------------")
     except Exception as e:
          print e

