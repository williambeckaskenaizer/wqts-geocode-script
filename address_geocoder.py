#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Geocoding script made by William Beck-Askenaizer for WQTS 
A little hacky, but it'll do.

7/9/2020
"""

import platform
import os
import subprocess
import csv
import sys
import time
import tkinter as tk
from tkinter import filedialog

"""
Make sure packages are installed
"""


def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
install_list = ["pandas", "geopy", "xlrd"]
for package in install_list:
    install(package)
import pandas as pd
from geopy.geocoders import ArcGIS, Bing, Nominatim, OpenCage, GoogleV3, OpenMapQuest

def print_logo():
    print("\t **       **   *******    **********  ********")
    print("\t/**      /**  **/////**  /////**///  **////// ")
    print("\t/**   *  /** **     //**     /**    /**       ")
    print("\t/**  *** /**/**      /**     /**    /*********")
    print("\t/** **/**/**/**    **/**     /**    ////////**")
    print("\t/**** //****//**  // **      /**           /**")
    print("\t/**/   ///** //******* **    /**     ******** ")
    print("\t//       //   /////// //     //     ////////  ")
    print("\n")
    print("\t Geolocation Script for Address Batches\n\n")


"""
 State codes for later
"""
abbr = ["AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA",
        "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD",
        "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
        "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC",
        "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY"]


def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])


install_list = ["pandas", "geopy"]
for package in install_list:
    install(package)

"""
clear the window :)
"""
os.system('cls') if platform.platform() == "Windows" else os.system('clear')
print_logo()


""" 
initialize geocoders (just using ArcGIS for now)
"""
arcgis = ArcGIS(timeout=100)
# nominatim = Nominatim(user_agent="WQTS", timeout=100)
# opencage = OpenCage('your-API-key', timeout=100)
# openmapquest = OpenMapQuest('api-key', timeout=100)
# geolocator = Nominatim(user_agent="WQTS")


"""
get excel file location/info from user
"""
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
base_file = pd.ExcelFile(file_path)


print("Available sheets in selected file are: \n")
sheet_num = 1
for sheet in base_file.sheet_names:
    print(sheet_num, ":", sheet)
    sheet_num += 1
print('\n')

address_num = input(
    "Enter the number next to the desired sheet in the list above (e.g. 1 for " + str(base_file.sheet_names[0]) + "): ")


address_sheet = base_file.sheet_names[int(address_num) - 1]
print("Selected sheet:", address_sheet)


header_sheet = base_file.parse(address_sheet, index_col=0)

sheet_columns = header_sheet.columns.to_list()

col_count = 1
print("Available columns in '", address_sheet, "' are: \n")
for column in sheet_columns:
    if("Unnamed" not in column):
        print(col_count, ":", column)
        col_count += 1

address_column_num = input(
    "Enter the number next to the column containing the addresses: ")

address_column = sheet_columns[int(address_column_num)-1]
print("Selected column:", address_column, "\n")


state = input("Enter the state code (e.g. CA for California): ").upper()
while state not in abbr:
    state = input(
        "That doesn't seem to be a state code\nPlease enter the state code: ").upper()
print("\n")

df = pd.ExcelFile(file_path).parse(address_sheet)
addr_list = df[address_column].to_list()


"""
Progress bar function courtesy of StackOverflow :)
"""


def printProgressBar(iteration, total, prefix='', suffix='', decimals=1, length=100, fill='█', printEnd="\r"):
    percent = ("{0:." + str(decimals) + "f}").format(100 *
                                                     (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end=printEnd)
    # Print New Line on Complete
    if iteration == total:
        print()


geocoded = [[]]
count = 0
l = len(addr_list)
for item in addr_list:
    count += 1
    if isinstance(item, str):
        item += ' ' + state
    full_address = arcgis.geocode(item)
    if(full_address == state):
        print("unable to locate for " + item)
    else:
        geocoded.append(
            [full_address, full_address.latitude, full_address.longitude])
    printProgressBar(count, l, prefix='Geocoding Addresses:',
                     suffix='[' + str(count) + '/' + str(l) + ']', length=50)
"""
write to file, display progress
"""
output_name = address_sheet.replace(" ", "_")
with open(output_name + '_formatted_with_lat_long.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["address", "latitude", "longitude"])
    count = 0
    for entry in geocoded:
        count += 1
        l = len(geocoded)
        printProgressBar(count, l, prefix='Writing Addresses:  ',
                         suffix='[' + str(count) + '/' + str(l) + ']', length=50)
        # because it's more satisfying to see the bar fill ...
        time.sleep(0.009)
        writer.writerow(entry)
print("Job's done! Saved data to \"" +
      output_name + '_formatted_with_lat_long.csv\"')
