#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Geocoding script made by William Beck-Askenaizer for WQTS 
A little hacky, but it'll do.

7/9/2020
"""
import time
import sys
import csv
import subprocess
import os
import platform

"""
Make sure packages are installed
"""
def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

install_list = ['pandas', 'geopy', 'xlrd', 'commonregex']
reqs = subprocess.check_output([sys.executable, '-m', 'pip', 'freeze'])
installed_packages = [r.decode().split('==')[0] for r in reqs.split()]
for package in install_list:
    if package not in installed_packages:
        install(package)

import pandas as pd
from geopy.geocoders import ArcGIS, Bing, Nominatim, OpenCage, GoogleV3, OpenMapQuest
from commonregex import CommonRegex
from tkinter import filedialog
import tkinter as tk

"""
Colors for text formatting
"""

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    CYAN  = "\033[1;36m"


"""
clear the window :)
"""


def clear_terminal():
    os.system('cls') if platform.system() == "Windows" else os.system('clear')


clear_terminal()

"""
Flashy :)
"""


def print_logo():
    print(bcolors.OKBLUE +
          "\t **       **   *******    **********  ********" + bcolors.ENDC)
    print(bcolors.OKBLUE +
          "\t/**      /**  **/////**  /////**///  **////// " + bcolors.ENDC)
    print(bcolors.OKBLUE +
          "\t/**   *  /** **     //**     /**    /**       " + bcolors.ENDC)
    print(bcolors.OKBLUE +
          "\t/**  *** /**/**      /**     /**    /*********" + bcolors.ENDC)
    print(bcolors.OKBLUE +
          "\t/** **/**/**/**    **/**     /**    ////////**" + bcolors.ENDC)
    print(bcolors.OKBLUE +
          "\t/**** //****//**  // **      /**           /**" + bcolors.ENDC)
    print(bcolors.OKBLUE +
          "\t/**/   ///** //******* **    /**     ******** " + bcolors.ENDC)
    print(bcolors.OKBLUE +
          "\t//       //   /////// //     //     ////////  " + bcolors.ENDC)
    print(bcolors.OKBLUE + "\n")
    print(bcolors.OKBLUE +
          "\t Geolocation Script for Address Batches\n\n" + bcolors.ENDC)


"""
 State codes for input check
"""
abbr = ["AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA",
        "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD",
        "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
        "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC",
        "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY"]

"""
begin!
"""
clear_terminal()
print_logo()

""" 
initialize geocoders (just using ArcGIS for now)
"""
arcgis = ArcGIS(timeout=100)
nominatim = Nominatim(user_agent="WQTS", timeout=100)
# opencage = OpenCage('your-API-key', timeout=100)
openmapquest = OpenMapQuest('vooH8ziES69RZKTpR4LLUyuImVpaSY78')
geolocator = Nominatim(user_agent="WQTS")

geocoders = [openmapquest, arcgis, nominatim]

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


def check_for_address_format(addresses):
    final = []
    for item in addresses:
        parsed = CommonRegex(str(item))
        if parsed.street_addresses:
            final.append(parsed)
    return len(final)


"""
get list of columns
for each column in list
check the first 5 entries
whichever column has the most address-like entries is returned
"""


def auto_determine_address_col(working_sheet):
    tie = []
    addr_col = ""
    max_addrs = 0
    for column in working_sheet:
        addrs = []
        for entry in working_sheet[column].dropna().head(8):
            addrs.append(entry)
        if(check_for_address_format(addrs) > 0):
            tie.append(column)
            max_addrs = len(addrs)
            addr_col = column
    if max_addrs == 0:
        print(bcolors.FAIL + "Could not determine address column." + bcolors.ENDC)
        return False
    if len(tie) > 1:
       print(bcolors.WARNING + "warning: multiple address columns detected" + bcolors.ENDC)
    if len(tie) == 1:
        print(bcolors.OKGREEN + "Found column of addresses \"" + addr_col + "\"" + bcolors.ENDC)
        return addr_col


def get_list_from_file():
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

    address_num = input(bcolors.CYAN + 
        "Enter the number next to the desired sheet in the list above (e.g. 1 for " + str(base_file.sheet_names[0]) + "): " + bcolors.ENDC)
    address_sheet = base_file.sheet_names[int(address_num) - 1]
    print("Selected sheet:", address_sheet)

    header_sheet = base_file.parse(address_sheet, index_col=0)
    sheet_columns = header_sheet.columns.to_list()
    col_count = 1
    df = pd.ExcelFile(file_path).parse(address_sheet)
    auto_res = auto_determine_address_col(df)
    if auto_res:
        addr_set = set(df[auto_res].dropna().to_list())
    else:
        print("Available columns in '", address_sheet, "' are: \n")
        for column in sheet_columns:
            if("Unnamed" not in column):
                print(col_count, ":", column)
                col_count += 1
        address_column_num = input(bcolors.CYAN + 
            "Enter the number next to the column containing the addresses: " + bcolors.ENDC)
        address_column = sheet_columns[int(address_column_num)-1]
        addr_set = set(df[address_column].dropna().to_list())
        print("Selected column:", address_column, "\n")

    if check_for_address_format(addr_set) < len(addr_set)/4:
        print(bcolors.WARNING +
              "warning: Less than 25% of the entries in the column appear to be addresses." + bcolors.ENDC)
        proceed = input(bcolors.CYAN + "Proceed? (y/n): " + bcolors.ENDC)
        if proceed == "y":
            print(bcolors.WARNING + "warning: The script will still try, but makes no guarantees of the results. You should check the output manually." + bcolors.ENDC)
        if proceed == "n":
            print(bcolors.WARNING + "exiting..." + bcolors.ENDC)
            exit()
    return addr_set, address_sheet


def get_geolocation_data(batch):
    print("Geocode function received:", len(batch), "items to search for")
    count = 0
    city = input(bcolors.CYAN + "Enter the city " + bcolors.UNDERLINE +
                 "(spelling and spacing matter!): " + bcolors.ENDC + " ").title()
    state = input(bcolors.CYAN +  "Enter the state (Texas, California, etc.): " + bcolors.ENDC).title()
    print("\n")
    failed = []
    coded_batch = [[]]
    l = len(batch)
    for address in batch:
        address = str(address) + ', ' + city + ', ' + state
        try:
            printProgressBar(count, l, prefix='Geocoding Addresses: ',
                             suffix='[' + str(count) + '/' + str(l) + ']', length=50)
            for geocoder in geocoders:
                full_location = geocoder.geocode(address)
                if full_location != None:
                    coded_batch.append(full_location)
                    break
            count += 1
        except:
            failed.append(address)
    return coded_batch, failed


def write_to_file(coded_batch, sheet_name):
    output_name = sheet_name.replace(" ", "_")
    with open(output_name + '_formatted_with_lat_long.csv', 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["address", "latitude, longitude"])
        count = 0
        try:
            for entry in coded_batch:
                count += 1
                l = len(coded_batch)
                printProgressBar(count, l, prefix='Writing Addresses:   ',
                                 suffix='[' + str(count) + '/' + str(l) + ']', length=50)
                # because it's more satisfying to see the bar fill ...
                time.sleep(0.009)
                writer.writerow(entry)
        except:
            pass

    print("Job's done! Saved data to " +
          bcolors.OKBLUE + "\"" + output_name + '_formatted_with_lat_long.csv\"' + bcolors.ENDC)


def main():
    address_list, sheet_name = get_list_from_file()
    coded, failed = get_geolocation_data(address_list)
    write_to_file(coded, sheet_name)
    print(bcolors.OKGREEN + "Found geolocation data for all entries" + bcolors.ENDC) if len(failed) == 0 else print(
        bcolors.FAIL + str(len(failed)), "addresse(s)"+ str(failed) +" are probably on Mars. (No data)." + bcolors.ENDC)
    if len(failed) > len(address_list)/2:
        print("Over half of the addresses could not be found. Did you specify the right column?")
main()

def run():
    main()
