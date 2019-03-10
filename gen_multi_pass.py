#!/usr/bin/env python3
"""
  MultiPassword generator
  Generates 200 passwords With 
  8 characters one capital letter one special character
  and one integer
  11/27/2017

Copyright (C) 2019  Louis Scianni

    This program is free software; you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation; either version 2 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License along
    with this program; if not, write to the Free Software Foundation, Inc.,
    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.

    lscianniit@gmail.com

 
"""

import openpyxl
import random
from sys import argv



def make_pass():
    spreadsheet = argv[1]                                                               # get our spreadsheet file from commandline argunemt
    spchar = ['!', '@', '#', '$', '%', '^', '&', '*', '/', '\\', '<', '>', '(', ')', '+', '_', '-', '=']                                   # an array of special characters
                                                       
    
    wb = openpyxl.load_workbook(spreadsheet)   # open the file                          
    sheet = wb.get_sheet_by_name('Sheet1')     # select the sheet we are going to use
    
    with open('passwordlist.csv', 'a') as password_list:
        for col in sheet.iter_cols(min_row=1, max_row=200):    # iterate over the col in the sheet
            for cell in col:                                   # for each cell in the column grab the value of the cell
               words = cell.value
               cap_words = words.title() #words.replace(words[0], words[0].upper())            # capitilize the first letter
               sp_char = spchar[random.randint(0, 7)]                           # add a special character
               rand_digit = random.randint(0, 9)                                # select a random digit
               format_string = '%s%s%d\n' % (cap_words, sp_char, rand_digit)
               password_list.write(format_string)                               # format and print that to the screen
               
    with open('passwordlist.csv', 'r') as password_list:
        lines = password_list.readlines()
        for line in lines:
            print(line)
    
make_pass()  # call our function
