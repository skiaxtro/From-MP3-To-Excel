# -*- coding: utf-8 -*-
"""
Created on Sat Apr 16 17:43:37 2022

@author: Stylianos Eleftheriadis
"""
import os
from pandas import DataFrame
from mutagen.easyid3 import EasyID3

#path = "C:\\..\\..\\..\\<folder name>" # No input Method

# Input Method
path = input("Please enter a valid folder path\nFor example (Windows): C:\\\..\\\..\\\..\\\<name of folder>\n\n")
print(f'\n\nYou entered:\n{path}')
os.chdir(path)

# Set Excel's name from keyboard
name = input("Set the Excel's name:\n\n")
print(f'\n\nYou entered:\n{name}')

# Read every file in the folder
def read_folder(directory):
    
    files_name = []  

    for root, directories, files in os.walk(directory):
        for filename in files:
            filepath = os.path.join(filename)
            files_name.append(filepath)  

    return files_name  

folder = read_folder(path)
length = len(folder)

tick = [""] * length  # I need it to "tick" every song I played. Remove it if you dont want it!
number = list(range(1, length+1))
title = []
artist = []

for i in range(length):
    audio = EasyID3(folder[i])
    
    title.append(audio['title'][0])
    artist.append(audio['artist'][0])

# Create data frame
df = DataFrame({'N': tick, '#': number,'SONG': title,'ARTIST': artist})

# Export data frame to excel
df.to_excel(name+'.xlsx', sheet_name=name, index=False)

print("\n\nThe Excel file was added to the folder's path!\n\n\n")