# -*- coding: utf-8 -*-
"""
Created on Thu Oct  8 15:41:46 2020

@author: niels
"""

import pyperclip
import os

path = r"C:\Users\niels\OneDrive\OneDriveDocs\TA\Numerical Methods\student results" + "\\"
superfolder = r"intermediate_assignment4" + "\\" 
folder = r"feedback_files" + "\\"

location = path+superfolder+folder

file_obj = os.walk(location)

_,_,files = next(file_obj)

for file in files:
    print(file)
    with open(location+file, 'r') as f:
        pyperclip.copy(f.read())
    input("Next?")