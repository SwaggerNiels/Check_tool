# -*- coding: utf-8 -*-
"""
Created on Sun Sep 13 14:27:14 2020

@author: niels
"""

import os
import sys
import re
import zipfile
from shutil import rmtree

class convert_pdfs():
    def __init__(self, path):
        self.path = path
        self.names = []
    
    def check_if_exists(self):
        if os.path.isdir(self.path):
            if input(f"Are you sure you want to delete '{self.path}'? (yes/no):\n") == "yes":
                rmtree(self.path)
            else:
                sys.exit(f"must overwrite {self.path}")
    
    def unzip_files(self):
        with zipfile.ZipFile(self.path+".zip", 'r') as zip_ref:
            zip_ref.extractall(self.path)
        
        gen_obj = os.walk(self.path);
        next(gen_obj)
        
        py_files = [];
        pdf_files = [];
        
        for a in gen_obj:
            (dirpath,dirnames,filenames) = a
            
            # print(dirpath)
            name = re.findall(r" - ([\w -]*) -",dirpath)[0]
            lastname = re.findall(r"[^\ ]* (.*)", name)[0]
            # print(lastname)
        
            try:
                py_file = list(filter(lambda x : x.endswith(".py"), filenames))[0]
                py_files.append((lastname,dirpath,py_file))
                # print(py_file)
            except IndexError:
                print(f"\tpython file not found for {name}")   
            
            try:
                pdf_file = list(filter(lambda x : x.endswith(".pdf"), filenames))[0]
                pdf_files.append((lastname,dirpath,pdf_file))
                # print(pdf_file)
            except IndexError:
                print(f"\tpdf file not found for {name}")   
        
        py_files.sort(key = lambda x: x[0])
        pdf_files.sort(key = lambda x: x[0])
        
        for name,dirpath,file in py_files:
            os.rename(dirpath+"\\"+file,self.path+"\\"+name+".py")
            pass
            
        for name,dirpath,file in pdf_files:
            os.rename(dirpath+"\\"+file,self.path+"\\"+name+".pdf")
            self.names.append(name)
    
    def check_files(self):
        removal = []
        for name in self.names:
            if not os.path.isfile(self.path+"\\"+name+".pdf") or not os.path.isfile(self.path+"\\"+name+".py"):
                print("NOTE: incomplete files\n" +name+ " has not completed their hand-in")
                removal.append(name)
                
        for name in removal:
            self.names.remove(name)
        
        if removal == []:
            print("All files complete")
    
    def convert(self):
        self.check_if_exists()
        self.unzip_files()
        self.check_files()
        
        return(self.names)
    
if __name__ == "__main__":
    path = r"C:\Users\niels\OneDrive\OneDriveDocs\TA\Numerical Methods\student results\week3\Assignment 2 Download 21 September, 2020 1622"
    obj = convert_pdfs(path)
    names = obj.convert()
    
    
    
    
    