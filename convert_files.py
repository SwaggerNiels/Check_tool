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
from scroll_pdf import save_jpg

def sstr(s):
    return s.encode('ascii', errors='ignore').decode('utf-8')

class convert_pdfs():
    def __init__(self, path):
        self.path = path
        self.names = []
    
    def check_if_dir_exists(self):
        if os.path.isdir(self.path):
            if input(f"You want to remake '{self.path}'? (yes/no):\n") == "yes":
                rmtree(self.path)
            else:
                return True
                
        if os.path.isfile(self.path+".zip"):
            self.unzip_files()
            return True
        else:
            sys.exit(f"File {self.path}.zip not found")
    
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
            lastname = sstr(lastname).lower()
            print(lastname)
        
            try:
                py_file = list(filter(lambda x : x.endswith(".py"), filenames))[0]
                py_files.append([lastname,dirpath,py_file])
                # print(py_file)
            except IndexError:
                print(f"\tpython file not found for {name}")   
            
            try:
                pdf_file = list(filter(lambda x : x.endswith(".pdf"), filenames))[0]
                if pdf_file != "answers.pdf":
                    print("answers found in wrong directory - move this file 1 directory up")
                    pdf_files.append([lastname,dirpath,pdf_file])
                # print(pdf_file)
            except IndexError:
                print(f"\tpdf file not found for {name}")   
        
        #make sure spaces have priority
        for i,files in enumerate([py_files, pdf_files]):
            j = 0
            for name,_,_ in files:
                if i == 0: #py_files
                    py_files[j][0] = py_files[j][0].replace(' ',"AAA")
                if i == 1: #pdf_files
                    pdf_files[j][0] = pdf_files[j][0].replace(' ',"AAA")
                j+=1
        #sort
        py_files.sort(key = lambda x: x[0])
        pdf_files.sort(key = lambda x: x[0])
        print(py_files)
        print(pdf_files)
        #remove AAA's for spaces
        for i,files in enumerate([py_files, pdf_files]):
            j = 0
            for name,_,_ in files:
                if i == 0: #py_files
                    py_files[j][0] = py_files[j][0].replace("AAA",' ')
                if i == 1: #pdf_files
                    pdf_files[j][0] = pdf_files[j][0].replace("AAA",' ')
                j+=1
        
        for name,dirpath,file in py_files:
            old_filename = dirpath+"\\"+file
            new_filename = self.path+"\\"+name+".py"
            if os.path.isfile(new_filename):
                inplace_file_dirpath = [dirpath for inplace_name,dirpath,_ in py_files if inplace_name == name][0]
                if (input(f"{new_filename} already exists. You want to change file from:\n{inplace_file_dirpath}\nto\n{dirpath}? (yes/no)\n") == "yes"):
                    os.remove(new_filename)
                else:
                    continue
            
            os.rename(old_filename,new_filename)
            pass
            
        for name,dirpath,file in pdf_files:
            old_filename = dirpath+"\\"+file
            new_filename = self.path+"\\"+name+".pdf"
            if os.path.isfile(new_filename):
                inplace_file_dirpath = [dirpath for inplace_name,dirpath,_ in pdf_files if inplace_name == name][0]
                if (input(f"{new_filename} already exists. You want to change file from:\n{inplace_file_dirpath}\nto\n{dirpath}? (yes/no)\n") == "yes"):
                    os.remove(new_filename)
                else:
                    continue
            else:
                self.names.append(name)
            
            os.rename(old_filename,new_filename)
            
    
    def check_files(self):
        removal = []
        for name in self.names:
            if (not os.path.isfile(self.path+"\\"+name+".pdf") 
                or not os.path.isfile(self.path+"\\"+name+".py") 
                or not os.path.isfile(self.path+"\\"+name+".jpg")):
                
                print("NOTE: incomplete files\n" +name+ " has not completed their hand-in")
                removal.append(name)
                
        for name in removal:
            self.names.remove(name)
        
        if removal == [] and self.names != []:
            print("All files complete")
            return True
        else:
            return False
    
    def convert(self):
        if not self.check_files():
            self.check_if_dir_exists()
            self.make_jpgs()
            self.check_files()
        else:
            gen_obj = os.walk(self.path);
            
            for a in gen_obj:
                (dirpath,dirnames,filenames) = a
                for filename in filenames:
                    if filename[-3:] == "pdf":
                        name = filename[:-4]
                        self.names.append(sstr(name))
            self.names.sort()
            print("Files were already converted")
        
        return(self.names)
    
    def make_jpgs(self):
        for name in self.names:
            if not os.path.isfile(self.path+"\\"+name+".jpg"):
                save_jpg(self.path+"\\"+name+".pdf")
        print("Finished converting to jpg")
    
if __name__ == "__main__":
    path = r"C:\Users\niels\OneDrive\OneDriveDocs\TA\Numerical Methods\student results\week3\Assignment 2 Download 21 September, 2020 1622"
    obj = convert_pdfs(path)
    names = obj.convert()
    print(names)
    
    
    
    
    