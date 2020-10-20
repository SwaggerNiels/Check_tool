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
from scroll_pdf import save_jpg,save_pycode_jpg
import matplotlib.pyplot as plt

def sstr(s):
    return s.encode('ascii', errors='ignore').decode('utf-8')

class convert_pdfs():
    def __init__(self, path, py_files = None):
        self.path = path
        self.names = []
        self.py_files = py_files #contains tags of the py files that are needed
    
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
        _,upper_dirnames,_ = next(gen_obj)
        print(upper_dirnames)
        
        submissions = dict()
        for dirname in upper_dirnames:
            name = re.findall(r" - ([\w -]*) -",dirname)[0]
            lastname = re.findall(r"[^\ ]* (.*)", name)[0]
            lastname = sstr(lastname).lower()
            
            if lastname not in submissions.keys():
                submissions[lastname] = 1
            else:
                submissions[lastname] += 1
        print(submissions)
        
        py_files = [];
        pdf_files = [];
        
        if self.py_files != None:
            print(f"Multiple py files will be searched with tags: {self.py_files}")
        
        number_of_files_found = dict()
        
        for a in gen_obj:
            (dirpath,dirnames,filenames) = a
            
            name = re.findall(r" - ([\w -]*) -",dirpath)[0]
            lastname = re.findall(r"[^\ ]* (.*)", name)[0]
            lastname = sstr(lastname).lower()
            submissions[lastname] -= 1
            
            if lastname not in number_of_files_found.keys():
                number_of_files_found[lastname] = 0
                print(lastname)
            
            if self.py_files == None: #single python file
                try:
                    py_file = list(filter(lambda x : x.endswith(".py"), filenames))[0]
                    py_files.append([lastname,dirpath,py_file])
                    # print(py_file)
                except IndexError:
                    print(f"\tpython file not found for {lastname}")   
            else:
                for tag in self.py_files:
                    try:
                        files_found = list(filter(lambda x : x.endswith(".py"), filenames))
                        py_file = list(filter(lambda x : tag in x.lower(), files_found))[0]
                        
                        py_files.append([lastname,dirpath,py_file,tag])
                        number_of_files_found[lastname] += 1
                        # print(py_file)
                    except IndexError:
                        print(f"\tpython file {tag} not found for {lastname}") 
                
                print(number_of_files_found[lastname])
                #check if all files found
                while number_of_files_found[lastname] < len(self.py_files):
                    if submissions[lastname] > 0:
                        break
                    for tag in self.py_files:
                        print(f"Wrong format detected in {lastname}")
                        files_found = list(filter(lambda x : x.endswith(".py"), filenames))
                        files_found = list(filter(lambda x : tag in x.lower(), files_found))
                        
                        if files_found == []:
                            print(f"No files found... {tag} --> create blank file for structure:")
                            blank_path = self.path+"\\"+ lastname + '_' + tag + ".py"
                            try:
                                with open(blank_path, "w+"):
                                    pass
                            except:
                                print(f"Couldn't create blank file {blank_path}")
        
                            print(f"{blank_path}")
                            
                            py_files.append([lastname,dirpath,py_file,tag])
                            number_of_files_found[lastname] += 1
                            continue
                            
                        selection_string = "\n".join([str(i+1) + ") " + files_found[i] for i in range(len(files_found))])
                        index = int(input(f"Which is {tag}? (give number) \n{selection_string}\n")) - 1
                        py_file = files_found[index]
                    
                        selection_string = "\n".join([str(i+1) + ") " + self.py_files[i] for i in range(len(self.py_files))])
                        index = int(input(f"For which py file tag? (give number) \n{selection_string}\n")) - 1
                        py_files.append([lastname,dirpath,py_file,tag])
                        number_of_files_found[lastname] += 1
                
                
            try:
                pdf_file = list(filter(lambda x : x.endswith(".pdf"), filenames))[0]
                if pdf_file == "answers.pdf":
                    print("answers found in wrong directory - move this file 1 directory up")
                else:
                    pdf_files.append([lastname,dirpath,pdf_file])
                # print(pdf_file)
            except IndexError:
                print(f"\tpdf file not found for {lastname}")   
        
        #make sure spaces have priority
        for i,files in enumerate([py_files, pdf_files]):
            j = 0
            for name,*_ in files:
                if i == 0: #py_files
                    py_files[j][0] = py_files[j][0].replace(' ',"AAA")
                if i == 1: #pdf_files
                    pdf_files[j][0] = pdf_files[j][0].replace(' ',"AAA")
                j+=1
        #sort
        py_files.sort(key = lambda x: x[0])
        pdf_files.sort(key = lambda x: x[0])
        # print(py_files)
        # print(pdf_files)
        
        #remove AAA's for spaces
        for i,files in enumerate([py_files, pdf_files]):
            j = 0
            for name,*_ in files:
                if i == 0: #py_files
                    py_files[j][0] = py_files[j][0].replace("AAA",' ')
                if i == 1: #pdf_files
                    pdf_files[j][0] = pdf_files[j][0].replace("AAA",' ')
                j+=1
        
        if self.py_files == None: #single python file
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
        else:
            for tag in self.py_files:
                for name,dirpath,file,file_tag in py_files:
                    if file_tag == tag:
                        old_filename = dirpath+"\\"+file
                        new_filename = self.path+"\\"+name+'_'+tag+".py"
                        if os.path.isfile(new_filename):
                            inplace_file_dirpath = [dirpath for inplace_name,dirpath,_,_ in py_files if inplace_name == name][0]
                            if (input(f"{new_filename} '{tag}' already exists. You want to change file from:\n{inplace_file_dirpath}\nto\n{dirpath}? (yes/no)\n") == "yes"):
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
        
        if self.names == []:
            gen_obj = os.walk(self.path);
            
            for a in gen_obj:
                (dirpath,dirnames,filenames) = a
                for filename in filenames:
                    if filename[-3:] == "pdf" and filename[-11:] != "_pycode.pdf":
                        name = filename[:-4]
                        self.names.append(sstr(name))
            self.names.sort()
            print("Files were already converted")
        
        return(self.names)
    
    def make_jpgs(self):
        print("Converting ---------")
        for name in self.names:
            if not os.path.isfile(self.path+"\\"+name+".jpg"):
                print(name, "report...")
                save_jpg(self.path+"\\"+name+".pdf")
                print(name, "code...")
                if self.py_files == None: #single python file
                    save_pycode_jpg(self.path+"\\"+name+".py")
                else:
                    for tag in self.py_files:
                        save_pycode_jpg(self.path+"\\"+name+'_'+tag+".py")
                
        print("Finished converting to jpg")
    
if __name__ == "__main__":
    path = r"C:\Users\niels\OneDrive\OneDriveDocs\TA\Numerical Methods\student results\assignment3\Assignment 3 Download 15 October, 2020 1511"
    obj = convert_pdfs(path)
    names = obj.convert()
    print(names)
    
    
    
    
    