# -*- coding: utf-8 -*-
"""
Created on Mon Sep 14 16:55:18 2020

@author: niels
"""
import numpy as np
import pandas as pd
import xlwings as xw

def c(col,row):
    abc = list("0ABCDEFGHIJKLMNO")
    return abc[col]+str(row)

class excel_feedback():
    def __init__(self, path, file):
        self.path = path
        self.file = file
        
        self.grade = None
        self.grader = None
        self.total_points = None
        self.df = None
        self.name_number = 22
        
        self._ask_info()
    
    def _ask_info(self):
        m = "What is the thing you are grading? (correctness/timeliness):\n"
        self.grade = input(m)
        
        m = "What do you want to be refered to as? (e.g.: 'TA Niels'):\n"
        self.grader = input(m)
        
        m = "What is the maximum points attainable? (e.g.: '21'):\n"
        self.total_points = input(m)
        
        
    def make_feedback(self):
        if len(xw.apps.keys()) > 0:
            try:
                excel_pid = xw.apps.keys()[0]
                wb = xw.apps[excel_pid].books(self.file)
            except:
                print("Error: make sure you close all other excel-sheets than scoresheet")
        else:
            wb = xw.Book(self.path+self.file)
        app = xw.apps.active    

        sheets = []
        for sheet in wb.sheets:
            if sheet.name != "template":
                sheets.append(sheet.name)
        
        names = wb.sheets[sheets[0]].range(c(1,3),c(1,self.name_number)).value
        while None in names:
            names.remove(None)
        
        self.feedback = pd.DataFrame(index = names, columns = ["feedback", "total"])
        
        self.feedback["feedback"] = ""
        self.feedback["total"] = 0
        
        #columns
        col_t = 10 #timeliness
        col_c = 11 #correctness
        col_r = 12 #remarks
        
        for sheet in sheets:
            data = wb.sheets[sheet]
            getval = lambda x,y : data.range(c(x,y)).value
            getrange = lambda x1,y1,x2,y2 : data.range(c(x1,y1),c(x2,y2)).value
            
            question_name = getval(1,1)
            points = getrange(col_c,3,col_c,2+len(names))
            remarks = getrange(col_r,3,col_r,2+len(names))
            options = getrange(3,1,col_t-1,1)
            deduction = getrange(3,2,col_t-1,2)
            ticks = getrange(3,3,col_t-1,2+len(names))
            
                        
            #get option (mistake) remarks
            df_options = dict()
            for i,name in list(enumerate(names)):
                for j,tick in list(enumerate(ticks[i])):
                    if tick == 'x':
                        if j == 0:
                            df_options[name] = options[j] + "\n"
                        elif name in df_options.keys():
                            df_options[name] = df_options[name] + \
                                options[j] + " --> -" + str(deduction[j]) + "\n"
                        else:
                            df_options[name] = options[j] + " --> -" + str(deduction[j]) + "\n"
                        
            df_i = pd.DataFrame(index = names)
            df_i["points"] = points
            df_i["remarks"] = remarks
            df_options = pd.DataFrame.from_dict(df_options, orient= "index",columns = ["options"])
            df_i = pd.merge(df_i, df_options, how='left', left_index=True, right_index=True)
            
            points = df_i["points"].apply(lambda x : round(x,3))
            remarks = df_i["remarks"].fillna('')
            options = df_i["options"].fillna('')
            
            # df_i["remarks"].replace(
            #     r"^0$", '', 
            #     regex=True).replace(
            #     r"^\. $", "feedback missing", 
            #     regex=True)
            
            # print(question_name)
            # print(points.tolist())
            # print(remarks.tolist())
            
            out_of = "/1"
            
            # print(question_name + ":\n" + remarks + "\n-->" + points.astype(str) + out_of + "\n\n")
            # print(len(question_name + ":\n" + remarks + "\n-->" + points.astype(str) + out_of + "\n\n"))
            
            self.feedback[sheet+"_feedback"] \
                = np.array(question_name + ":\n" + options + remarks + \
                           "\nsubtotal --> " + points.astype(str) + out_of + "\n\n")
            self.feedback[sheet+"_points"] = np.array(points)
            
            self.feedback["feedback"] = self.feedback["feedback"] + self.feedback[sheet+"_feedback"]
            self.feedback["total"] = self.feedback["total"] + self.feedback[sheet+"_points"]
        
        app.quit()
        return(self.feedback)
        # print(feedback)
        # print(feedback.loc[:,["feedback","total"]])

    def write_txts(self, path = ""):
        path = self.path if path == "" else path
        
        for person in self.feedback.iloc:
            name = person.name
            
            try:
                print(name+"'s \tfeedback generated")
                f = open(path + name + ".txt","w")
                
                all_feedback = person[0]
                f.write(all_feedback)
                
                points = str(round(person[1],3))
                f.write("Total:"+points+\
                        f"/{self.total_points} points {self.grade} - {self.grader}\n")
                f.write("____________________________")
            finally:
                f.close()

if __name__ == '__main__':
    path = r"C:\Users\niels\OneDrive\OneDriveDocs\TA\Numerical Methods\student results\week3" + "\\"
    file = "scores_week3.xlsx"
    
    fb = excel_feedback(path, file)
    df = fb.make_feedback()
    # fb.write_txts()