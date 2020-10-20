# -*- coding: utf-8 -*-
"""
Created on Tue Sep 22 20:30:34 2020

@author: niels
"""
import PIL
from PIL import Image,ImageTk
import tkinter as tk
import pdf2image as pi
import numpy as np
import os
import code2pdf as cp

PP = r"C:\Users\niels\Desktop\release\poppler-0.90.1\bin"

def pdf_to_jpg(file):
    images = pi.convert_from_path(file, poppler_path = PP)
    min_shape = sorted( [(np.sum(i.size), i.size ) for i in images])[0][1]
    page_break = np.array([[np.array([[0,0,0]]*min_shape[0],dtype=np.uint8)]]*10).squeeze()
    image = np.vstack([np.vstack([np.asarray(i.resize(min_shape)), page_break]) for i in images])
    image = PIL.Image.fromarray( image)
    return(image)
  
def save_jpg(file):
    image = pdf_to_jpg(file)
    print(file[:-3]+"jpg", "made")
    image.save(file[:-3]+"jpg")
    return(file[:-3]+"jpg")

def concate_image_array(images, page_break_height, cut):
    min_shape = sorted( [(np.sum(i.size), i.size ) for i in images])[0][1]
    # print(min_shape)
    # [print(i.size) for i in images]
    # page_break = np.array([[np.array([[0,0,0]]*min_shape[0],dtype=np.uint8)]]*10).squeeze()
    page_break = np.array([[np.array([[0,0,0]]*(min_shape[0]-2*cut[0]),dtype=np.uint8)]]*page_break_height).squeeze()
    # print("page break:",page_break.shape)
    # print("with breaks:")
    if cut[0] != 0 and cut[1] != 0:
        # [print(np.asarray(i.resize(min_shape)).shape) for i in images]
        # print("to")
        # [print(np.asarray(i.resize(min_shape))[cut[1]:-cut[1],cut[0]:-cut[0],:].shape) for i in images]
        images = np.vstack([np.vstack([np.asarray(i.resize(min_shape))[cut[1]:-cut[1],cut[0]:-cut[0],:], page_break]) for i in images])
    else:
        images = np.vstack([np.vstack([np.asarray(i.resize(min_shape)), page_break]) for i in images])
    return images

def pypdf_to_jpg(file, page_break_height=10, cut=(0,0)):
        images = pi.convert_from_path(file, poppler_path = PP)
        image = concate_image_array(images,page_break_height,cut)
        image = PIL.Image.fromarray(image)
        return(image)

def py_to_jpg(file):
        postfix = "_pycode"
        if cp.main(file,postfix) == 0:
            return pypdf_to_jpg(file[:-3]+postfix+".pdf", 2, (250,200))          

def save_pycode_jpg(file):
    image = py_to_jpg(file)
    print(file[:-3]+"_pycode.jpg", "made")
    image.save(file[:-3]+"_pycode.jpg")
    return(file[:-3]+"_pycode.jpg")
    
class pdf_canvas():
    def __init__(self, root, file, zoom = 6, poppler_path = r"C:\Users\niels\Desktop\release\poppler-0.90.1\bin"):
        self.root = root
        self.file = file
        self.zoom = zoom
        self.pp = poppler_path
        
        self.func_image = self.pdf_to_jpg
        self.func_scroll = None
        
        self.canvas = None
        self.scroll_bar = None
        self.image = None
        self.photo = None
        self.image_on_canvas = None
        
    def pdf_to_jpg(self, file, page_break_height=10, cut=(0,0)):
        images = pi.convert_from_path(file, poppler_path = self.pp)
        image = concate_image_array(images,page_break_height,cut)
        image = PIL.Image.fromarray(image)
        return(image)
    
    def py_to_jpg(self, file, page_break_height=10, cut=(0,0)):
        return py_to_jpg(file)
    
    def get_jpg(self, file):
        return (PIL.Image.open(file[:-3]+"jpg"))
    
    def get_canvas(self):
        self.scroll_bar = tk.Scrollbar(self.root, orient=tk.VERTICAL) 
        self.scroll_bar.pack(side = tk.RIGHT, fill = tk.Y ) 
        
        self.image = self.func_image(self.file)
        
        self.basewidth = 150*self.zoom
        self.canvas = tk.Canvas(self.root, height=1000, width=self.basewidth)
        self.wpercent = (self.basewidth / float(self.image.size[0]))
        
        self.hsize = int((float(self.image.size[1]) * float(self.wpercent)))
        self.image = self.image.resize((self.basewidth, self.hsize), PIL.Image.ANTIALIAS)
        self.photo = PIL.ImageTk.PhotoImage(self.image)
        self.image_on_canvas = self.canvas.create_image(0,0, anchor = tk.NW, image=self.photo)
        
        self.scroll_bar.config(command = self.canvas.yview ) 
        self.canvas.config(yscrollcommand=self.scroll_bar.set, scrollregion=(0,0,self.basewidth,self.hsize))
            
        self.func_scroll = lambda event: self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        self.root.bind_all("<MouseWheel>", self.func_scroll)
        
        return(self.canvas,self.photo)
    
    def __change_size(self, f):
        '''This function doesn't work yet'''
        self.zoom = f
        
        self.basewidth = self.zoom*150
        # self.basewidth = int(self.basewidth)
        self.canvas.config(width=self.basewidth)
        self.wpercent = (self.basewidth / float(self.image.size[0]))
        return(self.change_canvas(self.file))
    
    def change_canvas(self, file):
        self.file = file
        
        self.image = self.func_image(file)
        self.wpercent = (self.basewidth / float(self.image.size[0]))
        self.hsize = int((float(self.image.size[1]) * float(self.wpercent)))
        self.image = self.image.resize((self.basewidth, self.hsize), PIL.Image.ANTIALIAS)
        self.photo = PIL.ImageTk.PhotoImage(self.image)
        self.canvas.itemconfig(self.image_on_canvas, image = self.photo)
        self.canvas.config(yscrollcommand=self.scroll_bar.set, scrollregion=(0,0,self.basewidth,self.hsize))
        
        return(self.canvas, self.photo)
    
    def __str__(self):
        return(f'''self.basewidth = {self.basewidth}
              self.zoom = {self.zoom}
              self.wpercent = {self.wpercent}
              self.hsize = {self.hsize}
              ''')

if __name__ == "__main__":
    root = tk.Tk()
    
    path = "C:\\Users\\niels\\OneDrive\\OneDriveDocs\\TA\\Numerical Methods\\"
    file = path+"Assignment1-2020-answers.pdf"
    file2 = path+"Assignment2-2020-answers.pdf"
    file3 = path+"Assignment2-2020.py"
    
    # file2 = save_jpg(file2)
    # file3 = save_pycode_jpg(file3)
    file2 = file2[:-3]+"jpg"
    file3 = file3[:-3]+"_pycode.jpg"
    
    files = [file2,file3]
    index = False
    
    pdf = pdf_canvas(root, file3, zoom = 5)
    pdf.func_image = pdf.get_jpg
    
    p1,a = pdf.get_canvas()
    p1.pack(side = tk.TOP, expand=True, fill=tk.BOTH)
    
    # def f_file(self):
    #     print("change...")
    #     p1,p2 = pdf.change_canvas(file2)
    #     return(p1,p2)
    
    # root.bind("<Return>",f_file)
    
    def click(event):
        global index,pdf,p1,a
        
        print(pdf)
        index = not index
        print(index)
        file = files[int(index)]
        p1,a = pdf.change_canvas(file)
        print(pdf)

    root.bind("<Button-1>", click)

    root.mainloop()