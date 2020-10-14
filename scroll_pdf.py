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
    
    def pdf_to_jpg(self, file):
        images = pi.convert_from_path(file, poppler_path = self.pp)
        min_shape = sorted( [(np.sum(i.size), i.size ) for i in images])[0][1]
        page_break = np.array([[np.array([[0,0,0]]*min_shape[0],dtype=np.uint8)]]*10).squeeze()
        image = np.vstack([np.vstack([np.asarray(i.resize(min_shape)), page_break]) for i in images])
        image = PIL.Image.fromarray( image)
        return(image)
    
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
    
    pdf = pdf_canvas(root, file, zoom = 5)
    p1,_ = pdf.get_canvas()
    p1.pack(side = tk.TOP, expand=True, fill=tk.BOTH)
    
    def f_file(self):
        print("change...")
        p1,p2 = pdf.change_canvas(file2)
        return(p1,p2)
    
    root.bind("<Return>",f_file)
    
    def click(event):
        print(pdf)
        pdf.change_size(pdf.zoom+1)
        print(pdf)

    root.bind("<Button-1>", click)

    root.mainloop()