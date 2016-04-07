#!/usr/bin/python
# -*- coding: iso-8859-15 -*-
import sys
# IO package für Pfadnamen
stflist = ['C:\\Program Files\\Stimfit 0.14\\wx-3.0-msw', 
           'C:\\Program Files\\Stimfit 0.14', 
           'C:\\Program Files\\Stimfit 0.14\\stf-site-packages', 
           'C:\\WINDOWS\\SYSTEM32\\python27.zip', 
           'C:\\Users\\c-sch_000\\Anaconda\\Lib', 
           'C:\\Users\\c-sch_000\\Anaconda\\DLLs', 
           'C:\\Python27\\Lib', 
           'C:\\Python27\\DLLs', 
           'C:\\Python27\\Lib\\lib-tk', 
           'C:\\Python27', 
           'C:\\Python27\\lib\\site-packages']
sys.path = list(set(sys.path + stflist))
from Tkinter import *
import tkMessageBox as box
import tkFileDialog 
import stfio
import tkSimpleDialog 

# Imports for Export to Excel
import openpyxl
from openpyxl import Workbook
from openpyxl.compat import range
from openpyxl.cell import get_column_letter
from openpyxl.styles import Font, Fill

# Import for CSV file reading
import csv


# Imports for Peak Detection and Calculations
import numpy as np
from math import pi, log
import pylab
from scipy import fft, ifft, stats
from scipy.optimize import curve_fit

# Imports for Matplotlib
import matplotlib.pyplot as plt

# Import for directory settings

import os
from time import strftime

rec = stfio.read("14122000.abf")
last_value = 0
complete_dataset = np.array([0,0])
rig = 0
abffile = "None"
iAP_file = "None"
input_resistance = 0
capacitance = 1000
openDirectory = "C:\\"
saveDirectory = "C:\\temp\\"

#Create Excel file
wb = Workbook()
#dest_filename = saveDirectory + "\\"  + strftime("%Y-%m-%d_%H-%M-%S") + ".xlsx"
dest_filename = "F:\\Programmierung\\Python\\"
dest_directory = ""

ws1 = wb.active

# Filling of Column A with Descriptions
ws1['A1'] = "Well"
ws1['A5'] = "A1"
ws1['A6'] = "A2"
ws1['A7'] = "A3"
ws1['A8'] = "A4"
ws1['A9'] = "A5"  
ws1['A10'] = "A6"
ws1['A11'] = "B1"
ws1['A12'] = "B2"
ws1['A13'] = "B3"
ws1['A14'] = "B4"
ws1['A15'] = "B5"
ws1['A16'] = "B6"
ws1['A17'] = "C1"
ws1['A18'] = "C2"
ws1['A19'] = "C3"
ws1['A20'] = "C4"
ws1['A21'] = "C5"
ws1['A22'] = "C6"
ws1['A23'] = "D1"
ws1['A24'] = "D2"
ws1['A25'] = "D3"
ws1['A26'] = "D4"
ws1['A27'] = "D5"
ws1['A28'] = "D6"

ws1['B1'] = "Electrodes with spikes"

ws1['C1'] = "Electrodes without spikes"

ws1['D1'] = "Mean Spike frequency of electrodes with spikes [Hz]"


        


class Example(Frame):
  
    def __init__(self, parent):
        Frame.__init__(self, parent)   
         
        self.parent = parent        
        self.initUI()
        
    def initUI(self):
      
        self.parent.title("MEA analysis file summarizer for the PJK lab")
        self.pack(fill=BOTH, expand=1)
        menubar = Menu(self.parent)
        self.parent.config(menu=menubar)
        
        fileMenu = Menu(menubar, tearoff=0)
        fileMenu.add_command(label="Exit", command=self.onExit)    
        menubar.add_cascade(label="File", menu=fileMenu)     
        
        optionMenu = Menu(menubar, tearoff=0)
        optionMenu.add_command(label="Set directory for generated excel files", command=self.askdirectorySave)
        menubar.add_cascade(label="Options", menu=optionMenu)           

        optionMenu = Menu(menubar, tearoff=0)
        optionMenu.add_command(label="Read Dose response csv file", command=self.onMEADoseResponse)
        menubar.add_cascade(label="Analysis", menu=optionMenu)               
        
        helpMenu = Menu(menubar, tearoff=0)
        helpMenu.add_command(label="Help", command=self.onHelp)
        helpMenu.add_command(label="About", command=self.onAbout)    
        menubar.add_cascade(label="?", menu=helpMenu)              

    
        #self.txt = Text(self)
        #self.txt.pack(fill=BOTH, expand=1)
        
    def askdirectory(self):
        global openDirectory, saveDirectory, dest_filename, dest_directory
        """Returns a selected directoryname."""
        openDirectory = tkFileDialog.askdirectory()
        saveDirectory = openDirectory
        dest_directory = openDirectory
        print dest_directory
        
    def askdirectorySave(self):
        global saveDirectory, dest_filename, dest_directory
        """Returns a selected directoryname."""
        saveDirectory = tkFileDialog.askdirectory()  
        dest_directory  = saveDirectory
        #dest_filename = saveDirectory +  "\\"  + abffile[0] + ".xlsx"
        print dest_directory

    def onExit(self):
        root.destroy()
        
    def onAbout(self):
        box.showinfo("About MEA analysis file summarizer", "Version 0.1, April 2016\n\n Copyright: Christian Schnell (cschnell@schnell-thiessen.de)\n\n https://github.com/schrist81/MEAAnalyzer") 
    
    def onHelp(self):
        pass     
    

        


    def onMEADoseResponse(self):
        ftypes = [('comma separated files', '*.csv'), ('All files', '*')]
        dlg = tkFileDialog.Open(self, filetypes = ftypes, initialdir = openDirectory)
        fl = dlg.show()
        
        with open(fl, 'rb') as csvfile:
            #https://docs.python.org/2/library/csv.html
            csvreader = csv.DictReader(csvfile)
            spikeRate = []
            
            for row in csvreader:
                spikeRate.append(row['Spike Rate [Hz]'])

        numberOfElectrodes = 12
        numberOfWells = 24
        emptyElectrodes = 0
        spikeFrequency = 0 # in Hz
        for k in xrange(numberOfWells):
            for l in xrange(numberOfElectrodes):
                #print str(l+1) + ": " + str(k*numberOfElectrodes+l+1)
                if float(spikeRate[k*numberOfElectrodes+l]) == 0:
                    emptyElectrodes = emptyElectrodes + 1
                else:
                    spikeFrequency = spikeFrequency+float(spikeRate[k*numberOfElectrodes+l])
            fieldforSavingC = "C"+str(5+k)
            fieldforSavingB = "B"+str(5+k)
            ws1[fieldforSavingC] =  emptyElectrodes
            ws1[fieldforSavingB] =  numberOfElectrodes-emptyElectrodes

            fieldforSavingD = "D"+str(5+k)
            ws1[fieldforSavingD] = spikeFrequency/(numberOfElectrodes-emptyElectrodes)

            spikeFrequency = 0
            emptyElectrodes = 0

            

        dest_filename = dest_directory + "\\SpikeAnalysis.xlsx"
        wb.save(filename = dest_filename)




root = Tk()
ex = Example(root)


root.geometry("300x250+300+300")
root.mainloop()
