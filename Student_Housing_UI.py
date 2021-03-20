#!/usr/bin/env python
# coding: utf-8

# In[12]:


import tkinter as tk
from tkinter import ttk
from tkinter import font
import pandas as pd
import numpy as np
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import h5py
from tabulate import tabulate

class MainApp(tk.Frame):
    def __init__(self,parent,*args,**kwargs):
        tk.Frame.__init__(self,parent, *args,**kwargs)
        self.label = tk.Label(self.master, text="Student Housing", font = ("Arial",25))
        self.label.grid(column = 0, row= 0)
        self.boton = tk.Button(self.master, text = "Enter Data:", command = self.new_window)
        self.boton.grid(column = 3, row =1)
        
    def new_window(self):
        self.newWindow = tk.Toplevel(self.master)
        self.app1 = App2(self.newWindow)
        
class App2:
    def __init__(self, master):
        self.master = master
        
    #initialize treeview
        self.inserttree()
        
    #Save & Export
        self.savebutton = tk.Button(self.master, text = "Save", command = self.saveinfo)
        self.savebutton.grid(column=4, row=36)
        self.exportbutton = tk.Button(self.master, text = "Export", command = self.export)
        self.exportbutton.grid(column=5, row=36)
        self.submitbutton = tk.Button(self.master, text = "Insert", command = self.insert_data)
        self.submitbutton.grid(column=6, row=36)
        self.deletebutton = tk.Button(self.master, text = "Delete", command = self.delete_data)
        self.deletebutton.grid(column=7, row=36)
        
    #Initialize variable classes
        self.choiceVar1 = tk.StringVar()
        self.choiceVar2 = tk.IntVar()
        self.choiceVar3 = tk.IntVar()
        self.choiceVar4 = tk.DoubleVar()
        self.choiceVar5 = tk.DoubleVar()
        self.choiceVar6 = tk.StringVar()
        self.choiceVar7 = tk.StringVar()
        self.choiceVar8 = tk.StringVar()
        self.choiceVar9 = tk.IntVar()
        self.choiceVar10 = tk.DoubleVar()
        self.choiceVar11 = tk.IntVar()
        self.choiceVar12 = tk.StringVar()
        self.choiceVar13 = tk.DoubleVar()
        self.choiceVar14 = tk.DoubleVar()
        self.choiceVar15 = tk.DoubleVar()
        self.choiceVar16 = tk.IntVar()
        self.choiceVar17 = tk.IntVar()
        self.choiceVar18 = tk.DoubleVar()
        self.choiceVar19 = tk.DoubleVar()
        self.choiceVar20 = tk.StringVar()
        self.choiceVar21 = tk.StringVar()
        self.choiceVar22 = tk.StringVar()
        self.choiceVar23 = tk.IntVar()
        self.choiceVar24 = tk.StringVar()
        self.choiceVar25 = tk.StringVar()
        self.choiceVar26 = tk.StringVar()
        self.choiceVar27 = tk.StringVar()
        self.choiceVar28 = tk.StringVar()
        self.choiceVar29 = tk.StringVar()
        self.choiceVar30 = tk.StringVar()
        self.choiceVar31 = tk.DoubleVar()
        self.choiceVar32 = tk.DoubleVar()
        self.choiceVar33 = tk.DoubleVar()
        
   #Attributes
        self.title0 = tk.Label(self.master, text="School Attributes:", font = 'Arial 10 bold')
        self.title0.grid(column = 0, row = 0)
        
        self.e1 = tk.Label(self.master, text = "University Name", font = ('Arial',8))
        self.e1.grid(column = 0, row = 1)
        self.entry1 = tk.Entry(self.master, textvariable = self.choiceVar1, width=20)
        self.entry1.grid(column = 1, row = 1)
        
        self.e2 = tk.Label(self.master, text = "Current Fall Enrollment", font = ('Arial',8))
        self.e2.grid(column = 0, row = 2)
        self.entry2 = tk.Entry(self.master, textvariable = self.choiceVar2, width=20)
        self.entry2.grid(column = 1, row = 2)
        
        self.e3 = tk.Label(self.master, text = "Prior Year Enrollment", font = ('Arial',8))
        self.e3.grid(column = 0, row = 3)
        self.entry3 = tk.Entry(self.master, textvariable = self.choiceVar3, width=20)
        self.entry3.grid(column = 1, row = 3)
        
        self.e4 = tk.Label(self.master, text = "Enrollment Change %", font = ('Arial',8))
        self.e4.grid(column = 0, row = 4)
        self.entry4 = tk.Entry(self.master, textvariable = self.choiceVar4, width=20)
        self.entry4.grid(column = 1, row = 4)
        
        self.e5 = tk.Label(self.master, text = "Acceptance Rate %", font = ('Arial',8))
        self.e5.grid(column = 0, row = 5)
        self.entry5 = tk.Entry(self.master, textvariable = self.choiceVar5, width=20)
        self.entry5.grid(column = 1, row = 5)
        
        
        self.e6 = tk.Label(self.master, text = "Campus Type", font = ('Arial',8))
        self.e6.grid(column = 0, row = 6)
        self.entry6 = ttk.Combobox(self.master, textvariable = self.choiceVar6,values = ["On-Campus", "Off-Campus"]) 
        self.entry6.grid(column = 1, row = 6)
       
        self.e7 = tk.Label(self.master, text = "Program Type", font = ('Arial',8))
        self.e7.grid(column = 0, row = 7)
        self.entry7 = ttk.Combobox(self.master, textvariable = self.choiceVar7,values = ["Undergraduate", "Graduate"]) 
        self.entry7.grid(column = 1, row = 7)
        
        
    #Financials
        self.title1 = tk.Label(self.master, text="Financials:", font = 'Arial 10 bold')
        self.title1.grid(column = 0, row = 8)

        self.e8 = tk.Label(self.master, text = "Debt Schedule", font = ('Arial',8))
        self.e8.grid(column = 0, row = 9)
        self.entry8 = ttk.Combobox(self.master, textvariable = self.choiceVar8,values = ["Aggressive", "Moderate", "Level"]) 
        self.entry8.grid(column = 1, row = 9)
        
        self.e9 = tk.Label(self.master, text = "MADS($):", font = ('Arial',8))
        self.e9.grid(column = 0, row = 10)
        self.entry9 = tk.Entry(self.master, textvariable = self.choiceVar9, width=20)
        self.entry9.grid(column = 1, row = 10)
        
        self.e10 = tk.Label(self.master, text = "Stabilized Coverage Ratio", font = ('Arial',8))
        self.e10.grid(column = 0, row = 11)
        self.entry10 = tk.Entry(self.master, textvariable = self.choiceVar10, width=20)
        self.entry10.grid(column = 1, row = 11)
        
        self.e11 = tk.Label(self.master, text = "Amount of Cap.I. ($):", font = ('Arial',8))
        self.e11.grid(column = 0, row = 12)
        self.entry11 = tk.Entry(self.master, textvariable = self.choiceVar11, width=20)
        self.entry11.grid(column = 1, row = 12)
        
        self.e12 = tk.Label(self.master, text = "Cap. I. End Date", font = ('Arial',8))
        self.e12.grid(column = 0, row = 13)
        self.entry12 = tk.Entry(self.master, textvariable = self.choiceVar12, width=20)
        self.entry12.grid(column = 1, row = 13)
        
        self.e13 = tk.Label(self.master, text = "Operating Reserve % of Expenses", font = ('Arial',8))
        self.e13.grid(column = 0, row = 14)
        self.entry13 = tk.Entry(self.master, textvariable = self.choiceVar13, width=20)
        self.entry13.grid(column = 1, row = 14)
        
        self.e14 = tk.Label(self.master, text = "Surplus Fund Amount ($):", font = ('Arial',8))
        self.e14.grid(column = 0, row = 15)
        self.entry14 = tk.Entry(self.master, textvariable = self.choiceVar14, width=20)
        self.entry14.grid(column = 1, row = 15)
        
        self.e15 = tk.Label(self.master, text = "Release Test Requirement", font = ('Arial',8))
        self.e15.grid(column = 0, row = 16)
        self.entry15 = tk.Entry(self.master, textvariable = self.choiceVar15, width=20)
        self.entry15.grid(column = 1, row = 16)
        
        self.e16 = tk.Label(self.master, text = "Ground Lease Payment($):", font = ('Arial',8))
        self.e16.grid(column = 0, row = 17)
        self.entry16 = tk.Entry(self.master, textvariable = self.choiceVar16, width=20)
        self.entry16.grid(column = 1, row = 17)
        
        self.e17 = tk.Label(self.master, text = "Ground Lease Payment Start Date", font = ('Arial',8))
        self.e17.grid(column = 0, row = 18)
        self.entry17 = tk.Entry(self.master, textvariable = self.choiceVar17, width=20)
        self.entry17.grid(column = 1, row = 18)
        
     #Project Attributes
        self.title2 = tk.Label(self.master, text="Project Attributes", font = 'Arial 10 bold')
        self.title2.grid(column = 0, row = 19)
        
        self.e18 = tk.Label(self.master, text = "Pro Forma Occupancy (%)", font = ('Arial',8))
        self.e18.grid(column = 0, row = 20)
        self.entry18 = tk.Entry(self.master, textvariable = self.choiceVar18, width=20)
        self.entry18.grid(column = 1, row = 20)
        
        self.e19 = tk.Label(self.master, text = "Breakeven Occupancy (%)", font = ('Arial',8))
        self.e19.grid(column = 0, row = 21)
        self.entry19 = tk.Entry(self.master, textvariable = self.choiceVar19, width=20)
        self.entry19.grid(column = 1, row = 21)
        
        self.e20 = tk.Label(self.master, text = "New Construction?", font = ('Arial',8))
        self.e20.grid(column = 0, row = 22)
        self.entry20 = ttk.Combobox(self.master, textvariable = self.choiceVar20,values = ["Yes", "No"]) 
        self.entry20.grid(column = 1, row = 22)
        self.entry20.set(self.entry20.cget("values")[0])
        
        self.e21 = tk.Label(self.master, text = "Expected Completion Date:", font = ('Arial',8))
        self.e21.grid(column = 0, row = 23)
        self.entry21 = tk.Entry(self.master, textvariable = self.choiceVar21, width=20)
        self.entry21.grid(column = 1, row = 23)
        
        self.e22 = tk.Label(self.master, text = "School Start Date:", font = ('Arial',8))
        self.e22.grid(column = 0, row = 24)
        self.entry22 = tk.Entry(self.master, textvariable = self.choiceVar22, width=20)
        self.entry22.grid(column = 1, row = 24)
        
        self.e23 = tk.Label(self.master, text = "Existing Number of Beds:", font = ('Arial',8))
        self.e23.grid(column = 0, row = 25)
        self.entry23 = tk.Entry(self.master, textvariable = self.choiceVar23, width=20)
        self.entry23.grid(column = 1, row = 25)
        
        self.e24 = tk.Label(self.master, text = "Project Purpose", font = ('Arial',8))
        self.e24.grid(column = 0, row = 26)
        self.entry24 = ttk.Combobox(self.master, textvariable = self.choiceVar24,values = ["Adding New Beds", "Replace Beds", "Combination","Refunding"]) 
        self.entry24.grid(column = 1, row = 26)
    
        self.e25 = tk.Label(self.master, text = "Number of Net New Beds", font = ('Arial',8))
        self.e25.grid(column = 0, row = 27)
        self.entry25 = tk.Entry(self.master, textvariable = self.choiceVar25, width=20)
        self.entry25.grid(column = 1, row = 27)
        
        self.e26 = tk.Label(self.master, text = "Most Common Unit Type", font = ('Arial',8))
        self.e26.grid(column = 0, row = 28)
        self.entry26 = tk.Entry(self.master, textvariable = self.choiceVar26, width=20)
        self.entry26.grid(column = 1, row = 28)
        
        self.e27 = tk.Label(self.master, text = "Affiliation Agreement/Other Covenant ft.", font = ('Arial',8))
        self.e27.grid(column = 0, row = 29)
        self.entry27 = tk.Entry(self.master, textvariable = self.choiceVar27, width=20)
        self.entry27.grid(column = 1, row = 29)
        
        self.e28 = tk.Label(self.master, text = "Project Manager.", font = ('Arial',8))
        self.e28.grid(column = 0, row = 30)
        self.entry28 = tk.Entry(self.master, textvariable = self.choiceVar28, width=20)
        self.entry28.grid(column = 1, row = 30)
        
     #Market Study Attributes
        self.title3 = tk.Label(self.master, text="Market Study", font = 'Arial 10 bold')
        self.title3.grid(column = 0, row = 31)
        
        self.e29 = tk.Label(self.master, text = "Market Study?", font = ('Arial',8))
        self.e29.grid(column = 0, row = 32)
        self.entry29 = ttk.Combobox(self.master, textvariable = self.choiceVar29,values = ["Yes", "No"]) 
        self.entry29.grid(column = 1, row = 32)
        self.entry29.set(self.entry20.cget("values")[0])
        
        self.e30 = tk.Label(self.master, text = "Competitiveness", font = ('Arial',8))
        self.e30.grid(column = 0, row = 33)
        self.entry30 = ttk.Combobox(self.master, textvariable = self.choiceVar30,values = ["Competitive", "Overpriced", "Affordable"]) 
        self.entry30.grid(column = 1, row = 33)
        
        self.e31 = tk.Label(self.master, text = "Avg. Project Price/Sqft ($)", font = ('Arial',8))
        self.e31.grid(column = 0, row = 34)
        self.entry31 = tk.Entry(self.master, textvariable = self.choiceVar31, width=20)
        self.entry31.grid(column = 1, row = 34)
        
        self.e32 = tk.Label(self.master, text = "Avg. Comp Price/Sqft ($)", font = ('Arial',8))
        self.e32.grid(column = 0, row = 35)
        self.entry32 = tk.Entry(self.master, textvariable = self.choiceVar32, width=20)
        self.entry32.grid(column = 1, row = 35)
        
        
        self.e33 = tk.Label(self.master, text = "Assumed Capture Rate (%)", font = ('Arial',8))
        self.e33.grid(column = 0, row = 36)
        self.entry33 = tk.Entry(self.master, textvariable = self.choiceVar32, width=20)
        self.entry33.grid(column = 1, row = 36)
        
        
    #Conditional Combobox Variables
        self.choiceVar13.trace("w", self.on_trace_choice)
        self.refresh()
        self.choiceVar20.trace("w", self.on_trace_choice)
        self.refresh()
        self.choiceVar29.trace("w", self.on_trace_choice)
        self.refresh()
        self.data = []
        
    #Conditional combobox callback function
    def on_trace_choice(self,name, index, mode):
        self.refresh()
        
    #Conditional combobox refresh functions
    def refresh(self):
        choice13 = self.entry13.get()
        if choice13 == "None":
            self.entry14.configure(state="disabled")
        else:
            self.entry14.configure(state="normal")
            
        choice20= self.entry20.get()
        if choice20 == "Yes":
            self.entry21.configure(state="normal")     
            self.entry22.configure(state="normal") 
            self.entry23.configure(state="normal") 
            self.entry24.configure(state="normal") 
            self.entry25.configure(state="normal") 
            self.entry26.configure(state="normal") 
            self.entry27.configure(state="normal") 
            self.entry28.configure(state="normal") 
            
        else:
            self.entry21.configure(state="disabled")     
            self.entry22.configure(state="disabled") 
            self.entry23.configure(state="disabled") 
            self.entry24.configure(state="disabled") 
            self.entry25.configure(state="disabled") 
            self.entry26.configure(state="disabled") 
            self.entry27.configure(state="disabled") 
            self.entry28.configure(state="disabled")
        
        choice29= self.entry29.get()
        if choice29 == "Yes":
            self.entry30.configure(state="normal")     
            self.entry31.configure(state="normal") 
            self.entry32.configure(state="normal") 
            self.entry33.configure(state="normal") 
            
        else:
            self.entry30.configure(state="disabled")     
            self.entry31.configure(state="disabled") 
            self.entry32.configure(state="disabled") 
            self.entry33.configure(state="disabled") 
            
    #save & append function

    def saveinfo(self):
        v1 = self.entry1.get()
        v2 = self.entry2.get()
        v3 = self.entry3.get()
        v4 = self.entry4.get()
        v5 = self.entry5.get()
        v6 = self.entry6.get()
        v7 = self.entry7.get()
        v8 = self.entry8.get()
        v9 = self.entry9.get()
        v10 = self.entry10.get()
        v11= self.entry11.get()
        v12 = self.entry12.get()
        v13 = self.entry13.get()
        v14 = self.entry14.get()
        v15 = self.entry15.get()
        v16 = self.entry16.get()
        v17 = self.entry17.get()
        v18 = self.entry18.get()
        v19 = self.entry19.get()
        v20 = self.entry20.get()
        v21 = self.entry21.get()
        v22 = self.entry22.get()
        v23 = self.entry23.get()
        v24 = self.entry24.get()
        v25 = self.entry25.get()
        v26 = self.entry26.get()
        v27 = self.entry27.get()
        v28 = self.entry28.get()
        v29 = self.entry29.get()
        v30 = self.entry30.get()
        v31 = self.entry31.get()
        v32 = self.entry32.get()
        v33 = self.entry33.get()
        
        self.data.append([v1,v2,v3,v4,v5,v6,v7,v8,v9,v10,v11,v12,v13,v14,v15,v16,
                          v17,v18,v19,v20,v21,v22,v23,v24,v25,v26,v27,v28,v29,v30,v31,v32,v33])
      
        
   #Export Function
    def export(self):
        df = pd.DataFrame(self.data, columns = ["UniversityName","CurrentFallEnrollment","PriorFallEnrollment","EnrollmentChange",
                                               "AcceptanceRate","CampusType","ProgramType","DebtSchedule","MADS","StabilizedCoverageRatio",
                                               "AmountOfCap_I","CapI_EndDate","OperatingReserve_PctOf_Expense", "SurplusFundAmount",
                                               "ReleaseTestRequirement","GroundLeasePayment","GroundLeasePaymentStartDate","ProFormaOccupancy",
                                               "BreakevenOccupancy","NewConstruction","ExpectedCompletionDate","SchoolStartDate","ExistingBeds",
                                               "ProjectPurpose","NetNewBeds", "MostCommonUnitType","AffiliationAgreement_OtherCovenantFeats",
                                                "Project_Manager","MarketStudy","Competitiveness","AvgProjPriceSqft","AvgCompPriceSqFt",
                                                "AssumedCaptureRate"])
            
            
        pd.set_option('io.hdf.default_format','table')
        store = pd.HDFStore('store.h5')
        df.to_hdf('store.h5','table',append =True, data_columns=True,min_itemsize={'values':250})
        writer = pd.ExcelWriter(('DataBase.xlsx'), engine = 'xlsxwriter')
        store['table'].to_excel(writer, index=False, sheet_name = 'Sheet1', startcol=1, startrow=0,header=True)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        worksheet.set_zoom(90)
        
        format1 = workbook.add_format({'num_format': '#,##0.00'})
        format2 = workbook.add_format({'num_format': '#,##0.00'})
        format3 = workbook.add_format({'num_format': '#,##0.00'})
        format4 = workbook.add_format({'num_format': '0%'})
        format5 = workbook.add_format({'num_format': '0%'})
        
        worksheet.set_column('L:L',None, format1)
        worksheet.set_column('M:M',None, format2)
        worksheet.set_column('I:I',None, format3)
        worksheet.set_column('E:E',None, format4)
        worksheet.set_column('F:F',None, format5)
     
        
        worksheet.set_column('B:AZ',15)
        header_format = workbook.add_format({
            'bold':True,
            'text_wrap':True,
            'border':1 })
        
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0,col_num+1,value,header_format)
            
        pd.set_option('max_columns',None)
     
    #Tree Creation
    def inserttree(self):
        self.frame = tk.Frame(self.master)
        self.tree = ttk.Treeview(self.frame,columns = (['',"UniversityName","CurrentFallEnrollment","PriorFallEnrollment","EnrollmentChange",
                                               "AcceptanceRate","CampusType","ProgramType","DebtSchedule","MADS","StabilizedCoverageRatio",
                                               "AmountOfCap_I","CapI_EndDate","OperatingReserve_PctOf_Expense", "SurplusFundAmount",
                                               "ReleaseTestRequirement","GroundLeasePayment","GroundLeasePaymentStartDate","ProFormaOccupancy",
                                               "BreakevenOccupancy","NewConstruction","ExpectedCompletionDate","SchoolStartDate","ExistingBeds",
                                               "ProjectPurpose","NetNewBeds", "MostCommonUnitType","AffiliationAgreement_OtherCovenantFeats",
                                                "Project_Manager","MarketStudy","Competitiveness","AvgProjPriceSqft","AvgCompPriceSqFt",
                                                "AssumedCaptureRate"]))
        
        #set the headings
        
        self.tree.heading('0',text = 'blank', anchor='w')
        self.tree.column('0',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('1',text = 'University Name', anchor='w')
        self.tree.column('1',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('2',text = 'Current Fall Enrollment', anchor='w')
        self.tree.column('2',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('3',text = 'Prior Fall Enrollment', anchor='w')
        self.tree.column('3',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('4',text = 'Enrollment Change %', anchor='w')
        self.tree.column('4',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('5',text = 'Acceptance Rate %', anchor='w')
        self.tree.column('5',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('6',text = 'Campus Type', anchor='w')
        self.tree.column('6',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('7',text = 'Program Type', anchor='w')
        self.tree.column('7',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('8',text = 'Debt Schedule', anchor='w')
        self.tree.column('8',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('9',text = 'MADS ($)', anchor='w')
        self.tree.column('9',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('10',text = 'Stabilized Coverage Ratio', anchor='w')
        self.tree.column('10',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('11',text = 'Amount of Cap. I ($)', anchor='w')
        self.tree.column('11',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('12',text = 'Cap. I. End Date', anchor='w')
        self.tree.column('12',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('13',text = 'Operating Reserve % of Expenses (%)', anchor='w')
        self.tree.column('13',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('14',text = 'Surplus Fund Amount ($)', anchor='w')
        self.tree.column('14',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('15',text = 'Release Test Requirement', anchor='w')
        self.tree.column('15',width = 1, minwidth = 1, stretch=tk.YES)
        self.tree.heading('16',text = 'Ground Lease Requirement ($)', anchor='w')
        self.tree.column('16',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('17',text = 'Ground Lease Payment ($)', anchor='w')
        self.tree.column('17',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('18',text = 'Ground Lease Payment Start Date', anchor='w')
        self.tree.column('18',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('19',text = 'Pro Forma Occupancy (%)', anchor='w')
        self.tree.column('19',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('20',text = 'Breakeven Occupancy (%)', anchor='w')
        self.tree.column('20',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('21',text = 'New Construction', anchor='w')
        self.tree.column('21',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('22',text = 'Expected Completion Date', anchor='w')
        self.tree.column('22',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('23',text = 'School Start Date', anchor='w')
        self.tree.column('23',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('24',text = 'Existing Number of Beds', anchor='w')
        self.tree.column('24',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('25',text = 'Project Purpose', anchor='w')
        self.tree.column('25',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('26',text = 'Number of Net New Beds', anchor='w')
        self.tree.column('26',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('27',text = 'Most Common Unit Type', anchor='w')
        self.tree.column('27',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('28',text = 'Affiliation Agreement/Other Covenant Feat.', anchor='w')
        self.tree.column('28',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('29',text = 'Project Manager', anchor='w')
        self.tree.column('29',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('30',text = 'Market Study?', anchor='w')
        self.tree.column('30',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('31',text = 'Competitiveness', anchor='w')
        self.tree.column('31',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('32',text = 'Avg. Project Price/Sqft', anchor='w')
        self.tree.column('32',width = 20, minwidth = 200, stretch=tk.YES)
        self.tree.heading('33',text = 'Assumed Capture Rate (%)', anchor='w')
        self.tree.column('33',width = 20, minwidth = 200, stretch=tk.YES)
        
      #scroll creation
    
        ysb = ttk.Scrollbar(self.frame, orient = 'vertical', command = self.tree.yview)
        xsb = ttk.Scrollbar(self.frame, orient = 'horizontal', command = self.tree.xview)
        self.tree.configure(yscroll=ysb.set,xscroll=xsb.set)
        
        self.tree.grid(row=0, column = 0, rowspan = 1, columnspan =4, sticky = 'nsew')
        ysb.grid(row=0, column=6, columnspan=4,pady=20,sticky='nsew')
        xsb.grid(row=1, column=6, columnspan=4,pady=20,sticky='nsew')
        self.frame.grid(row=40, column = 0,padx=20, rowspan = 1, columnspan =4, sticky = 'nsew')
        self.treeview = self.tree
        
    #initialization of tree insert loop
        self.id = 0
        self.iid = 0
        
    #Tree insertion
    def insert_data(self):
        self.treeview.insert('','end', iid=self.iid, text='item_'+str(self.id), 
                            values=(self.entry1.get(), self.entry2.get(), self.entry3.get(),
                                    self.entry4.get(), self.entry5.get(),self.entry6.get(),
                                  self.entry7.get(),self.entry8.get(),self.entry9.get(),self.entry10.get(),
                                    self.entry11.get(),self.entry12.get(),self.entry13.get(),self.entry14.get(),
                                   self.entry15.get(),self.entry16.get(),self.entry17.get(),self.entry18.get(),
                                   self.entry19.get(),self.entry20.get(),self.entry21.get(), self.entry22.get(),
                                   self.entry23.get(),self.entry24.get(),self.entry25.get(),self.entry25.get(),
                                   self.entry26.get(),self.entry27.get(),self.entry28.get(),self.entry29.get(),
                                   self.entry30.get(),self.entry31.get(),self.entry32.get(),self.entry33.get()))
        
    #completion of tree insert loop
        self.iid = self.iid + 1
        self.id = self.id +1
        
    #Tree row deletion
    
    def delete_data(self):
        row_id=int(self.tree.focus())
        self.treeview.delete(row_id)
        
    #Close app
def main():
    root= tk.Tk()
    app=MainApp(root)
    root.mainloop()

if __name__== '__main__':
    main()
        


# In[2]:


nbconvert --to script [Student_Housing_UI].ipynb


# In[ ]:




