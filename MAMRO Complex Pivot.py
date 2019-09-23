

import pandas as pd
import matplotlib as plt
import numpy as np
from pandas import ExcelWriter
import openpyxl 
from openpyxl import Workbook
from openpyxl import load_workbook

def pivot_yearly_average(table,value):
    table = pd.pivot_table(table, values=[value], index=['Family','MDS','MAMRO'],columns='Year', aggfunc=np.sum)
    table[value,"Yearly Average"] = table.iloc[:,:3].mean(axis=1)
    return table

def cost_per_count(costs,counts,average_tech_salary):
    costs['A-Cost Per']=0
    i=0
    while i<costs['MAMRO'].count():
        if costs.loc[i,'MAMRO']=="Techs":
            if counts.loc[i,'TAI']==0:
                costs.loc[i,'A-Cost Per']=0
            else:
                costs.loc[i,'A-Cost Per']=costs.loc[i,'A-Cost']/counts.loc[i,'TAI']/average_tech_salary
            
        else:
            if counts.loc[i,'Hours']==0:
                costs.loc[i,'A-Cost Per']=0
            else:
                costs.loc[i,'A-Cost Per']=costs.loc[i,'A-Cost']/counts.loc[i,'Hours']
        i=i+1
    return costs


def aggregate_costs (raw):
    costs = raw.groupby(['Family','MDS','Year','MAMRO'],as_index=True)['Cost','A-Cost'].sum()
    costs = costs.drop(index=['0','Hours','TAI','TOC'],level='MAMRO')
    costs.reset_index(inplace=True)
    return costs

def aggregate_counts (raw):
    counts = raw.groupby(['Family','MDS','Year','MAMRO','T/M/S'],as_index=True)['Hours','TAI'].mean()
    counts = counts.groupby(['Family','MDS','Year','MAMRO']).sum()
    counts = counts.drop(index=['0','Hours','TAI','TOC'],level='MAMRO')
    counts.reset_index(inplace=True)
    return counts

def import_rawfile(rawname):
    rawfile = pd.read_csv(rawname)
    rawfile['Year']=rawfile[['Year']].astype(str)
    return rawfile

def export_to_file (table,output_file,sheetname):
    book = load_workbook(output_file)
    book[sheetname].delete_rows(1,5000)
    
    writer = pd.ExcelWriter(output_file, engine='openpyxl') 
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    table.to_excel(writer, sheetname,index=False)
    
    writer.save()

def mamro_pivot(input_names,output_file,average_tech_salary):
    pd.options.display.float_format = '{:,.1f}'.format    
    writer = pd.ExcelWriter(output_file)
    i=0
    
    while i<2:
        rawfile=import_rawfile(input_files[i])
        
        agg_costs = aggregate_costs(rawfile)
        agg_counts = aggregate_counts(rawfile)
        agg_costs = cost_per_count(agg_costs,agg_counts,average_tech_salary)
            
        export_costs = pd.concat([pivot_yearly_average(agg_costs,"A-Cost"),pivot_yearly_average(agg_costs,"A-Cost Per")], axis=1)
        export_costs = export_costs.stack()    

        export_to_file(export_costs.reset_index(),output_file,input_files[i].replace('.csv',''))      
    
        i=i+1

input_files=['aftoc.csv','vamosc.csv']
output_file='MAMRO 2020 Python.xlsx'
average_tech_salary=91284
mamro_pivot(input_files,output_file,average_tech_salary)




