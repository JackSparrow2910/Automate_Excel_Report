import pandas as pd
import os

#This function creates pivot table in report.xlsx entering path of file
def create_pivot_table(month,application_path):
    path=input("Enter the path of the excel file: ")
    final_path=os.path.join(application_path,path)
    df = pd.read_excel(final_path)
    df = df[['Gender','City','Total']]
    pivot_table = df.pivot_table(index='Gender',columns='City',values='Total',aggfunc='sum').round(0)
    file = f'report_{month}.xlsx'
    pivot_table.to_excel(file,startrow=4)
    return file

