# -*- coding: utf-8 -*-
"""
Created on Tue Jul 31 04:41:07 2018

@author: Rubiks
"""

import pandas as pd
import numpy as np

#file name to be uploaded
filename = 'Book3.xlsx'

#client name, output file will be named with their name
outfile = 'Ronel Wallace'

# changes X's to 1s and NaN values to 0s
def create_dataframe(filename):
    df = pd.read_excel(filename)
    df.reset_index(inplace=True)
    last_row = pd.Series([0,0,0,0,0,0,0,0,0,0,0,0,0],index=df.columns)
    df = df.append(last_row,ignore_index=True)
    df.replace(np.nan, 0, inplace=True)
    df.replace('x', 'X', inplace=True)
    df.columns = ['index', 'Item', 'Strongly Disagree', 'Disagree', 'Not sure', 'Agree',
       'Strongly Agree', 'Strongly Disagree.1', 'Disagree.1', 'Not sure.1',
       'Agree.1', 'Strongly Agree.1', 'Denotion']
    
    #create a column of skill talley to be updated later
    df['Skill Talley'] = 0
    return df


# standards total
def skill_std(data, standard_df, skill_type):
    skill_total = data[final_data['Denotion'] == skill_type]


    SD = len(skill_total['Strongly Disagree'].drop_duplicates())
    
    D = len(skill_total['Disagree'].drop_duplicates())
    NT = len(skill_total['Not sure'].drop_duplicates())
    A = len(skill_total['Agree'].drop_duplicates())
    SA = len(skill_total['Strongly Agree'].drop_duplicates())
    
    if SD > 1:
        standard_df['Strongly Disagree'][skill_type] = sum(skill_total['Strongly Disagree.1'][skill_total['Disagree'] == 'X'])
    else:
        standard_df['Strongly Disagree'][skill_type] = 0
        
    
    if D > 1:
        standard_df['Disagree'][skill_type] = sum(skill_total['Disagree.1'][skill_total['Disagree'] == 'X'])
    else:
        standard_df['Disagree'][skill_type] = 0
        
    if NT > 1:
        standard_df['Not sure'][skill_type] = sum(skill_total['Not sure.1'][skill_total['Not sure'] == 'X'])
    else:
        standard_df['Not sure'][skill_type] = 0
    
    if A > 1:
        standard_df['Agree'][skill_type]=sum(skill_total['Agree.1'][skill_total['Agree'] == 'X'])
    else:
        standard_df['Agree'][skill_type] = 0
        
    if SA > 1:
        standard_df['Strongly Agree'][skill_type] = sum(skill_total['Strongly Agree.1'][skill_total['Strongly Agree'] == 'X'])
    else:
        standard_df['Strongly Agree'][skill_type] = 0
        
    standard_df['Std'][skill_type] = skill_total.shape[0] * 100
        
    return standard_df



# create proficeicy column
def create_proficiency_column(data):
    sections = []
    index = []
    count = 0
    for j in data['Item'].str.contains('SECTION'):
        if j :
            sections.append(1)
            count = 0
            index.append(count)
        else:
            count += 1
            index.append(count)
            sections.append(0)
    
    data['index'] = index
    data['Prof.level'] = sections
    
    return data

# grading with the standards provided
def grading(data):
    for i in range(len(data)):
        if data['Strongly Disagree'].iloc[i] == 'X':
            data['Skill Talley'].iloc[i] = data['Strongly Disagree.1'].iloc[i]
        elif data['Disagree'].iloc[i] == 'X':
            data['Skill Talley'].iloc[i] = data['Disagree.1'].iloc[i]
        elif data['Not sure'].iloc[i] == 'X':
            data['Skill Talley'].iloc[i] = data['Not sure.1'].iloc[i]
        elif data['Agree'].iloc[i] == 'X':
            data['Skill Talley'].iloc[i] = data['Agree.1'].iloc[i]
        elif data['Strongly Agree'].iloc[i] == 'X':
            data['Skill Talley'].iloc[i] = data['Strongly Agree.1'].iloc[i]
    return data

def calculate_skills_total(data):
    section_score = []
    section_tot = []
    proficient_level = []
    count_score, count_tot = 0, 0
    
    #--------------------------------------------------------
    for i in range(len(data)):
        if data['Prof.level'].iloc[i] == 0:
            count_score += data['Skill Talley'].iloc[i]
            count_tot += 1
        else:
            section_score.append(count_score)
            section_tot.append(count_tot*100)
            count_score = 0
            count_tot = 0
    
    #--------------------------------------------------------
    for p in range(len(section_score)):
        if section_tot[p] > 0:
            proficient_level.append(round((section_score[p] / section_tot[p])*100,2))
        else:
            proficient_level.append(0)
    
    
    del proficient_level[0]
    
    #--------------------------------------------------------
    
    for i in range(len(data)):
        if data['Prof.level'].iloc[i] == 1:
            
            talley = section_score[0]
            skill_tot = section_tot[0]
            
            data['Denotion'].iloc[i] = skill_tot
            data['Skill Talley'].iloc[i] = talley
    
            section_score.remove(talley)
            section_tot.remove(skill_tot)
    
    
    #--------------------------------------------------------
    proficient_level.append(0.0)
    
    
    prof = 0
    for i in range(len(data)):
        if data['Prof.level'].iloc[i] == 1:
            prof_level = proficient_level[prof]
            data['Prof.level'].iloc[i] = prof_level
            prof += 1
    
    return data



df = create_proficiency_column(create_dataframe(filename))

graded_df = grading(df)

final_data = calculate_skills_total(graded_df)

final_data.set_index("index",inplace=True)

final_data.to_excel(outfile + str('.xlsx'),outfile)

std_index = np.arange(len(final_data), len(final_data) + 3)

std = pd.DataFrame(columns=['Strongly Disagree','Disagree','Not sure',
                            'Agree','Strongly Agree','Std'], 
                            index=['T','S','K'])


std_new = skill_std(final_data, std, "S")
std_new = skill_std(final_data, std, "T")
std_new = skill_std(final_data, std, "K")









#############################################################################

final_data.index.name = ""

writer = pd.ExcelWriter(outfile + str(' new.xlsx'), engine='xlsxwriter')

final_data.to_excel(writer, sheet_name=outfile)



worksheet = writer.sheets[outfile]


workbook = writer.book



header_format_two = workbook.add_format(
    {
        "bg_color":"#92d050",
        "font":"Calibri",
        "font_size":10,
        "bold":True,
        "border":1,
        "text_wrap":True,
        'valign': 'top'
    }
)


header_format_one = workbook.add_format(
    {
        "bg_color":"#ffffff",
        "font":"Calibri",
        "font_size":10,
        "bold":True,
        "border":1,
        "text_wrap":True,
        'valign': 'top'
    }
)

item_format = workbook.add_format(
    {
        "font":"Calibri",
        "font_size":10,
        "border":1,
        "text_wrap":True,
    }
)


item_100_format = workbook.add_format(
    {
        "bg_color":"#92d050",
        "valign": "center",
        "font":"Calibri",
        "font_size":10,
        "border":1,
        "text_wrap":True,
    }
)



section_format = workbook.add_format(
    {
        "bg_color":"#ffc000",
        "bold":True,
        "font":"Calibri",
        "font_size":12
    }
)


red_alert = workbook.add_format(
    {
        "bg_color":"#e8715c",
        "bold":True,
        "font":"Calibri",
        "font_size":12
    }
)

prof_color = workbook.add_format(
    {
        "bg_color":"#449eed",
        "font":"Calibri",
        "font_size":10,
        "bold":True,
        "border":1,
        "text_wrap":True,
        'valign': 'top'
    }
)

standard_format = workbook.add_format(
    {
        "bg_color":"#909296",
        "font":"Calibri",
        "font_size":10,
        "border":1,
        "text_wrap":True,
        'valign': 'right'
    }
)


worksheet.set_column('B:B', 68)

worksheet.set_column('C:O', 8.43)


for col_num, value in enumerate(final_data.columns.values):
    for row_num, row_value in enumerate(final_data.values):
        worksheet.write(row_num + 1, col_num + 1, row_value[col_num], item_format)
        
        
for col_num, value in enumerate(final_data.columns.values[:-2]):
    for row_num, row_value in enumerate(final_data.values):
        if row_value[col_num] == 100:
            worksheet.write(row_num + 1, col_num + 1, row_value[col_num], item_100_format)
            
for row_num, value in enumerate(final_data.index.values):
    if value == 0 :
        worksheet.write(row_num + 1, 1, final_data['Item'].iloc[row_num], section_format)
        for col_num2, value2 in enumerate(final_data.columns.values[1:-1]):
            worksheet.write(row_num + 1, col_num2 + 2 , final_data[value2].iloc[row_num], section_format)
            if float(final_data["Prof.level"].iloc[row_num]) < 50:
                worksheet.write(row_num + 1, 14, final_data['Prof.level'].iloc[row_num], red_alert)
        


# header style
for col_num, value in enumerate(final_data.columns.values):
    if col_num <= 5:
        worksheet.write(0, col_num + 1, value, header_format_one)
    elif col_num < 13:
            worksheet.write(0, col_num + 1, value, header_format_two)
    else:
        worksheet.write(0, col_num + 1, value, prof_color)
        
#cleaning up the excel file




        

worksheet.write(len(final_data), 14 , "", item_format)
worksheet.write(len(final_data), 1 , "", section_format)


#clean by removing zeroes from rows with white heading backgrounds
for i in range(len(final_data)):
    if final_data['Strongly Disagree'].iloc[i] == 0.0:
        worksheet.write(i +1, 2, '', item_format)
        
    if final_data['Disagree'].iloc[i] == 0.0:
        worksheet.write(i +1, 3, '', item_format)
        
    if final_data['Not sure'].iloc[i] == 0.0:
        worksheet.write(i +1, 4, '', item_format)
        
    if final_data['Agree'].iloc[i] == 0.0:
        worksheet.write(i +1, 5, '', item_format)
        
    if final_data['Strongly Agree'].iloc[i] == 0.0:
        worksheet.write(i +1, 6, '', item_format)
        
#remove zeroes from sections' rows (the yellow ones)
for row_num, value in enumerate(final_data.index.values):

    if value == 0 :
        worksheet.write(row_num + 1, 0 , "", item_format)
        for col_num2, value2 in enumerate(final_data.columns.values[1:-3]):
            worksheet.write(row_num + 1, col_num2 + 2 , "", section_format)



#standards printing


row_start = len(final_data) + 3
worksheet.write("C" + str(row_start), "SD", standard_format)
worksheet.write("D" + str(row_start), "D", standard_format)
worksheet.write("E" + str(row_start), "NS", standard_format)
worksheet.write("F" + str(row_start), "A", standard_format)
worksheet.write("G" + str(row_start), "SA", standard_format)
worksheet.write("H" + str(row_start), "Std", standard_format)


for col_num, value in enumerate(std_new.columns.values):
    worksheet.write(row_start , col_num + 2, value, standard_format)
    for row_num, row_value in enumerate(std_new.index.values):
        worksheet.write(row_start + row_num, 1, row_value, standard_format)
        worksheet.write(row_start + row_num, col_num + 2, std_new[value].iloc[row_num], standard_format)


workbook.close()