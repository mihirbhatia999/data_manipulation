import pandas as pd 
import numpy as np 
from functools import reduce
import os 
import openpyxl

'''
df_1 = dataframe with only the value of total intake, meal size etc 
df_2 = dataframe with value of total intake, meal size etc + column for rat type too 

'''
#function that converts the data into tables------------------------------------------------------------
def tabulate(filenames, result_page , rat_category):
    day_wise_result = []
    final_df = pd.DataFrame()
    #for each day in filenames 
    for day in filenames: 
        if not day.endswith('.xlsx'):
            day = day + '.xlsx'
        
        df = pd.read_excel(day,sheet_name = 'PSC Totals', skiprows = 8)
        
        if result_page =='total intake':
            df_1 = df.iloc[:,7]
            
        elif result_page == 'meal number':
            df_1 = df.iloc[:,8]
            
        elif result_page == 'meal size':
            df_1 = df.iloc[:,12]
            
        elif result_page == 'intermeal interval':
            df_1 = df.iloc[:,11]
        else: 
            print("invalid result page")
            
        
    #making a list of type of rats and convert to dataframe 
        type_of_rat = [] 
        for index, row in df.iterrows():
            if row[0] in rat_category['controls']:
                type_of_rat.append('controls')
                
            elif row[0] in rat_category['shams']:
                type_of_rat.append('shams')
                
            elif row[0] in rat_category['stimulated']:
                type_of_rat.append('stimulated')
        
        type_of_rat = pd.DataFrame(np.array(type_of_rat))
        type_of_rat.columns = ['type_of_rat']
        
        df_2 = pd.concat([df_1, type_of_rat ], ignore_index=False, axis = 1)
        #now group df2 by type 
        df_2 = df_2.sort_values(by = "type_of_rat")
        #print(df_2)
        
        #now take only the
        #print(" ")
        #print("GROUP BY")
        #print(df_2)
        day_wise_result.append(df_2)
    
    
    #merging all the days dataframes into a single one    
    #result = reduce(lambda x, y: pd.merge(x, y, on = 'type_of_rat'), day_wise_result)
    for dfs in day_wise_result: 
        final_df = pd.concat([final_df,dfs.iloc[:,0]], axis = 1 )
    final_df.columns = filenames   
    final_df = pd.concat([final_df,type_of_rat.sort_values(by = 'type_of_rat')], axis = 1 )
    return final_df
    
    



#function that transfers data ----------------------------------------------------------------------------------
def transfer_data(filenames, result_page_list , rat_category, output_filename):
    writer = pd.ExcelWriter(output_filename + '.xlsx', engine='xlsxwriter')
    
    for sheetname in result_page_list: 
        if not output_filename.endswith('.xlsx'):
            output_filename = output_filename + '.xlsx'
            
        #using the tabulate function to obtain the final dataframe from the requested files    
        final_df = tabulate(filenames, sheetname, rat_category)
        
        #making dataframe containing category wise matrices 
        averages = final_df.groupby(['type_of_rat']).mean()
        
        #adding a column containing row-wise means 
        final_df['Mean'] = final_df.mean(axis=1)
        
        #rearranging the columns of final dataframe
        new_col_names = []
        for col in final_df.columns:
            if col != 'type_of_rat':
                new_col_names.append(col)
        new_col_names.append('type_of_rat')
        
        final_df = final_df[new_col_names]
        print(sheetname + " transferred")
        
        
        #writing the dataframes to the excel sheet 
        final_df.to_excel(writer, sheet_name=sheetname, index=False)
        averages.to_excel(writer, sheet_name=sheetname, index=False, startrow=(final_df.shape[0]) + 2 , startcol= 0) 
        
    #calcuating differences for each excel sheet 
    avg_shams_controls = averages.loc['shams'] - averages.loc['controls'] 
    avg_stimulated_controls = averages.loc['stimulated'] - averages.loc['controls']
    sum_avg_shams_controls = sum(avg_shams_controls)
    sum_avg_stimulated_controls = sum(avg_stimulated_controls)
    writer.save()
    
    print("Result saved in file : " + output_filename)
        

#final execution of the functions ----------------------------------------------------------------------------------------
filenames = [str(x) for x in input("Enter the name of files with space in between : ").split()]

#list_control = [27,29,31] 
#list_shams = [5,9,11,19,21,23]
#list_stimulated = [13,15,17]

list_control = [int(x) for x in input("Enter the CONTROL RAT NO.S : ").split()]
list_shams = [int(x) for x in input("Enter the SHAM RAT NO.S : ").split()]
list_stimulated = [int(x) for x in input("Enter the STIMULATED RAT NO.S : ").split()]


rat_category = {'controls': list_control,'shams': list_shams,'stimulated': list_stimulated}

output_filename = input("Enter the ouput filename : ")

final_df , averages , avg_shams_controls, avg_stimulated_controls, sum_avg_shams_controls, sum_avg_stimulated_controls = transfer_data(filenames = filenames, result_page_list=['total intake','meal number','meal size','intermeal interval'],
              rat_category = rat_category, output_filename = output_filename)


# now we manually calculate some of the sums from each sheet of the resultant workbook and paste the results there-------

