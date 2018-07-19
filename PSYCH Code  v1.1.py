
# coding: utf-8

# In[3]:


import pandas as pd 
import numpy as np 
from functools import reduce
import os 
import openpyxl


# In[6]:


list_control = [27,29,31] 
list_shams = [5,9,11,19,21,23]
list_stimulated = [13,15,17]
rat_category = {'controls': list_control,'shams': list_shams,'stimulated': list_stimulated}


# In[14]:


print(rat_category )


# In[8]:


'''
df_1 = dataframe with only the value of total intake, meal size etc 
df_2 = dataframe with value of total intake, meal size etc + column for rat type too 

'''

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
    
    


# In[17]:


averages = to_be_printed.groupby(['type_of_rat']).mean()


# In[18]:


averages


# In[ ]:


def transfer_data(filenames, result_page_list , rat_category, output_filename):
    writer = pd.ExcelWriter(output_filename + '.xlsx', engine='xlsxwriter')
    
    for sheetname in result_page_list: 
        if not output_filename.endswith('.xlsx'):
            output_filename = output_filename + '.xlsx'
            
            
        final_df = tabulate(filenames, sheetname, rat_category)
        
        final_df.to_excel(writer, sheet_name=sheetname)
    
    writer.save()
    print("Result saved in file : " + output_filename)
        
  


# In[ ]:


#final function 
filenames = [str(x) for x in input("Enter the name of files with space in between : ").split()]
list_control = [27,29,31] 
list_shams = [5,9,11,19,21,23]
list_stimulated = [13,15,17]
rat_category = {'controls': list_control,'shams': list_shams,'stimulated': list_stimulated}

output_filename = input("Enter the ouput filename : ")

transfer_data(filenames = filenames, result_page_list=['total intake','meal number','meal size','intermeal interval'],
              rat_category = rat_category, output_filename = output_filename)

