import glob
import pandas as pd
import numpy as np

path='path/to/sheets/to/combine' # use your path
allFiles=glob.glob(path + "/*.xlsx")
frame=pd.DataFrame()
list_=[]
for file_ in allFiles:
    df=pd.read_excel(file_,sheet_name='June Queries',index_col=None, header=0)
    column_indices=[0,1,2,3,4,5]
    new_names=['a','b','c','d','e','f']
    old_names=df.columns[column_indices]
    df.rename(columns=dict(zip(old_names, new_names)), inplace=True)
    ur_row=df.loc[df['a'].str.contains('Final',na=False)].index.tolist()
    df=df.iloc[:ur_row[0]]
    list_.append(df)
frame=pd.concat(list_)
df_comb=frame

df_comb=df_comb[['a','b','c','d','e','f']]
df_comb=df_comb[~df_comb['a'].astype(str).str.contains('agent')]
df_comb.dropna(subset=['e'],inplace=True)

df_comb['a']=df_comb['a'].ffill()
df_comb['b']=df_comb['b'].ffill()
df_comb['c']=df_comb['c'].ffill()

df_comb=df_comb.rename(columns={
          'a': 'Protocols', 
          'b': 'OARS Page', 
          'c': 'Site', 
          'd': 'Subject ID',
          'e': 'Query',
          'f': 'Query Date'})

writer=pd.ExcelWriter('combined_queries.xlsx', engine='xlsxwriter', date_format='mm/dd/yyyy', datetime_format='mm/dd/yyyy')

df_comb.to_excel(writer, sheet_name='combined')
workbook =writer.book
worksheet1=writer.sheets['combined']
writer.save()
