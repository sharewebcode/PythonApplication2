import pandas as pd
import openpyxl
from openpyxl.chart import Reference, Series, BarChart
from openpyxl.styles import Font, Alignment
from openpyxl.styles import Border, Side, Color, PatternFill

#data = pd.read_excel('pandasSample.xlsx', sheet_name=0)
#df = pd.DataFrame(data)
 
#data = {'Name':['Tom', 'Jack', 'Steve', 'Ricky'],'Age':[28,34,29,42]}
#df = pd.DataFrame(data, index=['rank1','rank2','rank3','rank4'])
#print (data)


# 파일 읽어서 병합하기
df_Sam = pd.read_excel('2017Sam.xlsx')
df_Sam.set_index('date', inplace=True)

df_LG = pd.read_excel('2017LG.xlsx')
df_LG.set_index('date', inplace=True)

df_merge = pd.DataFrame()
print (df_merge)

print ('=======================================')

df_merge['삼성'] = df_Sam['total']
df_merge['LG'] = df_LG['total']

print (df_merge)