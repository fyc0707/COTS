import pandas as pd 

df = pd.read_excel('log/2020-03-15/rcvList.xlsx', header=7)
df = df[df['Event Type'].str.contains('RCT|RCV')]
df = df.drop(['2nd UD field Reception', 'Event Type'], axis=1)
df.columns = ['CQC#','Type','CQE','Customer','Part Name','Trace Code','Instruction','B2B']
df['B2B'] = df['B2B'].apply(lambda x: False if pd.isna(x) else True)
df['Instruction'] = df['Instruction'].apply(lambda x: str(x)[19:])
print(df)
