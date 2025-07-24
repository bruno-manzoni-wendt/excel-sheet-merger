# %%
import pandas as pd

file_path = r'your_file_path_here.xlsx'  # replace with your actual file path
df_all = pd.read_excel(file_path, sheet_name=None, header=None)

# %%
sheet_names = list(df_all.keys())
for nome in ['Dashboard', 'Página36', 'Página35', 'MP A']:
    sheet_names.remove(nome)
list_df = []

for i, sheet in enumerate(sheet_names, 1):
    print(i, sheet)
    df = df_all[sheet]
    col = df[0].loc[6:].reset_index(drop=True) # column names from row 6 of column 0
    data = df.loc[6:, 2:].reset_index(drop=True) # data from columns 1 starting from row 6

    unpivot_df = pd.DataFrame(data).T  # transpose to align properly
    unpivot_df.columns = col  #set columns names
    for char in ['', ' ', '-']:
        unpivot_df = unpivot_df.replace(char, pd.NA).infer_objects(copy=False) #replace empty strings, spaces and dashes with pd.NA

    unpivot_df = unpivot_df.dropna(subset=unpivot_df.columns[2:], how='all') #drop empty rows not checking the first two columns
    unpivot_df['Matéria Prima'] = df.loc[4, 2]
    list_df.append(unpivot_df) #append the treated dataframe to the list

# %%
consolidated_df = pd.concat(list_df, ignore_index=True) #concatenate all dataframes into one
consolidated_df = consolidated_df.drop(columns=['Produtos em que a MP é REPROVADA', 'p&d '])
consolidated_df['País de Origem'] = consolidated_df['País de Origem'].str.replace('Brasil ', 'Brasil')

coluns_list = list(consolidated_df.columns)
coluns_list = [coluns_list[-1]] + coluns_list[:-1]
consolidated_df = consolidated_df[coluns_list] # reorder columns

consolidated_df.to_excel(r'Consolidated_data.xlsx', sheet_name='Data', index=False)
print("Finished")
