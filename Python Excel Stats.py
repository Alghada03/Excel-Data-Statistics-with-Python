import pandas as pd
import numpy as np
from tabulate import tabulate

excel_file = 'C:/Users/ghada/Downloads/Data Set.xlsx'   
sheet_names = ['compact cars', 'midsize cars', 'large cars']
column_names = ['mpg', 'weight']  

stats_datasets = {}

for sheet_name in sheet_names:
    stats_datasets[sheet_name] = []

for sheet_name in sheet_names:
    data = pd.read_excel(excel_file, sheet_name=sheet_name)

    data_values1 = data[column_names[0]].values
    data_values2 = data[column_names[1]].values

    sample_size = len(data_values1)

    mean1 = np.mean(np.nan_to_num(data_values1, nan=0))
    median1 = np.median(np.nan_to_num(data_values1, nan=0))
    std_dev1 = np.std(np.nan_to_num(data_values1, nan=0))
    variance1 = np.var(np.nan_to_num(data_values1, nan=0))
    maximum1 = np.max(np.nan_to_num(data_values1, nan=0))
    minimum1 = np.min(np.nan_to_num(data_values1, nan=0))
    Q1_1 = np.percentile(np.nan_to_num(data_values1, nan=0), 25)
    Q3_1 = np.percentile(np.nan_to_num(data_values1, nan=0), 75)
    IQR1 = Q3_1 - Q1_1

    mean2 = np.mean(np.nan_to_num(data_values2, nan=0))
    median2 = np.median(np.nan_to_num(data_values2, nan=0))
    std_dev2 = np.std(np.nan_to_num(data_values2, nan=0))
    variance2 = np.var(np.nan_to_num(data_values2, nan=0))
    maximum2 = np.max(np.nan_to_num(data_values2, nan=0))
    minimum2 = np.min(np.nan_to_num(data_values2, nan=0))
    Q1_2 = np.percentile(np.nan_to_num(data_values2, nan=0), 25)
    Q3_2 = np.percentile(np.nan_to_num(data_values2, nan=0), 75)
    IQR2 = Q3_2 - Q1_2

    stats_datasets[sheet_name].append(['mpg', sample_size, mean1, median1, std_dev1, variance1, maximum1, minimum1, IQR1])
    
    stats_datasets[sheet_name].append(['weight', sample_size, mean2, median2, std_dev2, variance2, maximum2, minimum2, IQR2])
 
for i, sheet_name in enumerate(sheet_names):
    data = pd.read_excel(excel_file, sheet_name=sheet_name)
    print(f"Data for sheet '{sheet_name}':\n{data.head()}\n")
    
for sheet_name, stats in stats_datasets.items():
    print(f"Statistics for {sheet_name}:")
    headers = ['Column', 'Sample Size', 'Mean', 'Median', 'Std Dev', 'Variance', 'Maximum', 'Minimum', 'IQR']
    print(tabulate(stats, headers=headers, tablefmt="grid"))
    print()
