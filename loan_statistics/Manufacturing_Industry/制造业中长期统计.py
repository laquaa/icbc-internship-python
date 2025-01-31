import pandas as pd

#Read
loan_2023_12 = pd.read_excel('202312反传数据.xlsx', sheet_name='贷款 ')
loan_2024_06_xls = pd.read_excel('202406返传数据.xls', sheet_name='贷款')
loan_2024_06_xls.to_excel('202406返传数据.xlsx', engine='openpyxl', index=False)
loan_2024_06 = pd.read_excel('202406返传数据.xlsx', sheet_name='Sheet1')

#Select
selected_columns_2023_12 = loan_2023_12[['客户全称', '客户经理', '支行考核口径部门','贷款余额','贷款期限','国家代码','融资投向']]
selected_columns_2024_06 = loan_2024_06[['客户全称', '客户经理', '支行考核口径部门','贷款余额','贷款期限','国家代码','融资投向']]

filtered_2023_12 = selected_columns_2023_12[(selected_columns_2023_12['国家代码'] == 156) & (selected_columns_2023_12['融资投向'].str.startswith('C'))]
filtered_2024_06 = selected_columns_2024_06[(selected_columns_2024_06['国家代码'] == 156) & (selected_columns_2024_06['融资投向'].str.startswith('C'))]

#divide
loan_2023_12_long_medium_0 = filtered_2023_12[filtered_2023_12['贷款期限'] != 1]
loan_2024_06_long_medium_0 = filtered_2024_06[filtered_2024_06['贷款期限'] != 1]

loan_2023_12_long_medium_1 = loan_2023_12_long_medium_0.drop(columns=['贷款期限','国家代码','融资投向'])
loan_2024_06_long_medium_1 = loan_2024_06_long_medium_0.drop(columns=['贷款期限','国家代码','融资投向'])

#customer
customer_2023_12_long_medium = loan_2023_12_long_medium_1.groupby(['客户全称'], as_index=False).agg({'支行考核口径部门':'first',
                                                                        '客户经理':'first',
                                                                        '贷款余额': 'sum'})
customer_2024_06_long_medium = loan_2024_06_long_medium_1.groupby(['客户全称'], as_index=False).agg({'支行考核口径部门':'first',
                                                                        '客户经理':'first',
                                                                        '贷款余额': 'sum'})
#department
department_2023_12_long_medium = loan_2023_12_long_medium_1.groupby(['支行考核口径部门'], as_index=False).agg({'贷款余额': 'sum'})
department_2024_06_long_medium = loan_2024_06_long_medium_1.groupby(['支行考核口径部门'], as_index=False).agg({'贷款余额': 'sum'})

#merge
customer_2023_12_long_medium = customer_2023_12_long_medium.rename(columns={'贷款余额': '贷款余额_2023_12'})
customer_2024_06_long_medium = customer_2024_06_long_medium.rename(columns={'贷款余额': '贷款余额_2024_06'})

merged_customer_long_medium = pd.merge(customer_2023_12_long_medium, customer_2024_06_long_medium, on=['客户全称','支行考核口径部门','客户经理'], how='outer')

department_2023_12_long_medium = department_2023_12_long_medium.rename(columns={'贷款余额': '贷款余额_2023_12'})
department_2024_06_long_medium = department_2024_06_long_medium.rename(columns={'贷款余额': '贷款余额_2024_06'})

merged_department_long_medium = pd.merge(department_2023_12_long_medium, department_2024_06_long_medium, on='支行考核口径部门', how='outer')

merged_customer_long_medium = merged_customer_long_medium.fillna(0)
merged_department_long_medium = merged_department_long_medium.fillna(0)

merged_customer_long_medium['增量'] = merged_customer_long_medium['贷款余额_2024_06'] - merged_customer_long_medium['贷款余额_2023_12']
merged_department_long_medium['增量'] = merged_department_long_medium['贷款余额_2024_06'] - merged_department_long_medium['贷款余额_2023_12']

sum_customer_2023_12 = merged_customer_long_medium['贷款余额_2023_12'].sum()
sum_customer_2024_06 = merged_customer_long_medium['贷款余额_2024_06'].sum()
sum_customer_increase = merged_customer_long_medium['增量'].sum()
new_row_customer = pd.DataFrame([['/','/','(合计)', sum_customer_2023_12, sum_customer_2024_06, sum_customer_increase]], columns=merged_customer_long_medium.columns)
merged_customer_long_medium = pd.concat([merged_customer_long_medium, new_row_customer], ignore_index=True)

sum_department_2023_12 = merged_department_long_medium['贷款余额_2023_12'].sum()
sum_department_2024_06 = merged_department_long_medium['贷款余额_2024_06'].sum()
sum_department_increase = merged_department_long_medium['增量'].sum()
new_row_department = pd.DataFrame([['(合计)', sum_department_2023_12, sum_department_2024_06, sum_department_increase]], columns=merged_department_long_medium.columns)
merged_department_long_medium = pd.concat([merged_department_long_medium, new_row_department], ignore_index=True)

#save
merged_customer_long_medium.to_excel('制造业中长期（客户）.xlsx', engine='openpyxl', index=False)
merged_department_long_medium.to_excel('制造业中长期（支行部门）.xlsx', engine='openpyxl', index=False)