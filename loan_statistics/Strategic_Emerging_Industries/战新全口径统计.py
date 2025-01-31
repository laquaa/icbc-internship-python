import pandas as pd

#Read
loan_2023_12_xls = pd.read_excel('二营12月末--提供.xls', sheet_name='Sheet1')
loan_2023_12_xls.to_excel('二营12月末--提供.xlsx', engine='openpyxl', index=False)
loan_2023_12 = pd.read_excel('二营12月末--提供.xlsx', sheet_name='Sheet1')
loan_2024_06 = pd.read_excel('战新完成数字明细202406--提供.xlsx', sheet_name='Sheet6')

#Select
loan_2023_12 = loan_2023_12[['客户全称','客户经理','部门','投向战新余额（折人民币）']]
loan_2024_06 = loan_2024_06[['客户全称','客户经理','部门','投向战新余额（折人民币）']]

#customer
customer_2023_12 = loan_2023_12.groupby(['客户全称'], as_index=False).agg({'部门':'first',
                                                                        '客户经理':'first',
                                                                        '投向战新余额（折人民币）': 'sum'})
customer_2024_06 = loan_2024_06.groupby(['客户全称'], as_index=False).agg({'部门':'first',
                                                                        '客户经理':'first',
                                                                        '投向战新余额（折人民币）': 'sum'})

#department
department_2023_12 = loan_2023_12.groupby(['部门'], as_index=False).agg({'投向战新余额（折人民币）': 'sum'})
department_2024_06 = loan_2024_06.groupby(['部门'], as_index=False).agg({'投向战新余额（折人民币）': 'sum'})

#merge
customer_2023_12_final = customer_2023_12.rename(columns={'投向战新余额（折人民币）': '投向战新余额（折人民币）_2023_12'})
customer_2024_06_final = customer_2024_06.rename(columns={'投向战新余额（折人民币）': '投向战新余额（折人民币）_2024_06'})
merged_customer = pd.merge(customer_2023_12_final, customer_2024_06_final, on=['客户全称','部门','客户经理'], how='outer')

department_2023_12_final = department_2023_12.rename(columns={'投向战新余额（折人民币）': '投向战新余额（折人民币）_2023_12'})
department_2024_06_final = department_2024_06.rename(columns={'投向战新余额（折人民币）': '投向战新余额（折人民币）_2024_06'})
merged_department = pd.merge(department_2023_12_final, department_2024_06_final, on='部门', how='outer')

merged_customer = merged_customer.fillna(0)
merged_department = merged_department.fillna(0)

merged_customer['增量'] = merged_customer['投向战新余额（折人民币）_2024_06'] - merged_customer['投向战新余额（折人民币）_2023_12']
merged_department['增量'] = merged_department['投向战新余额（折人民币）_2024_06'] - merged_department['投向战新余额（折人民币）_2023_12']

sum_customer_2023_12 = merged_customer['投向战新余额（折人民币）_2023_12'].sum()
sum_customer_2024_06 = merged_customer['投向战新余额（折人民币）_2024_06'].sum()
sum_customer_increase = merged_customer['增量'].sum()
new_row_customer = pd.DataFrame([['/','/','(合计)', sum_customer_2023_12, sum_customer_2024_06, sum_customer_increase]], columns=merged_customer.columns)
merged_customer = pd.concat([merged_customer, new_row_customer], ignore_index=True)

sum_department_2023_12 = merged_department['投向战新余额（折人民币）_2023_12'].sum()
sum_department_2024_06 = merged_department['投向战新余额（折人民币）_2024_06'].sum()
sum_department_increase = merged_department['增量'].sum()
new_row_department = pd.DataFrame([['(合计)', sum_department_2023_12, sum_department_2024_06, sum_department_increase]], columns=merged_department.columns)
merged_department = pd.concat([merged_department, new_row_department], ignore_index=True)

#save
merged_customer.to_excel('战新全口径（客户）.xlsx', engine='openpyxl', index=False)
merged_department.to_excel('战新全口径（支行部门）.xlsx', engine='openpyxl', index=False)