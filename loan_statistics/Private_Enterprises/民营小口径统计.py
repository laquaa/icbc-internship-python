import pandas as pd

#Read
loan_2023_12 = pd.read_excel('202312--202406民营数据明细.xlsx', sheet_name='202312')
loan_2024_06 = pd.read_excel('202406民营数据明细.xlsx', sheet_name='Sheet1')

#Select
columns_to_check_23 = ['小企业', '大客', '涉房', '供应链']
columns_to_check_24 = ['小企业', '大客户中心', '涉房', '供应链']

filtered_2023_12 = loan_2023_12[loan_2023_12[columns_to_check_23].isna().all(axis=1)]
filtered_2024_06 = loan_2024_06[loan_2024_06[columns_to_check_24].isna().all(axis=1)]

#customer
customer_2023_12 = filtered_2023_12.groupby(['客户全称'], as_index=False).agg({'支行考核口径部门':'first',
                                                                        '客户经理':'first',
                                                                        '贷款余额': 'sum'})
customer_2024_06 = filtered_2024_06.groupby(['客户全称'], as_index=False).agg({'支行考核口径部门':'first',
                                                                        '客户经理':'first',
                                                                        '贷款余额': 'sum'})

#department
department_2023_12 = filtered_2023_12.groupby(['支行考核口径部门'], as_index=False).agg({'贷款余额': 'sum'})
department_2024_06 = filtered_2024_06.groupby(['支行考核口径部门'], as_index=False).agg({'贷款余额': 'sum'})

#merge
customer_2023_12 = customer_2023_12.rename(columns={'贷款余额': '贷款余额_2023_12'})
customer_2024_06 = customer_2024_06.rename(columns={'贷款余额': '贷款余额_2024_06'})
merged_customer = pd.merge(customer_2023_12, customer_2024_06, on=['客户全称','支行考核口径部门','客户经理'], how='outer')

department_2023_12 = department_2023_12.rename(columns={'贷款余额': '贷款余额_2023_12'})
department_2024_06 = department_2024_06.rename(columns={'贷款余额': '贷款余额_2024_06'})
merged_department = pd.merge(department_2023_12, department_2024_06, on='支行考核口径部门', how='outer')

merged_customer = merged_customer.fillna(0)
merged_department = merged_department.fillna(0)

merged_customer['增量'] = merged_customer['贷款余额_2024_06'] - merged_customer['贷款余额_2023_12']
merged_department['增量'] = merged_department['贷款余额_2024_06'] - merged_department['贷款余额_2023_12']

sum_customer_2023_12 = merged_customer['贷款余额_2023_12'].sum()
sum_customer_2024_06 = merged_customer['贷款余额_2024_06'].sum()
sum_customer_increase = merged_customer['增量'].sum()
new_row_customer = pd.DataFrame([['/','/','(合计)', sum_customer_2023_12, sum_customer_2024_06, sum_customer_increase]], columns=merged_customer.columns)
merged_customer = pd.concat([merged_customer, new_row_customer], ignore_index=True)

sum_department_2023_12 = merged_department['贷款余额_2023_12'].sum()
sum_department_2024_06 = merged_department['贷款余额_2024_06'].sum()
sum_department_increase = merged_department['增量'].sum()
new_row_department = pd.DataFrame([['(合计)', sum_department_2023_12, sum_department_2024_06, sum_department_increase]], columns=merged_department.columns)
merged_department = pd.concat([merged_department, new_row_department], ignore_index=True)

#save
merged_customer.to_excel('民营小口径（客户）.xlsx', engine='openpyxl', index=False)
merged_department.to_excel('民营小口径（支行部门）.xlsx', engine='openpyxl', index=False)