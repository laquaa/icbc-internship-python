import pandas as pd

origin = pd.read_excel('所属集团信息.xlsx', sheet_name='Sheet1')
filtered_origin = origin[origin['所属集团名称'] != '/']

filtered_origin.to_excel('集团信息汇总.xlsx', index=False)
#没有英文公司