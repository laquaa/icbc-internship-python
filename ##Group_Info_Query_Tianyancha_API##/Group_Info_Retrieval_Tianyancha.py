import pandas as pd
import requests
import urllib.parse

origin = pd.read_excel('20240101-0630.xls', sheet_name='20240101-0630')
origin.to_excel('20240101-0630.xlsx', engine='openpyxl', index=False)
origin = pd.read_excel('20240101-0630.xlsx', sheet_name='Sheet1')

company_name = origin['户名'].unique()
company_name_list = company_name.tolist()

groupRename = []
groupUUID = []

token = "TIANYANCHA_API_TOKEN"

for name in company_name_list:
    company_name = name
    encoded_company_name = urllib.parse.quote(name)
    url = f"http://open.api.tianyancha.com/services/open/group/base?keyword={name}"
    headers = {'Authorization': token}

    response = requests.get(url, headers=headers)
    data = response.json()
    error_code = data.get('error_code')
    if error_code == 0:

        group_info = data.get('result', {})
        group_rename = group_info.get('groupRename')
        groupRename.append(group_rename)
        group_uuid = group_info.get('groupUUID')
        groupUUID.append(group_uuid)

    else:
        if error_code == 300006:
            print('余额不足')
        groupRename.append('/')
        groupUUID.append('/')


result = {
    '户名': company_name_list,
    '所属集团名称': groupRename,
    '所属集团UID': groupUUID
}
df = pd.DataFrame(result)
df.to_excel('集团信息获取（天眼查）.xlsx', index=False)