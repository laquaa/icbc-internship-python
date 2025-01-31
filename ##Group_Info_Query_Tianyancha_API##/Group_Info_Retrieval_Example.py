import requests
import urllib.parse

# token 可以从 数据中心 -> 我的接口 中获取
token = "TIANYANCHA_API_TOKEN"

# 查询的公司名称，确保正确编码
company_name = "北京百度网讯科技有限公司"
encoded_company_name = urllib.parse.quote(company_name)

# 构建 URL
url = f"http://open.api.tianyancha.com/services/open/group/base?keyword={encoded_company_name}"

# 构建请求头
headers = {'Authorization': token}

try:
    # 发送请求
    response = requests.get(url, headers=headers)

    # 检查响应状态码
    if response.status_code == 200:
        data = response.json()
        # 提取 groupRename 和 groupUUID
        group_info = data.get('result', {})
        group_rename = group_info.get('groupRename')
        group_uuid = group_info.get('groupUUID')

        # 打印提取的信息
        print(f"groupRename: {group_rename}")
        print(f"groupUUID: {group_uuid}")
    else:
        print(f"请求失败，状态码：{response.status_code}")
        print(response.text)
except requests.exceptions.RequestException as e:
    # 捕获请求异常
    print("请求发生错误:", e)
