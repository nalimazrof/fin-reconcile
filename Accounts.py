import os
import requests

if not os.path.exists('.simplefin_url'):
    print("❌ 找不到钥匙文件！")
    exit()

with open('.simplefin_url', 'r') as f:
    access_url = f.read().strip()

scheme, rest = access_url.split('//', 1)
auth, rest = rest.split('@', 1)
url = scheme + '//' + rest + '/accounts'
username, password = auth.split(':', 1)

print("🌐 正在拉取 API 花名册...\n")
response = requests.get(url, auth=(username, password))
data = response.json()

accounts = data.get('accounts', [])
print(f"✅ 成功获取到 {len(accounts)} 个账户。具体名单如下：")
print("=" * 50)

for i, acc in enumerate(accounts, 1):
    org_name = acc.get('org', {}).get('name', '未知机构')
    acc_name = acc.get('name', '未知账户名')
    print(f"{i:02d}. 机构: 【{org_name}】 | 账户: 【{acc_name}】")

print("=" * 50)