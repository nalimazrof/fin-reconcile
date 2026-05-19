import requests
import json
import os
import time
from datetime import datetime, timedelta

# 1. 读取本地永久钥匙
if not os.path.exists('.simplefin_url'):
    print("❌ 找不到钥匙文件！请先运行一次 main.py。")
    exit()

with open('.simplefin_url', 'r') as f:
    access_url = f.read().strip()

# 2. 解析 URL
scheme, rest = access_url.split('//', 1)
auth, rest = rest.split('@', 1)
url = scheme + '//' + rest + '/accounts'
username, password = auth.split(':', 1)

# 3. 设定时间范围：当前时间往前推 60 天
end_date = datetime.now()
start_date = end_date - timedelta(days=60)
start_ts = int(time.mktime(start_date.timetuple()))
end_ts = int(time.mktime(end_date.timetuple()))

print(f"📅 强制查询范围: {start_date.strftime('%Y-%m-%d')} 至 {end_date.strftime('%Y-%m-%d')}")
print("🌐 正在连接 API 获取全量原始数据...\n")

# 将时间参数加入请求
params = {
    "start-date": start_ts,
    "end-date": end_ts
}
response = requests.get(url, auth=(username, password), params=params)
data = response.json()

# 4. 锁定 Citi 账户
found = False
for acc in data.get('accounts', []):
    org_name = acc.get('org', {}).get('name', '').upper()
    acc_name = acc.get('name', '').upper()
    
    if 'CITI' in org_name or 'CITI' in acc_name:
        found = True
        print(f"✅ 成功锁定目标！")
        print(f"🏛️ 机构: 【{acc.get('org', {}).get('name')}】")
        print(f"💳 账户: 【{acc.get('name')}】")
        print("=" * 60)
        print("👇 原始交易明细 (JSON 格式):")
        print("=" * 60)
        
        transactions = acc.get('transactions', [])
        if not transactions:
            print("  ℹ️ 即使加了长达 60 天的时间范围，API 依然返回 0 笔流水。")
            print("  👉 终极结论：聚合网关的爬虫存在抓取盲区，彻底漏掉了过渡期的数据。")
        else:
            print(json.dumps(transactions, indent=2))
            print(f"\n📊 统计：共抓取到 {len(transactions)} 笔流水。")
            
            # 帮助快速寻找有没有麦当劳
            mcdonalds_count = sum(1 for t in transactions if 'MCDONALD' in str(t.get('payee', '')).upper() or 'MCDONALD' in str(t.get('description', '')).upper())
            print(f"🍔 麦当劳匹配测试：在这个列表里找到了 {mcdonalds_count} 笔麦当劳。")
        print("=" * 60)

if not found:
    print("⚠️ 依旧没有找到匹配的账户。")