import os
import requests
import time
import re
from datetime import datetime

# 1. 自动读取本地永久钥匙
TOKEN_FILE = '.simplefin_url'
if not os.path.exists(TOKEN_FILE):
    print("❌ 找不到钥匙文件！请先确保主程序已经成功运行并保存了凭证。")
    exit()

with open(TOKEN_FILE, 'r') as f:
    access_url = f.read().strip()

# 2. 获取用户输入
print("="*60)
print("🔍 专属 API 银行流水快速查询器")
print("="*60)

target_suffix = input("👉 请输入要查询的账户后四位 (例如 4316): ").strip()
start_str = input("👉 请输入开始日期 (格式 YYYY-MM-DD): ").strip()
end_str = input("👉 请输入结束日期 (格式 YYYY-MM-DD): ").strip()

try:
    # 转换日期并计算时间戳 (结束日期强制设为当天的 23:59:59 以包含全天)
    start_date = datetime.strptime(start_str, "%Y-%m-%d")
    end_date = datetime.strptime(end_str, "%Y-%m-%d").replace(hour=23, minute=59, second=59)
    start_ts = int(time.mktime(start_date.timetuple()))
    end_ts = int(time.mktime(end_date.timetuple()))
except ValueError:
    print("\n❌ 日期格式输入有误，请按 YYYY-MM-DD 的格式重新运行。")
    exit()

# 3. 组装请求
scheme, rest = access_url.split('//', 1)
auth, rest = rest.split('@', 1)
url = scheme + '//' + rest + '/accounts'
username, password = auth.split(':', 1)

print("\n🌐 正在向 SimpleFIN API 发送查询请求...")
try:
    response = requests.get(
        url, 
        auth=(username, password),
        params={"start-date": start_ts, "end-date": end_ts},
        timeout=10
    )
    response.raise_for_status()
    data = response.json()
except Exception as e:
    print(f"❌ API 请求失败: {e}")
    exit()

# 4. 寻找匹配账户并打印
found_match = False

for acc in data.get('accounts', []):
    acc_name = acc.get('name', '')
    org_name = acc.get('org', {}).get('name', '未知机构')
    
    # 提取账户名中的所有数字组
    digits_groups = re.findall(r'\d+', acc_name)
    
    # 判断输入的后四位是否匹配该账户 (为了严谨，检查数字组是否以目标尾号结尾)
    is_match = False
    for group in digits_groups:
        if group.endswith(target_suffix):
            is_match = True
            break
            
    # 如果数字匹配不上，直接看字符串里有没有包含 (兜底)
    if not is_match and target_suffix in acc_name:
        is_match = True
        
    if is_match:
        found_match = True
        print("\n✅ 匹配成功！")
        print(f"🏛️ 机构: 【{org_name}】")
        print(f"💳 账户: 【{acc_name}】")
        print("-" * 60)
        
        transactions = acc.get('transactions', [])
        if not transactions:
            print("  ℹ️ 在您输入的日期范围内，该账户没有产生任何已入账的流水。")
        else:
            print(f"{'入账日期':<12} | {'金额':<10} | {'商户/描述'}")
            print("-" * 60)
            
            # 按日期排序打印，方便阅读
            transactions.sort(key=lambda x: x['posted'])
            for t in transactions:
                # 把时间戳转回可读日期
                dt = datetime.fromtimestamp(t['posted']).strftime('%Y-%m-%d')
                amount = float(t['amount'])
                desc = t.get('payee', '') or t.get('description', '')
                
                # 截断过长的描述以保持排版整洁
                if len(desc) > 35:
                    desc = desc[:32] + "..."
                    
                print(f"{dt:<12} | {amount:>8.2f}   | {desc}")
                
            print("-" * 60)
            print(f"📊 统计: 共查询到 {len(transactions)} 笔流水。")

if not found_match:
    print(f"\n⚠️ 在 API 中未找到尾号为 '{target_suffix}' 的账户。")
    print("💡 提示：请确保该账户已在 SimpleFIN 成功连接，且尾号输入正确。")