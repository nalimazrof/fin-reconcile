import os
import glob
import pandas as pd
import re
import time
import requests
import base64
import warnings
from datetime import datetime

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ==========================================
# 1. 基础与路径配置
# ==========================================
EXCEL_FILE = "Finances.xlsx"
CSV_DIR = "CSV for validation"
YEAR_TAB = "2026"
ASSETS_TAB = "US assets"
MAX_DAYS_TOLERANCE = 5 
TOKEN_FILE = ".simplefin_url" # 用于安全保存 Access URL 的本地文件

def get_date_range():
    print("\n📅 请输入需要核对的时间范围 (格式: YYYY-MM-DD)")
    start_str = input("请输入开始日期 (例如 2026-04-15): ").strip()
    end_str = input("请输入结束日期 (例如 2026-05-15): ").strip()
    
    try:
        start_date = pd.to_datetime(start_str)
        end_date = pd.to_datetime(end_str)
        return start_date, end_date
    except Exception as e:
        print("❌ 日期格式输入有误，请按 YYYY-MM-DD 格式重新运行。")
        exit()

def get_or_claim_access_url():
    """读取本地凭证，或通过 Setup Token 向服务器兑换新凭证"""
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'r') as f:
            return f.read().strip()

    print("\n⚠️ 未在本地找到 API 凭证。")
    setup_token = input("👉 请粘贴在 SimpleFIN 生成的 Setup Token (直接回车则跳过并使用纯CSV模式): ").strip()
    
    if not setup_token:
        return None

    try:
        # 1. 解码 Setup Token (本质是 base64 编码的 Claim URL)
        claim_url = base64.b64decode(setup_token).decode('utf-8')
        print("🔄 正在向 SimpleFIN 兑换永久 Access URL...")
        
        # 2. 发起 POST 请求换取永久链接
        response = requests.post(claim_url, timeout=10)
        response.raise_for_status()
        access_url = response.text.strip()
        
        # 3. 保存至本地隐藏文件
        with open(TOKEN_FILE, 'w') as f:
            f.write(access_url)
            
        print("✅ 兑换成功！永久凭证已安全保存在本目录的 .simplefin_url 文件中。")
        return access_url
    except Exception as e:
        print(f"❌ Setup Token 兑换失败 (该Token可能已失效，请去官网重新生成): {e}")
        return None

def fetch_simplefin_data(start_date, end_date):
    """请求网关数据 (采用官方 Basic Auth 标准)"""
    access_url = get_or_claim_access_url()
    if not access_url:
        print("⚠️ 降级提示：将仅扫描本地 CSV 文件夹。")
        return []

    print("\n🌐 正在连接 SimpleFIN 网关...")
    try:
        # 官方标准的 URL 拆解与 Auth 配置方式
        scheme, rest = access_url.split('//', 1)
        auth, rest = rest.split('@', 1)
        url = scheme + '//' + rest + '/accounts'
        username, password = auth.split(':', 1)

        start_ts = int(time.mktime(start_date.timetuple()))
        end_ts = int(time.mktime(end_date.timetuple()))

        response = requests.get(
            url,
            auth=(username, password),
            params={"start-date": start_ts, "end-date": end_ts}
        )
        response.raise_for_status()
        data = response.json()
        print(f"✅ 成功连接云端！共获取 {len(data.get('accounts', []))} 个账户数据。")
        return data.get('accounts', [])
    except Exception as e:
        print(f"❌ API 连接失败 ({e})。\n🔄 已自动降级为纯本地 CSV 备用模式。")
        return []

def extract_last_4_digits(val):
    if pd.isna(val) or not val: return ""
    s = str(val).split('.')[0] 
    s = re.sub(r'\D', '', s)
    return s[-4:] if len(s) >= 4 else ""

def build_account_mapping(excel_path):
    try:
        df_assets = pd.read_excel(excel_path, sheet_name=ASSETS_TAB)
    except Exception:
        print(f"❌ 找不到 '{ASSETS_TAB}' 标签页，无法建立自动映射。")
        return None, None

    account_map = {}
    all_account_names = []

    for _, row in df_assets.iterrows():
        name = str(row.get('Name', '')).strip()
        if not name or name == 'nan': continue
            
        all_account_names.append(name)
        acct_type = str(row.get('Type', '')).strip().lower()
        card_4 = extract_last_4_digits(row.get('Card#'))
        acc_4 = extract_last_4_digits(row.get('Account#'))

        suffix = card_4 if acct_type == 'credit' and card_4 else acc_4 if acct_type != 'credit' and acc_4 else ""
        if suffix:
            account_map[suffix] = name
            
    # 【新增这一行】：按名字长度从长到短排序，先匹配 Flex 再匹配 Freedom
    all_account_names.sort(key=len, reverse=True) 
    return all_account_names, account_map

def standardize_bank_csv(csv_path):
    df = pd.read_csv(csv_path, index_col=False)
    df.columns = df.columns.str.strip()

    date_col = next((c for c in ['Posting Date', 'Trans. Date', 'Transaction Date', 'Date'] if c in df.columns), None)
    amount_col = next((c for c in ['Amount', 'Transaction Amount', 'Debit', 'Credit'] if c in df.columns), None)

    if not date_col or not amount_col: return None
    desc_col = next((c for c in ['Description', 'Payee', 'Details'] if c in df.columns), df.columns[0])

    df = df.rename(columns={date_col: 'Date', amount_col: 'Amount', desc_col: 'Description'})
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    
    df['Amount'] = df['Amount'].astype(str).str.replace(r'[^\d\.\-]', '', regex=True)
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce').abs()
    df = df.dropna(subset=['Date', 'Amount'])
    return df[['Date', 'Amount', 'Description']]

def run_reconciliation():
    print("="*60)
    print("🚀 个人财务全自动对账系统 (双轨混合引擎版)")
    print("="*60)

    if not os.path.exists(EXCEL_FILE):
        print(f"❌ 找不到 Excel 账本: {EXCEL_FILE}")
        return

    print(f"🔄 正在读取 Excel 建立账户尾号指纹库...")
    all_account_names, account_map = build_account_mapping(EXCEL_FILE)
    if not all_account_names: return

    start_date, end_date = get_date_range()
    bank_data_pool = {}

    # --- 阶段 A：填充 API 数据 ---
    sf_accounts = fetch_simplefin_data(start_date, end_date)
    for acc_data in sf_accounts:
        full_sf_name = f"{acc_data.get('org', {}).get('name', '')} {acc_data.get('name', '')}"
        matched_account = None

        for acc_name in all_account_names:
            if acc_name.lower() in full_sf_name.lower():
                matched_account = acc_name
                break
        if not matched_account:
            digits_groups = re.findall(r'\d+', acc_data.get('name', ''))
            if digits_groups:
                for suffix, acc_name in account_map.items():
                    if digits_groups[-1].endswith(suffix):
                        matched_account = acc_name
                        break
        
        if matched_account and acc_data.get('transactions'):
            df_api = pd.DataFrame(acc_data['transactions'])
            df_api['Date'] = pd.to_datetime(df_api['posted'], unit='s')
            df_api['Amount'] = pd.to_numeric(df_api['amount'], errors='coerce').abs()
            df_api['Description'] = df_api['payee'].fillna('') + " " + df_api['description'].fillna('')
            
            mask = (df_api['Date'] >= start_date) & (df_api['Date'] <= end_date)
            bank_data_pool[matched_account] = { 'source': 'API', 'df': df_api.loc[mask] }

    # --- 阶段 B：扫描本地 CSV 并覆盖 (Override) ---
    if not os.path.exists(CSV_DIR):
        os.makedirs(CSV_DIR)
        
    csv_files = glob.glob(os.path.join(CSV_DIR, "*.csv"))
    for csv_file in csv_files:
        filename = os.path.basename(csv_file)
        matched_account = None

        for acc_name in all_account_names:
            if acc_name.lower() in filename.lower():
                matched_account = acc_name
                break
        if not matched_account:
            digits_groups = re.findall(r'\d+', filename)
            if digits_groups:
                for suffix, acc_name in account_map.items():
                    if digits_groups[0].endswith(suffix):
                        matched_account = acc_name
                        break

        if matched_account:
            df_csv = standardize_bank_csv(csv_file)
            if df_csv is not None and not df_csv.empty:
                mask = (df_csv['Date'] >= start_date) & (df_csv['Date'] <= end_date)
                bank_data_pool[matched_account] = { 'source': '本地 CSV', 'df': df_csv.loc[mask] }

    # --- 阶段 C：执行核对 ---
    if not bank_data_pool:
        print("\n⚠️ 没有获取到任何有效数据 (API无返回且本地CSV为空)。")
        return

    df_excel = pd.read_excel(EXCEL_FILE, sheet_name=YEAR_TAB)
    df_excel['Date'] = pd.to_datetime(df_excel['Date'])
    df_excel['Amount'] = pd.to_numeric(df_excel['Amount'], errors='coerce').abs()
    df_excel['from Account'] = df_excel['from Account'].fillna('')
    df_excel['to Account'] = df_excel['to Account'].fillna('')

    for matched_account, data_info in bank_data_pool.items():
        source_label = data_info['source']
        df_bank = data_info['df']
        
        print("-" * 60)
        print(f"⚙️ 核对账户: 【{matched_account}】 (数据源: {source_label})")
        
        excel_start = start_date - pd.Timedelta(days=MAX_DAYS_TOLERANCE)
        excel_end = end_date + pd.Timedelta(days=MAX_DAYS_TOLERANCE)
        
        excel_mask = (
            ((df_excel['from Account'] == matched_account) | (df_excel['to Account'] == matched_account)) &
            (df_excel['Date'] >= excel_start) & (df_excel['Date'] <= excel_end)
        )
        df_excel_acc = df_excel.loc[excel_mask].copy()
        
        bank_records = df_bank.to_dict('records')
        excel_records = df_excel_acc.to_dict('records')
        
        matched_e_idx, matched_b_idx = set(), set()

        for b_idx, b_row in enumerate(bank_records):
            for e_idx, e_row in enumerate(excel_records):
                if e_idx in matched_e_idx: continue 
                
                if abs(b_row['Amount'] - e_row['Amount']) < 0.01:
                    if abs((b_row['Date'] - e_row['Date']).days) <= MAX_DAYS_TOLERANCE:
                        matched_e_idx.add(e_idx)
                        matched_b_idx.add(b_idx)
                        break

        excel_only = [e for i, e in enumerate(excel_records) if i not in matched_e_idx and start_date <= e['Date'] <= end_date]
        bank_only = [b for i, b in enumerate(bank_records) if i not in matched_b_idx]

        if not excel_only and not bank_only:
            print("  ✅ 完美匹配！")
        else:
            if excel_only:
                df_e = pd.DataFrame(excel_only)
                df_e['Date'] = df_e['Date'].dt.date
                print("\n  ❌ 异常 A: Excel 有，银行没找到：")
                print(df_e[['Date', 'Amount', 'Comments']].to_string(index=False))
            if bank_only:
                df_b = pd.DataFrame(bank_only)
                df_b['Date'] = df_b['Date'].dt.date
                print("\n  ❌ 异常 B: 银行有，Excel 没记：")
                print(df_b[['Date', 'Amount', 'Description']].to_string(index=False))

if __name__ == "__main__":
    run_reconciliation()