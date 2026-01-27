import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import os

# ================= 配置区域 =================
# 你的 Excel 文件名
INPUT_FILE = "name.xlsx" 

# Excel 中存放 物质名称 的列名 (请确保和Excel里一致)
COL_NAME = "Name"

# 结果将写入的新列名
COL_RESULT = "Found CAS"
# ===========================================

def get_cas_by_name(name):
    """
    根据物质名称去 NIST 搜索 CAS 号。
    """
    if pd.isna(name) or str(name).strip() == "":
        return "\\"
    
    clean_name = str(name).strip()
    # NIST 的名称搜索 URL 结构
    # 这里的 options=On-off 表示精确匹配逻辑，通常直接搜 Name 即可
    url = "https://webbook.nist.gov/cgi/cbook.cgi"
    params = {
        "Name": clean_name,
        "Units": "SI"
    }
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }

    try:
        # 发送请求
        response = requests.get(url, params=params, headers=headers, timeout=15)
        
        # 如果找不到页面，直接返回
        if response.status_code != 200:
            return "Connection Error"
            
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 1. 检查是否进入了“搜索结果列表”页面 (即没有直接跳转到物质详情页)
        # 如果网页标题包含 "Search Results"，说明名字有歧义或者没完全匹配
        if "Search Results" in soup.title.text:
            # 这种情况下，通常我们尝试找第一个链接，或者直接标记为“需人工确认”
            # 为了简单，这里我们只抓取完全匹配进入详情页的情况
            return "Ambiguous/List Found"
            
        # 2. 检查是否是“Name Not Found”
        if "Name Not Found" in response.text:
            return "Not Found"

        # 3. 在详情页寻找 "CAS Registry Number:"
        # NIST 页面结构通常是: <li><strong>CAS Registry Number:</strong> 100-52-7</li>
        cas_label = soup.find('strong', string="CAS Registry Number:")
        
        if cas_label:
            # 获取 label 后面的文本
            cas_text = cas_label.next_sibling
            if cas_text:
                return cas_text.strip()
        
        # 备选方案：如果上面的结构变了，尝试搜索文本
        # 有时候页面结构不同，防止漏抓
        all_text = soup.get_text()
        if "CAS Registry Number:" in all_text:
             # 简单的文本截取兜底
             import re
             match = re.search(r'CAS Registry Number:\s*([\d\-]+)', all_text)
             if match:
                 return match.group(1)

        return "Not Found in Page"

    except Exception as e:
        print(f"Error processing {name}: {e}")
        return "Error"

def main():
    print(f"Reading file: {INPUT_FILE} ...")
    
    if not os.path.exists(INPUT_FILE):
        print(f"Error: File not found {INPUT_FILE}")
        return

    try:
        df = pd.read_excel(INPUT_FILE, dtype=str) # 全部按字符串读取，防止数字名称变型
    except Exception as e:
        print(f"Excel read error: {e}")
        return

    # 检查列名
    if COL_NAME not in df.columns:
        print(f"Error: Column '{COL_NAME}' not found.")
        print(f"Current columns: {list(df.columns)}")
        return

    print("Searching for CAS numbers... (Please wait)")

    total_rows = len(df)
    for index, row in df.iterrows():
        name = row[COL_NAME]
        
        print(f"[{index+1}/{total_rows}] Searching: {name} ... ", end="", flush=True)
        
        result = get_cas_by_name(name)
        
        df.at[index, COL_RESULT] = result
        print(f"Found: {result}")
        
        # 礼貌爬虫，防止封IP
        time.sleep(1)

    # 保存结果
    output_filename = INPUT_FILE.replace(".xlsx", "_with_cas.xlsx")
    df.to_excel(output_filename, index=False)
    
    print("-" * 30)
    print(f"Done! Saved to: {output_filename}")
    print("-" * 30)

if __name__ == "__main__":
    main()