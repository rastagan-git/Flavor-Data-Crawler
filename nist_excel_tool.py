import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import time
import os

# ================= 配置区域 (请在这里修改) =================
# 你的 Excel 文件名 (必须带后缀, 如 .xlsx)
INPUT_FILE = "data.xlsx" 

# Excel 中存放 CAS 号的列名 (请确保Excel里第一行是列名)
COL_CAS = "CAS Number"

# Excel 中存放 计算保留指数 的列名
COL_CALC_RI = "Calculated RI"

# 结果将写入的新列名
COL_RESULT = "NIST RI"
# ========================================================

def get_nist_ri(cas, calc_ri):
    """
    根据 CAS 号去 NIST 查找最接近的 RI 值。
    """
    # 如果 CAS 为空或无效，直接返回 \
    if pd.isna(cas) or str(cas).strip() == "":
        return "\\"
    
    clean_cas = str(cas).strip().replace('-', '')
    url = f"https://webbook.nist.gov/cgi/cbook.cgi?ID=C{clean_cas}&Units=SI&Mask=2000"
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }

    try:
        # 发送请求
        response = requests.get(url, headers=headers, timeout=15)
        if response.status_code != 200:
            return "\\"
            
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 定位目标表格
        target_title = "Normal alkane RI, non-polar column, custom temperature program"
        target_table = soup.find('table', attrs={'aria-label': target_title})

        # 如果通过 aria-label 找不到，尝试通过文本查找
        if not target_table:
            text_matches = soup.find_all(string=re.compile(re.escape(target_title)))
            for match in text_matches:
                parent = match.parent
                next_table = parent.find_next('table')
                if next_table:
                    target_table = next_table
                    break
        
        if not target_table:
            return "\\"

        # 找到 "I" (保留指数) 所在的列
        header_row = target_table.find('tr')
        headers_text = [th.get_text(strip=True) for th in header_row.find_all(['th', 'td'])]
        
        try:
            ri_col_index = headers_text.index('I')
        except ValueError:
            return "\\"

        ri_values = []
        
        # 遍历所有行提取数据
        rows = target_table.find_all('tr')[1:]
        for row in rows:
            cols = row.find_all(['td', 'th'])
            if len(cols) > ri_col_index:
                val_text = cols[ri_col_index].get_text(strip=True)
                # 清洗非数字字符
                val_clean = re.sub(r'[^\d\.]', '', val_text)
                try:
                    if val_clean:
                        ri_values.append(float(val_clean))
                except ValueError:
                    continue

        if not ri_values:
            return "\\"

        # 找出与计算值 (calc_ri) 差值最小的那个数
        closest_ri = min(ri_values, key=lambda x: abs(x - float(calc_ri)))
        
        # 格式化返回（如果是整数显示整数）
        if closest_ri.is_integer():
            return int(closest_ri)
        else:
            return closest_ri

    except Exception as e:
        print(f"Error processing CAS {cas}: {e}")
        return "\\"

def main():
    print(f"正在读取文件: {INPUT_FILE} ...")
    
    if not os.path.exists(INPUT_FILE):
        print(f"错误: 找不到文件 {INPUT_FILE}")
        return

    # 读取 Excel，强制将 CAS 列作为字符串读取，防止 Excel 自动转成日期
    try:
        df = pd.read_excel(INPUT_FILE, dtype={COL_CAS: str})
    except Exception as e:
        print(f"读取 Excel 失败: {e}")
        return

    # 检查列名是否存在
    if COL_CAS not in df.columns or COL_CALC_RI not in df.columns:
        print(f"错误: Excel 中找不到列名 '{COL_CAS}' 或 '{COL_CALC_RI}'。")
        print(f"当前 Excel 列名: {list(df.columns)}")
        return

    print("开始处理数据，请稍候...")
    print("注意: 为防止被 NIST 网站屏蔽，每条数据间会有 1 秒的间隔。")

    # 遍历 DataFrame 处理每一行
    total_rows = len(df)
    for index, row in df.iterrows():
        cas = row[COL_CAS]
        calc_ri = row[COL_CALC_RI]
        
        print(f"[{index+1}/{total_rows}] 处理 CAS: {cas} (计算值: {calc_ri}) ... ", end="", flush=True)
        
        result = get_nist_ri(cas, calc_ri)
        
        # 写入结果
        df.at[index, COL_RESULT] = result
        
        print(f"结果: {result}")
        
        # 休眠 1 秒，礼貌爬虫
        time.sleep(1)

    # 保存结果
    output_filename = INPUT_FILE.replace(".xlsx", "_result.xlsx")
    df.to_excel(output_filename, index=False)
    
    print("-" * 30)
    print(f"完成！结果已保存至: {output_filename}")
    print("-" * 30)

if __name__ == "__main__":
    main()