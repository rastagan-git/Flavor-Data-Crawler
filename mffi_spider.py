import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import os

# ================= 配置区域 =================
INPUT_FILE = "max.xlsx"
COL_CAS = "CAS Number"
# ===========================================

def init_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized") 
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument("--log-level=3")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def get_mffi_data(driver, cas):
    # 构造结果字典
    data = {
        "Chinese Name": "\\",
        "English Name": "\\",
        "Sensory Characteristics": "\\",
        "In Water": "\\"
    }

    try:
        # 1. 访问网页
        url = f"https://mffi.sjtu.edu.cn/database/search?value={cas}&keyword=all"
        driver.get(url)

        # 2. 等待表格出现 (最多等10秒)
        wait = WebDriverWait(driver, 10)
        # 只要 tbody 里有行(tr)出现，就开始处理
        rows = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody tr")))
        
        # 3. 遍历每一行寻找匹配的数据
        for row in rows:
            # 获取这一行的所有格子
            cells = row.find_elements(By.TAG_NAME, "td")
            
            # 如果格子太少，说明是空行或格式不对，跳过
            if len(cells) < 7:
                continue
            
            # 获取这一行里显示的 CAS 号 (第4列，索引是3)
            row_cas_text = cells[3].text.strip()
            
            # 检查 CAS 号是否匹配 (只要包含就可以)
            if cas in row_cas_text or row_cas_text in cas:
                # 提取我们需要的数据
                data["Chinese Name"] = cells[0].text.strip()
                data["English Name"] = cells[1].text.strip()
                data["Sensory Characteristics"] = cells[5].text.strip()
                data["In Water"] = cells[6].text.strip()
                # 找到了就不找了，直接退出循环
                break

    except Exception as e:
        # 如果出错（比如超时没找到），什么都不做，直接返回空数据
        # 这样程序就不会崩溃
        pass

    return data

def main():
    print(f"读取文件: {INPUT_FILE} ...")
    
    if not os.path.exists(INPUT_FILE):
        print(f"[错误] 找不到文件 {INPUT_FILE}")
        return

    try:
        df = pd.read_excel(INPUT_FILE, dtype={COL_CAS: str})
    except Exception as e:
        print(f"[错误] Excel 读取失败: {e}")
        return

    print("正在启动浏览器...")
    driver = init_driver()

    try:
        total = len(df)
        print("-" * 50)
        print(f"开始任务，共 {total} 条数据")
        print("-" * 50)

        for index, row in df.iterrows():
            # 获取 CAS 号并转为字符串
            cas = row[COL_CAS]
            
            # 如果为空则跳过
            if pd.isna(cas) or str(cas).strip() == "":
                continue
            
            cas = str(cas).strip()
            
            print(f"[{index+1}/{total}] 搜索: {cas} ... ", end="", flush=True)
            
            # 执行抓取
            result = get_mffi_data(driver, cas)
            
            # 填入表格
            df.at[index, "Chinese Name"] = result["Chinese Name"]
            df.at[index, "English Name"] = result["English Name"]
            df.at[index, "Sensory Characteristics"] = result["Sensory Characteristics"]
            df.at[index, "In Water"] = result["In Water"]
            
            # 打印结果
            if result["Chinese Name"] != "\\":
                print(f"成功 -> {result['Chinese Name']}")
            else:
                print("未找到")

            time.sleep(1)

        # 保存结果
        output_file = INPUT_FILE.replace(".xlsx", "_mffi_result.xlsx")
        df.to_excel(output_file, index=False)
        print("-" * 50)
        print(f"全部完成！结果已保存为: {output_file}")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()