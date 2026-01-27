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
INPUT_FILE = "Odor.xlsx"   
COL_CAS = "CAS Number"     
# ===========================================

def init_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument("--log-level=3")
    options.add_argument("--disable-notifications")
    options.add_argument("--blink-settings=imagesEnabled=false") 
    
    try:
        driver_path = ChromeDriverManager(url="https://npmmirror.com/metadata/chromedriver/").install()
    except Exception:
        driver_path = ChromeDriverManager().install()
        
    service = Service(driver_path)
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def get_chemicalbook_data_with_retry(driver, cas):
    """
    带有手动重试机制的抓取函数
    """
    while True:
        try:
            # 尝试执行抓取
            return _get_data_logic(driver, cas)
        except Exception as e:
            # 如果抓取过程中出现严重错误（通常是被验证码拦截导致超时）
            print("\n" + "!"*50)
            print(f"【暂停警告】在抓取 {cas} 时遇到问题！")
            print("可能是出现了：1.人机验证码  2.网络断开  3.页面加载极慢")
            print(f"错误信息: {e}")
            print("!"*50)
            
            # 这里暂停，等待用户输入
            user_input = input(">>> 请去浏览器手动解决验证码，完成后在此处按回车重试 (输入 n 跳过此条): ")
            
            if user_input.lower().strip() == 'n':
                # 如果用户输入 n，则返回空数据，跳过这条
                return {
                    "CB_Odor_Desc": "\\", "CB_Odor_Threshold": "\\", "CB_Odor_Type": "\\"
                }
            else:
                # 否则继续循环，重新尝试
                print(">>> 正在重试...")
                continue

def _get_data_logic(driver, cas):
    """
    核心抓取逻辑（不含重试循环）
    """
    data = {
        "CB_Odor_Desc": "\\",       
        "CB_Odor_Threshold": "\\",  
        "CB_Odor_Type": "\\"        
    }

    # 1. 访问搜索页
    url = f"https://www.chemicalbook.com/Search.aspx?keyword={cas}"
    driver.get(url)
    
    # 这里我们缩短一点超时时间，方便更快触发“手动验证”提示
    wait = WebDriverWait(driver, 5) 

    # 2. 点击链接
    try:
        xpath_query = "//a[text()='化学性质' or contains(@href, 'ProductChemicalProperties')]"
        target_link = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_query)))
        target_link.click()
    except Exception:
        # 如果找不到链接，这里抛出异常，触发外层的“重试询问”
        # 只有确实搜不到的，用户手动确认后可以选择跳过
        raise Exception("未找到'化学性质'链接，可能是被验证码拦截或无数据")

    # 3. 切换窗口
    windows = driver.window_handles
    driver.switch_to.window(windows[-1])

    # 4. 等待详情页
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "th")))

    # 5. 抓取数据
    try:
        ele = driver.find_element(By.XPATH, "//th[contains(text(), '气味')]/following-sibling::td")
        data["CB_Odor_Desc"] = ele.text.strip()
    except: pass

    try:
        ele = driver.find_element(By.XPATH, "//th[contains(text(), '嗅觉阈值')]/following-sibling::td")
        data["CB_Odor_Threshold"] = ele.text.strip()
    except: pass

    try:
        ele = driver.find_element(By.XPATH, "//th[contains(text(), '香型')]/following-sibling::td")
        data["CB_Odor_Type"] = ele.text.strip()
    except: pass

    # 关闭窗口
    if len(windows) > 1:
        driver.close()
        driver.switch_to.window(windows[0])

    return data

def main():
    print(f"读取文件: {INPUT_FILE} ...")
    if not os.path.exists(INPUT_FILE):
        print(f"文件不存在: {INPUT_FILE}")
        return

    try:
        df = pd.read_excel(INPUT_FILE, dtype={COL_CAS: str})
    except Exception as e:
        print(f"Excel 读取失败: {e}")
        return

    print("正在启动浏览器...")
    driver = init_driver()

    try:
        total = len(df)
        print(f"开始任务，共 {total} 条")
        print("-" * 50)

        for index, row in df.iterrows():
            cas = row[COL_CAS]
            if pd.isna(cas) or str(cas).strip() == "": continue
            cas = str(cas).strip()
            
            print(f"[{index+1}/{total}] 搜索: {cas} ... ", end="", flush=True)
            
            # 使用带有重试机制的函数
            result = get_chemicalbook_data_with_retry(driver, cas)
            
            df.at[index, "CB_Odor_Desc"] = result["CB_Odor_Desc"]
            df.at[index, "CB_Odor_Threshold"] = result["CB_Odor_Threshold"]
            df.at[index, "CB_Odor_Type"] = result["CB_Odor_Type"]
            
            found = []
            if result["CB_Odor_Desc"] != "\\": found.append("气味")
            if result["CB_Odor_Threshold"] != "\\": found.append("阈值")
            if result["CB_Odor_Type"] != "\\": found.append("香型")
            
            if found:
                print(f"成功 ({', '.join(found)})")
            else:
                print("未找到数据 (或跳过)")

            time.sleep(1.5)

        output_file = INPUT_FILE.replace(".xlsx", "_cb_result.xlsx")
        df.to_excel(output_file, index=False)
        print(f"完成！保存至: {output_file}")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()