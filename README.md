# Flavor-Data-Crawler
Tools for extracting chemical sensory data
# Flavor & Chemical Properties Crawler Tool

这是一个基于 Python 的自动化数据采集工具集，用于从多个化学数据库中批量检索化合物的属性信息。

## 功能模块 (Modules)

1.  **NIST WebBook Crawler**
    * 根据 CAS 号自动查找保留指数 (Retention Index)。以及根据英文名称查找CAS号
2.  **MFFI Database Crawler** (上海交大风味数据库)
    * 抓取化合物的中文名、感官特征 (Sensory Characteristics) 和水中阈值。
3.  **ChemicalBook Crawler**
    * 自动化检索气味描述、嗅觉阈值 (Odor Threshold) 和香型 (Odor Type)。

## 环境依赖 (Requirements)

* Python 3.x
* Selenium
* Pandas
* Requests

## 使用方法 (Usage)
### 建立虚拟环境
本项目代码基于虚拟环境的状态下，建议直接使用本项目的虚拟环境代码

```
# 1. 创建名为 myenv 的虚拟环境
python -m venv myenv

# 2. 激活虚拟环境 (Windows)
myenv\Scripts\activate

# 3. 激活虚拟环境 (Mac/Linux)
# source myenv/bin/activate
```

### 安装依赖：
    
    pip install -r requirements.txt
    

### 功能一、根据cas号和计算出的保留指数查找标准保留指数
你的excel文件名= "data.xlsx" 

Excel 中存放 CAS 号的列名 (请确保Excel里第一行是列名)
COL_CAS = "CAS Number"

Excel 中存放 计算保留指数 的列名
COL_CALC_RI = "Calculated RI"

结果将写入的新列名
COL_RESULT = "NIST RI"

双击start1.bat

### 功能二、根据名称查找cas号
你的 Excel 文件名 = "name.xlsx" 

Excel 中存放 物质名称 的列名 (请确保和Excel里一致)
COL_NAME = "Name"

结果将写入的新列名
COL_RESULT = "Found CAS"

双击start2.bat

### 功能三、根据cas号查找中英文名称，气味，水中的嗅觉阈值
你的excel文件名 = "max.xlsx"

Excel 中存放 CAS 号的列名 (请确保Excel里第一行是列名)
COL_CAS = "CAS Number"

双击start3.bat

### 功能四、根据cas号查找嗅觉阈值，气味，香型（自带手动防人机验证）
你的excel文件名
INPUT_FILE = "Odor.xlsx"

Excel 中存放 CAS 号的列名 (请确保Excel里第一行是列名)
COL_CAS = "CAS Number"

双击start4.bat

## 信息来源网站
https://webbook.nist.gov/chemistry/
https://mffi.sjtu.edu.cn/database/search?value=Propanal%2C+2-methyl-&keyword=all
https://www.chemicalbook.com/ProductIndex.aspx

## 注意事项
本项目仅供学习参考，禁止用于商业用途，合理使用工具，不要给网站服务器带来太多负担
