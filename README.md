# Wei Long North America Management Review Dashboard

这是 Wei Long North America 的管理层经营看板。项目将 Amazon、TikTok、Temu 三个平台的订单/结算数据清洗成统一经营数据表，并用 Streamlit 生成可交互网页。

## 项目结构

```text
app.py                    Streamlit Dashboard 入口
run_pipeline.py           本地数据清洗入口
export_report.py          静态 HTML 报告导出
src/                      数据清洗 pipeline
config:/                  本地配置文件，不建议直接上传公开 GitHub
raw_data:/                本地原始数据，不要上传 GitHub
output/                   本地输出文件
sample_master_data.xlsx   Streamlit Cloud Demo 兜底数据
requirements.txt          Streamlit Cloud Python 依赖
.gitignore                隐私与输出文件忽略规则
```

## 本地运行

首次本地使用时，先安装依赖：

```bash
pip install -r requirements.txt
```

如果已经有 `output/master_data.xlsx`，Dashboard 会优先读取它：

```bash
streamlit run app.py
```

如果还没有 `output/master_data.xlsx`，先运行清洗 pipeline：

```bash
python run_pipeline.py
streamlit run app.py
```

## 静态报告

```bash
python export_report.py
```

然后打开：

```text
output/static_report.html
```

## Streamlit Community Cloud 部署

目标：生成一个老板可以直接打开的 `https://xxx.streamlit.app` 网页链接。

### 1. 上传到 GitHub

新建一个 GitHub repository，然后把本项目代码上传。

建议上传：

```text
app.py
run_pipeline.py
export_report.py
requirements.txt
README.md
.gitignore
src/
sample_master_data.xlsx
```

如果要展示真实经营数据，请先确认 `output/master_data.xlsx` 已脱敏或允许上传，然后手动强制加入 Git：

```bash
git add -f output/master_data.xlsx
```

不建议上传：

```text
raw_data/
raw_data:/
*.csv
未脱敏的 *.xlsx
output/unmatched_sku.xlsx
config 中包含敏感映射或目标的数据文件
```

### 2. 在 Streamlit Community Cloud 创建应用

1. 打开 Streamlit Community Cloud。
2. 点击 New app。
3. 选择你的 GitHub repo。
4. Branch 选择 `main`。
5. Main file path 填：

```text
app.py
```

6. 点击 Deploy。

部署成功后，Streamlit 会生成一个类似这样的链接：

```text
https://your-dashboard-name.streamlit.app
```

把这个链接发给老板即可直接打开。

## 数据读取逻辑

`app.py` 在启动时按以下顺序读取数据：

1. 优先读取 `output/master_data.xlsx`
2. 如果不存在，读取 `sample_master_data.xlsx`
3. 如果两个文件都不存在，页面会提示先本地运行：

```bash
python run_pipeline.py
```

## 核心口径

- `Fact_Order`：订单粒度，用于订单数、总 GMV、平台 GMV、AOV。
- `Fact_Product`：SKU 粒度，用于商品、口味、规格、组包分析。
- `Product_Expanded`：组包拆解后的商品分析表。
- 平台 SKU 只作为映射键，不作为最终商品分析维度。
- 订单数始终使用 `distinct(order_id)`。
- 无法识别 SKU 会输出到 `output/unmatched_sku.xlsx`，不会自动猜测。

## Dashboard 页面

- Business Overview
- Product Performance
- Geo Analysis
- Bundle Performance
- Product Quadrant
