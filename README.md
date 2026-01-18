# 销售透视报表生成器（简体中文）

这个小工具可以把订单状态 Excel（多工作表）自动清洗并生成 Pivot 报表：
- 自动识别表头、自动匹配列名
- 自动规范化 Team（按规则）
- 输出 9 个工作表：**3 个工厂 × 3 个指标（Order Qty / SAH / Sales (USD)**）
- 每个工厂 sheet 顶部会先输出 **ALL DATA** 汇总表，然后按 **产品类型 + Team** 拆分多张表，便于查看/筛选

## 1) 安装

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

pip install -r requirements.txt
```

> 说明：如果你不需要处理 `.xls`（老格式），可以不装 `xlrd`。

## 2) 运行（Web 界面）

```bash
streamlit run app.py
```

打开浏览器后：上传 Excel → 点击“生成报表” → 下载生成的 `*_pivot.xlsx`。

## 3) 文件结构

- `process.py`：核心处理逻辑（无 UI）
- `app.py`：Streamlit 简体中文界面
- `cli.py`：命令行模式（可选）
- `requirements.txt`：依赖

## 4) （可选）命令行运行

如果你不想用 Streamlit，也可以直接命令行生成：

```bash
python cli.py -i "输入文件.xlsx" -o "输出文件_pivot.xlsx"
```

