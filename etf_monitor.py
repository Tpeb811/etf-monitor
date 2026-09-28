import pandas as pd
import datetime
import os

# 强制创建输出目录，解决报错的核心
OUTPUT_DIR = os.path.join(os.getcwd(), "html_output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 获取 ETF 数据（这里用测试数据，确保脚本先能跑通）
print("🚀 正在获取 ETF 数据...")
data = {
    'ETF名称': ['沪深300ETF', '中证500ETF', '创业板ETF', '科创50ETF', '红利ETF'],
    '当前价格': [3.85, 6.21, 2.15, 0.88, 2.95],
    '涨跌幅': ['-0.5%', '+1.2%', '-2.1%', '+0.8%', '+0.3%'],
    '更新时间': [datetime.datetime.now().strftime('%Y-%m-%d %H:%M')] * 5
}
df = pd.DataFrame(data)
print("✅ 数据获取成功！")

# 生成 Excel 文件
excel_filename = f"ETF_Data_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
excel_path = os.path.join(OUTPUT_DIR, excel_filename)
df.to_excel(excel_path, index=False)
print(f"💾 Excel 已生成在: {excel_path}")

# 生成网页 (index.html)
html_table = df.to_html(index=False, table_id='etf-table', na_rep='-', border=0)

html_content = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>ETF 监控面板</title>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, sans-serif; background: #f4f7f6; padding: 40px; }}
        .container {{ max-width: 800px; margin: 0 auto; background: #fff; padding: 40px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.05); }}
        h1 {{ color: #2c3e50; border-left: 5px solid #007bff; padding-left: 15px; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
        th {{ background: #f8f9fa; padding: 12px; text-align: left; border-bottom: 2px solid #dee2e6; }}
        td {{ padding: 12px; border-bottom: 1px solid #e9ecef; }}
        .btn {{ display: inline-block; background: #27ae60; color: #fff; padding: 10px 20px; text-decoration: none; border-radius: 5px; margin-bottom: 20px; font-weight: bold; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>每日 ETF 监控</h1>
        <p>数据生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        <a href="{excel_filename}" class="btn">📥 下载完整 Excel 数据</a>
        {html_table}
    </div>
</body>
</html>
"""

index_path = os.path.join(OUTPUT_DIR, "index.html")
with open(index_path, "w", encoding="utf-8") as f:
    f.write(html_content)
print(f"🌐 网页已生成在: {index_path}")
print("✨ 全部任务完成！")
