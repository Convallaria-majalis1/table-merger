# -*- coding: utf-8 -*-
"""
生成测试用的示例数据（故意做得很脏，用来验证清洗逻辑）
运行：python generate_samples.py
"""
from pathlib import Path

import pandas as pd

out_dir = Path(__file__).parent / "示例数据"
out_dir.mkdir(exist_ok=True)

# 1月：带空格、空行、重复行、长订单号
jan = pd.DataFrame(
    {
        "订单号": ["202601001", " 202601002 ", "202601003", "", "202601003", "  "],
        "客户名称": ["  兰州某某商贸  ", "张记餐饮", "王五 ", "", "王五", "空白行"],
        "金额": [1200.5, 800, 350.0, None, 350.0, None],
        "备注": ["首单", "", "  待确认", "", "  待确认", ""],
    }
)
jan.to_excel(out_dir / "1月订单.xlsx", index=False, sheet_name="订单")

# 2月：CSV，GB18030 编码，列结构不同（没有备注，多了业务员）
feb = pd.DataFrame(
    {
        "订单号": ["202602001", "202602002", "202602003"],
        "客户名称": ["李四", "赵六商贸", "某某科技"],
        "金额": [560, 2100, 90],
        "业务员": ["小张", "小张", "小李"],
    }
)
feb.to_csv(out_dir / "2月订单.csv", index=False, encoding="gb18030")

# 3月：一个文件两个 Sheet
mar_a = pd.DataFrame(
    {
        "订单号": ["202603001", "202603002"],
        "客户名称": ["甲公司", "乙公司 "],
        "金额": [300, 1500],
        "业务员": ["小张", "小李"],
    }
)
mar_b = pd.DataFrame(
    {
        "订单号": ["202603003", "202603004"],
        "客户名称": ["丙公司", "丁公司"],
        "金额": [220, 780],
    }
)
with pd.ExcelWriter(out_dir / "3月订单.xlsx", engine="openpyxl") as writer:
    mar_a.to_excel(writer, index=False, sheet_name="上半月")
    mar_b.to_excel(writer, index=False, sheet_name="下半月")

print(f"示例数据已生成：{out_dir}")
