# -*- coding: utf-8 -*-
"""桌面 API 单元测试：不弹对话框，直接验证 template_options / export_xls / save_file。"""
import base64
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import desktop  # noqa: E402

TPL = r"C:\Users\bwyyyyy\Desktop\房产表计单据安装点导入模板.xls"
OUT = r"D:\test\form-filler\samples\api-test.xls"

with open(TPL, "rb") as f:
    raw = f.read()
b64 = base64.b64encode(raw).decode()

api = desktop.Api()
info = api.template_options({"template": b64})
print("options:", [(v["col"], len(v["values"])) for v in info.get("validations", [])])

rows = []
for i in range(3):
    rows.append([
        "海口市测试路%d号" % i, "否", "北区3-", "阳台", "是", "", "否",
        "", "栋", 1, "层", "01", "户",
    ])

api._ask_save = lambda name: OUT
result = api.export_xls({
    "template": b64,
    "sheetName": "Sheet0",
    "headerRow": 1,
    "floorCol": 10,
    "roomCol": 12,
    "rows": rows,
})
print("export:", result)
print("size:", os.path.getsize(OUT) if os.path.exists(OUT) else 0)

# save_file 走同一路径
api._ask_save = lambda name: OUT.replace("api-test.xls", "api-save.xlsx")
print("save:", api.save_file(base64.b64encode(b"hello").decode(), "x.xlsx"))
