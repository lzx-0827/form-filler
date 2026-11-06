# -*- coding: utf-8 -*-
"""Excel COM：把行写进 .xls 模板（保留样式与下拉），并读取模板的下拉选项。"""
import os
import shutil
import tempfile

import pythoncom
import win32com.client as win32

XL_VALIDATE_LIST = 3
XL_VALID_ALERT_STOP = 1
XL_BETWEEN = 1
XL_DOWN = -4121
XL_FORMAT_FROM_LEFT_OR_ABOVE = 0
XL_EXCEL8 = 56


def fill_template(template_bytes, sheet_name, header_row, rows, floor_col=None, room_col=None):
    """打开模板 → 调整行数 → 写入 → 重设下拉 → 另存为 .xls，返回字节。"""
    workdir = tempfile.mkdtemp(prefix="meter-form-")
    src = os.path.join(workdir, "template.xls")
    out = os.path.join(workdir, "output.xls")
    with open(src, "wb") as f:
        f.write(template_bytes)

    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    try:
        wb = excel.Workbooks.Open(src)
        ws = wb.Worksheets(sheet_name)
        last_col = len(rows[0]) if rows else 13
        target_last = header_row + len(rows)

        # 记录每个数据列的下拉设置（取第一个数据行）
        validations = {}
        for col in range(1, last_col + 1):
            try:
                v = ws.Cells(header_row + 1, col).Validation
                if v.Type == XL_VALIDATE_LIST:
                    validations[col] = (v.Formula1, v.InCellDropdown, v.IgnoreBlank)
            except Exception:
                pass

        used = ws.UsedRange
        current_last = used.Row + used.Rows.Count - 1

        # 行数不足：在末尾插入，格式沿用上一行
        if current_last < target_last:
            need = target_last - current_last
            ws.Rows(f"{current_last + 1}:{current_last + need}").Insert(
                Shift=XL_DOWN, CopyOrigin=XL_FORMAT_FROM_LEFT_OR_ABOVE
            )
        elif current_last > target_last:
            ws.Rows(f"{target_last + 1}:{current_last}").Delete()

        # 严格保留模板格式：字体、数字格式都不动
        write_rows = []
        for row in rows:
            out_row = ["" if v is None else str(v) for v in row]
            if floor_col:
                try:
                    out_row[floor_col - 1] = int(float(row[floor_col - 1]))
                except (TypeError, ValueError):
                    pass
            write_rows.append(out_row)

        data_rng = ws.Range(ws.Cells(header_row + 1, 1), ws.Cells(target_last, last_col))
        data_rng.Value = write_rows

        # 房号 01–09 保持文本，10 以上仍是数字
        if room_col:
            room_vals = []
            for row in rows:
                value = "" if row[room_col - 1] is None else str(row[room_col - 1])
                if value.startswith("0"):
                    room_vals.append(["'" + value])
                elif value.isdigit():
                    room_vals.append([int(value)])
                else:
                    room_vals.append([value])
            ws.Range(ws.Cells(header_row + 1, room_col), ws.Cells(target_last, room_col)).Value = room_vals

        # 重新铺一遍下拉，覆盖整个数据区
        for col, (formula, dropdown, ignore_blank) in validations.items():
            cell_range = ws.Range(ws.Cells(header_row + 1, col), ws.Cells(target_last, col))
            try:
                cell_range.Validation.Delete()
            except Exception:
                pass
            cell_range.Validation.Add(
                XL_VALIDATE_LIST, XL_VALID_ALERT_STOP, XL_BETWEEN, formula
            )
            cell_range.Validation.InCellDropdown = bool(dropdown)
            cell_range.Validation.IgnoreBlank = bool(ignore_blank)

        wb.SaveAs(out, FileFormat=XL_EXCEL8)
        wb.Close(SaveChanges=False)
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    with open(out, "rb") as f:
        data = f.read()
    shutil.rmtree(workdir, ignore_errors=True)
    return data


def _flatten(value):
    if value is None:
        return []
    if isinstance(value, tuple):
        out = []
        for item in value:
            out.extend(_flatten(item))
        return out
    text = str(value).strip()
    return [text] if text else []


def resolve_validation(ws, formula):
    """把模板里的下拉来源解析成值列表（支持逗号列表、区域、定义名称）。"""
    if not formula:
        return []
    formula = formula.strip()
    if formula.startswith("="):
        try:
            rng = ws.Range(formula[1:])
            return _flatten(rng.Value)
        except Exception:
            return []
    return [item.strip() for item in formula.split(",") if item.strip()]


def template_options(template_bytes):
    """读取模板每一列的下拉选项，返回 {sheetName, validations:[{col, values}]}。"""
    workdir = tempfile.mkdtemp(prefix="meter-options-")
    src = os.path.join(workdir, "template.xls")
    with open(src, "wb") as f:
        f.write(template_bytes)

    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(src)
        ws = wb.Worksheets(1)
        last_col = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
        validations = []
        for col in range(1, last_col + 1):
            try:
                v = ws.Cells(2, col).Validation
                if v.Type != XL_VALIDATE_LIST:
                    continue
                values = resolve_validation(ws, v.Formula1)
                if values:
                    validations.append({"col": col, "values": values})
            except Exception:
                continue
        result = {"sheetName": ws.Name, "validations": validations}
        wb.Close(SaveChanges=False)
        return result
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
        shutil.rmtree(workdir, ignore_errors=True)
