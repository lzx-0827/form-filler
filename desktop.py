# -*- coding: utf-8 -*-
"""表计安装点填表器 · 桌面客户端（pywebview + WebView2）。

双击 exe 打开原生窗口；导出时调用 Excel 写模板，并弹出原生“另存为”。
"""
import base64
import datetime
import os
import pathlib
import sys

import webview

from excel_export import fill_template, template_options

# WebView2 的窗口遮挡计算会导致拖动窗口时假死，关闭该特性
os.environ.setdefault(
    "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS",
    "--disable-features=CalculateNativeWinOcclusion",
)


def resource_root():
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def default_name(ext):
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    return "房产表计安装点导入_%s.%s" % (stamp, ext)


class Api:
    def __init__(self):
        self.window = None

    def template_options(self, payload):
        try:
            data = base64.b64decode(payload["template"])
            return template_options(data)
        except Exception as exc:  # noqa: BLE001
            return {"error": str(exc)}

    def export_xls(self, payload):
        try:
            data = fill_template(
                base64.b64decode(payload["template"]),
                payload.get("sheetName") or "Sheet0",
                int(payload.get("headerRow") or 1),
                payload["rows"],
                int(payload["floorCol"]) if payload.get("floorCol") else None,
                int(payload["roomCol"]) if payload.get("roomCol") else None,
            )
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}

        path = self._ask_save(default_name("xls"))
        if not path:
            return {"ok": False, "canceled": True}
        try:
            with open(path, "wb") as f:
                f.write(data)
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}
        return {"ok": True, "path": path}

    def save_file(self, content, filename):
        path = self._ask_save(filename or default_name("xlsx"))
        if not path:
            return {"ok": False, "canceled": True}
        try:
            with open(path, "wb") as f:
                f.write(base64.b64decode(content))
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}
        return {"ok": True, "path": path}

    def _ask_save(self, filename):
        result = self.window.create_file_dialog(
            webview.FileDialog.SAVE,
            save_filename=filename,
            file_types=("Excel 文件 (*.xls;*.xlsx)", "所有文件 (*.*)"),
        )
        if not result:
            return None
        if isinstance(result, (list, tuple)):
            return result[0] if result else None
        return result


def main():
    api = Api()
    window = webview.create_window(
        "表计安装点填表器",
        pathlib.Path(resource_root(), "index.html").as_uri(),
        js_api=api,
        width=1280,
        height=880,
        min_size=(1080, 700),
        text_select=True,
        easy_drag=False,
    )
    api.window = window
    webview.start(debug=False)


if __name__ == "__main__":
    main()


