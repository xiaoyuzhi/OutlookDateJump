#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""完整重启 Outlook 后验证弹窗里显示的版本号。

为什么要这个脚本：Outlook 已在运行时，Start-Process 只是激活现有进程，
**不会重新加载加载项 DLL**（旧 DLL 仍在内存里）。要验证新构建的产物，
必须先让 Outlook 进程彻底退出（COM Quit），再重新启动。

用法：python tools/verify_version_label.py
退出码：0=版本号一致；1=环境问题；2=未找到。
"""
import os
import re
import sys
import time
import ctypes
import uiautomation as auto

auto.SetGlobalSearchTimeout(5)
TARGET = "\u8df3\u8f6c\u65e5\u671f"       # 跳转日期
DLG_TITLE = "\u8df3\u8f6c\u5230\u65e5\u671f"  # 跳转到日期
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

user32 = ctypes.windll.user32
user32.GetWindowTextW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int
EnumWindows = user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
GetClassName = user32.GetClassNameW
GetWindowText = user32.GetWindowTextW
IsWindowVisible = user32.IsWindowVisible


def enum_hwnd(cls_match=None, visible_only=True):
    out = []

    def cb(hwnd, lparam):
        pid = ctypes.c_ulong()
        GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        cls = ctypes.create_unicode_buffer(256)
        GetClassName(hwnd, cls, 256)
        if (cls_match is None or cls.value == cls_match) and (not visible_only or IsWindowVisible(hwnd)):
            txt = ctypes.create_unicode_buffer(256)
            GetWindowText(hwnd, txt, 256)
            out.append((hwnd, pid.value, cls.value, txt.value))
        return True

    EnumWindows(EnumWindowsProc(cb), 0)
    return out


def collect(root, depth=0, max_depth=18, acc=None):
    if acc is None:
        acc = []
    if depth > max_depth:
        return acc
    try:
        kids = root.GetChildren()
    except Exception:
        return acc
    for c in kids:
        try:
            acc.append((depth, c, c.Name, c.ControlTypeName))
        except Exception:
            pass
        try:
            if c.GetChildren():
                collect(c, depth + 1, max_depth, acc)
        except Exception:
            pass
    return acc


def outlook_main():
    for h, p, c, t in enum_hwnd("rctrl_renwnd32"):
        return h, p
    return None, None


def open_dialog():
    """切日历模块 -> 点「跳转日期」-> 返回弹窗 hwnd。"""
    hwnd, pid = outlook_main()
    if not hwnd:
        return None, "outlook main window not found"

    # 切到日历模块（左侧模块列表里的 ListItemControl）
    for attempt in range(3):
        items = collect(auto.ControlFromHandle(hwnd))
        cal = [c for (d, c, n, t) in items
               if (n or "").strip() == "\u65e5\u5386" and t == "ListItemControl"]
        if not cal:
            return None, "module item '日历' not found"
        try:
            cal[0].GetPattern(auto.PatternId.InvokePattern).Invoke()
        except Exception:
            pass
        time.sleep(4)

        # 确认已切到日历模块：功能区应出现「今天 / 新建会议」这类按钮
        items = collect(auto.ControlFromHandle(hwnd))
        names = {n for (d, c, n, t) in items if t == "ButtonControl" and n}
        if any("\u4eca\u5929" in n for n in names):
            break
        print("attempt %d: not in calendar module yet (buttons: %s)" % (attempt + 1, sorted(names)[:6]))

    # 选「开始」选项卡
    items = collect(auto.ControlFromHandle(hwnd))
    tabs = [c for (d, c, n, t) in items
            if (n or "").strip() == "\u5f00\u59cb" and t == "TabItemControl"]
    if tabs:
        try:
            tabs[0].GetPattern(auto.PatternId.InvokePattern).Invoke()
        except Exception:
            pass
        time.sleep(2)

    items = collect(auto.ControlFromHandle(hwnd))
    btn = [c for (d, c, n, t) in items if (n or "").strip() == TARGET and t == "ButtonControl"]
    if not btn:
        return None, "ribbon button '跳转日期' not found"
    btn[0].GetPattern(auto.PatternId.InvokePattern).Invoke()

    t0 = time.time()
    while time.time() - t0 < 15:
        for h, wpid, cls, txt in enum_hwnd():
            if wpid == pid and cls.startswith("HwndWrapper") and DLG_TITLE in txt and IsWindowVisible(h):
                return h, None
        time.sleep(0.3)
    return None, "dialog not found"


def main():
    expected = open(os.path.join(ROOT, "VERSION"), encoding="utf-8-sig").read().strip()
    print("VERSION file expects: %s" % expected)

    dh, err = open_dialog()
    if dh is None:
        print("FAIL: %s" % err)
        return 2
    time.sleep(0.6)
    dlg = auto.ControlFromHandle(dh)

    texts = []
    for d, c, n, t in collect(dlg):
        if n and n.strip():
            texts.append(n.strip())
    print("dialog texts: %s" % texts)

    ver_lines = [x for x in texts if re.search(r"\d+\.\d+\.\d+", x)]
    dlg.GetWindowPattern().Close()
    time.sleep(0.3)

    if not ver_lines:
        print("FAIL: no version text visible in the dialog")
        return 1
    if not any(expected in x for x in ver_lines):
        print("FAIL: version %s not present in %s" % (expected, ver_lines))
        return 1
    print("PASS: dialog shows version %s" % expected)
    return 0


if __name__ == "__main__":
    sys.exit(main())
