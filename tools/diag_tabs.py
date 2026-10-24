# -*- coding: utf-8 -*-
"""诊断：列出 Outlook 功能区里的选项卡(TabItem)/模块入口，用于切到日历模块。

用法：python tools/diag_tabs.py
"""
import sys
import ctypes
import uiautomation as auto

auto.SetGlobalSearchTimeout(5)

user32 = ctypes.windll.user32
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


def collect(root, depth=0, max_depth=16, acc=None):
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


def main():
    hwnd = pid = None
    for h, p, c, t in enum_hwnd("rctrl_renwnd32"):
        hwnd, pid = h, p
        break
    if not hwnd:
        print("FAIL: outlook main window not found")
        return 1
    w = auto.ControlFromHandle(hwnd)
    items = collect(w)
    print("nodes=%d" % len(items))
    print("=== TabItem / 选项卡 ===")
    for d, c, n, t in items:
        if t in ("TabItemControl", "TabControl"):
            print("  depth=%d %-16s %r" % (d, t, n))
    print("=== Text / 自定义类名里含 '日历' 的节点 ===")
    for d, c, n, t in items:
        if n and ("\u65e5\u5386" in n or "Calendar" in n):
            print("  depth=%d %-16s %r" % (d, t, n))
    print("=== 顶层模块类名（ribbon 容器） ===")
    for d, c, n, t in items:
        if d <= 2 and n:
            print("  depth=%d %-16s %r" % (d, t, n))
    return 0


if __name__ == "__main__":
    sys.exit(main())
