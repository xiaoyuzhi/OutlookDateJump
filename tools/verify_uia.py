# -*- coding: utf-8 -*-
"""OutlookDateJump 无头验证脚本（沙箱安全版）。

用法（需已安装 uiautomation 的 Python 3 环境，系统 python 通常没有）：
  python verify_uia.py

流程：UIA 找 Outlook 主窗口 -> 树遍历找「跳转日期」按钮 ->
      InvokePattern 触发 -> EnumWindows 找弹窗（UIA 根枚举看不到模态拥有窗口！）-
      测量窗口尺寸 -> WindowPattern 关窗（picked=NULL，不触发跳转）。

环境结论（2026-10-24 实测）：
- uiautomation 只装在受管 venv，必须用绝对路径调用。
- 沙箱禁 Add-Type（PowerShell 运行时编译），UIA 走 Python uiautomation。
- 屏幕截图三路全死：PIL ImageGrab / BitBlt（screen grab failed）、
  PrintWindow（WPF+DirectComposition 返回全黑或全白）、
  Alt+PrintScreen 剪贴板（合成按键被丢/剪贴板为空）。
  观感验证只能靠：UIA BoundingRectangle 实测尺寸 + odjdbg.log fit 行。
- UIA GetRootControl().GetChildren() 看不到本加载项的模态弹窗
  （Owner=Outlook 的 HwndWrapper 窗口），必须 EnumWindows 或 ControlFromHandle。
"""
import time
import ctypes
import sys
from ctypes import wintypes
import uiautomation as auto

auto.SetGlobalSearchTimeout(5)
TARGET = "\u8df3\u8f6c\u65e5\u671f"          # 跳转日期
DLG_TITLE = "\u8df3\u8f6c\u5230\u65e5\u671f"  # 跳转到日期

user32 = ctypes.windll.user32
EnumWindows = user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
GetClassName = user32.GetClassNameW
GetWindowText = user32.GetWindowTextW
IsWindowVisible = user32.IsWindowVisible


def enum_hwnd(cls_match=None, visible_only=True):
    out = []

    def cb(hwnd, lparam):
        pid = wintypes.DWORD()
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


def find_dialog(pid, timeout_s=12):
    t0 = time.time()
    while time.time() - t0 < timeout_s:
        for hwnd, wpid, cls, txt in enum_hwnd():
            if wpid == pid and cls.startswith("HwndWrapper") and DLG_TITLE in txt and IsWindowVisible(hwnd):
                return hwnd
        time.sleep(0.25)
    return None


def main():
    wins = [x for x in enum_hwnd("rctrl_renwnd32")]
    if not wins:
        print("Outlook main window not found (is Outlook running?)")
        return 1
    hwnd, pid, _, title = wins[0]
    print("outlook: pid=%d %r" % (pid, title))

    w = auto.ControlFromHandle(hwnd)

    btn = []

    def walk(root, max_depth=13):
        level = root.GetChildren()
        depth = 0
        while level and depth < max_depth:
            nxt = []
            for c in level:
                try:
                    if (c.Name or "") == TARGET and c.ControlTypeName == "ButtonControl" and not btn:
                        btn.append(c)
                except Exception:
                    pass
                try:
                    kids = c.GetChildren()
                    if kids:
                        nxt.extend(kids)
                except Exception:
                    pass
            level = nxt
            depth += 1

    walk(w)
    if not btn:
        print("button %r not found in ribbon tree" % TARGET)
        return 1
    b = btn[0]
    r = b.BoundingRectangle
    grp = b.GetParentControl()
    print("button: %r rect=(%d,%d)-(%d,%d) group=%r" % (b.Name, r.left, r.top, r.right, r.bottom, grp.Name))

    try:
        b.GetPattern(auto.PatternId.InvokePattern).Invoke()
        print("invoked")
    except Exception as e:
        print("invoke failed: %s" % e)
        return 1

    dh = find_dialog(pid)
    if dh is None:
        print("dialog not found")
        return 1
    time.sleep(0.3)
    c = auto.ControlFromHandle(dh)
    dr = c.BoundingRectangle
    print("dialog: %dx%d at (%d,%d)" % (dr.width(), dr.height(), dr.left, dr.top))

    try:
        c.GetWindowPattern().Close()
        print("closed (picked=NULL, no jump)")
    except Exception as e:
        print("close failed: %s" % e)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
