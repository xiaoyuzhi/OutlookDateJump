# -*- coding: utf-8 -*-
"""验证三个下拉的滑动窗口行为（revision 47+）。

流程（全程无鼠标，UIA 模式调用）：
  1. UIA 触发「跳转日期」按钮，EnumWindows 等弹窗出现；
  2. 读每个下拉的当前值，动态计算期望窗口 [v-5 .. v+5]
     （月跨年环绕、日夹紧当月）；
  3. ExpandCollapse 展开，读取 ListItem 集合比对；
  4. SelectionItemPattern 选中最后一项（边界项），
     等 SelectionChanged 重建窗口后重新展开，验证新窗口；
  5. WindowPattern 关窗（picked=NULL，不触发跳转）。

用法：
  python verify_combos.py    # 需已安装 uiautomation 的 Python 3 环境
"""
import time
import ctypes
import sys
from ctypes import wintypes
import uiautomation as auto

auto.SetGlobalSearchTimeout(5)
TARGET = "\u8df3\u8f6c\u65e5\u671f"          # 跳转日期
DLG_TITLE = "\u8df3\u8f6c\u5230\u65e5\u671f"  # 跳转到日期
SPAN = 5

user32 = ctypes.windll.user32
EnumWindows = user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
GetClassName = user32.GetClassNameW
GetWindowText = user32.GetWindowTextW
IsWindowVisible = user32.IsWindowVisible

FAILURES = []


def enum_hwnd(cls_match=None):
    out = []

    def cb(hwnd, lparam):
        pid = wintypes.DWORD()
        GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        cls = ctypes.create_unicode_buffer(256)
        GetClassName(hwnd, cls, 256)
        if cls_match is None or cls.value == cls_match:
            if IsWindowVisible(hwnd):
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
            if wpid == pid and cls.startswith("HwndWrapper") and DLG_TITLE in txt:
                return hwnd
        time.sleep(0.25)
    return None


def wrap_month(m):
    return ((m - 1) % 12 + 12) % 12 + 1


def days_in_month(y, m):
    if m == 2:
        leap = (y % 4 == 0 and y % 100 != 0) or (y % 400 == 0)
        return 29 if leap else 28
    return 31 if m in (1, 3, 5, 7, 8, 10, 12) else 30


def check(label, actual, expect):
    ok = actual == expect
    print("%s %s: %s" % ("PASS" if ok else "FAIL", label, actual))
    if not ok:
        print("      expected: %s" % expect)
        FAILURES.append(label)
    return ok


def collect_listitems(root, max_depth=14):
    found = []

    def walk(r, depth):
        if depth > max_depth:
            return
        try:
            kids = r.GetChildren()
        except Exception:
            return
        for c in kids:
            try:
                if c.ControlTypeName == "ListItemControl":
                    found.append(c)
            except Exception:
                pass
            walk(c, depth + 1)

    walk(root, 0)
    return found


def find_combo(dlg, auto_id):
    combo = []

    def walk(root, depth=0):
        if combo or depth > 12:
            return
        try:
            kids = root.GetChildren()
        except Exception:
            return
        for c in kids:
            try:
                if (c.AutomationId or "") == auto_id and c.ControlTypeName == "ComboBoxControl":
                    combo.append(c)
                    return
            except Exception:
                pass
            walk(c, depth + 1)

    walk(dlg)
    return combo[0] if combo else None


def combo_value(combo):
    try:
        return combo.GetValuePattern().CurrentValue
    except Exception:
        return combo.Name


def read_items(dlg, combo, pid):
    """展开 combo，收集 ListItem 名称（combo 子树 -> 弹窗子树 -> 进程窗口 三级兜底）。"""
    try:
        combo.GetExpandCollapsePattern().Expand()
    except Exception as e:
        return None, "expand failed: %s" % e
    time.sleep(0.5)
    items = collect_listitems(combo)
    if not items:
        items = collect_listitems(dlg)
    if not items:
        for h, wpid, cls, txt in enum_hwnd():
            if wpid == pid and cls.startswith("HwndWrapper"):
                items = collect_listitems(auto.ControlFromHandle(h))
                if items:
                    break
    try:
        combo.GetExpandCollapsePattern().Collapse()
    except Exception:
        pass
    return [c.Name for c in items], None


def select_last(dlg, combo, pid):
    """展开并选中最后一项（边界项），触发窗口重建。"""
    try:
        combo.GetExpandCollapsePattern().Expand()
    except Exception:
        return False
    time.sleep(0.5)
    items = collect_listitems(combo)
    if not items:
        items = collect_listitems(dlg)
    if not items:
        return False
    try:
        items[-1].GetSelectionItemPattern().Select()
        return True
    except Exception:
        return False
    finally:
        time.sleep(0.6)


def main():
    wins = enum_hwnd("rctrl_renwnd32")
    if not wins:
        print("Outlook main window not found")
        return 1
    hwnd, pid, _, title = wins[0]
    print("outlook: pid=%d %r" % (pid, title))

    w = auto.ControlFromHandle(hwnd)
    btn = []

    def walkb(root, depth=0):
        if btn or depth > 13:
            return
        try:
            kids = root.GetChildren()
        except Exception:
            return
        for c in kids:
            try:
                if (c.Name or "") == TARGET and c.ControlTypeName == "ButtonControl" and not btn:
                    btn.append(c)
                    return
            except Exception:
                pass
            walkb(c, depth + 1)

    walkb(w)
    if not btn:
        print("ribbon button not found")
        return 1
    btn[0].GetPattern(auto.PatternId.InvokePattern).Invoke()
    print("invoked")

    dh = find_dialog(pid)
    if dh is None:
        print("dialog not found")
        return 1
    time.sleep(0.8)
    dlg = auto.ControlFromHandle(dh)
    r = dlg.BoundingRectangle
    print("dialog: %dx%d" % (r.width(), r.height()))

    # ---- 年 ----
    cmbY = find_combo(dlg, "CmbYear")
    if cmbY is None:
        print("CmbYear not found")
        return 1
    y = int(combo_value(cmbY))
    items, err = read_items(dlg, cmbY, pid)
    if err:
        print("year:", err)
        return 1
    check("year window", items, [str(i) for i in range(y - SPAN, y + SPAN + 1)])
    if select_last(dlg, cmbY, pid):
        items, _ = read_items(dlg, cmbY, pid)
        check("year re-anchor", items, [str(i) for i in range(y, y + 2 * SPAN + 1)])

    # ---- 月 ----
    cmbM = find_combo(dlg, "CmbMonth")
    m = int(combo_value(cmbM))
    items, err = read_items(dlg, cmbM, pid)
    if err:
        print("month:", err)
        return 1
    check("month window", items, [str(wrap_month(m + i)) for i in range(-SPAN, SPAN + 1)])
    if select_last(dlg, cmbM, pid):
        new_m = wrap_month(m + SPAN)
        items, _ = read_items(dlg, cmbM, pid)
        check("month re-anchor", items, [str(wrap_month(new_m + i)) for i in range(-SPAN, SPAN + 1)])

    # ---- 日 ----
    cmbD = find_combo(dlg, "CmbDay")
    y2 = int(combo_value(find_combo(dlg, "CmbYear")))
    m2 = int(combo_value(find_combo(dlg, "CmbMonth")))
    d = int(combo_value(cmbD))
    dim = days_in_month(y2, m2)
    items, err = read_items(dlg, cmbD, pid)
    if err:
        print("day:", err)
        return 1
    lo = max(1, d - SPAN)
    hi = min(dim, d + SPAN)
    check("day window", items, [str(i) for i in range(lo, hi + 1)])
    if select_last(dlg, cmbD, pid):
        d2 = hi  # 选中的边界项成为新锚点
        lo2 = max(1, d2 - SPAN)
        hi2 = min(dim, d2 + SPAN)
        items, _ = read_items(dlg, cmbD, pid)
        check("day re-anchor", items, [str(i) for i in range(lo2, hi2 + 1)])

    # ---- 收尾 ----
    try:
        dlg.GetWindowPattern().Close()
        print("closed (picked=NULL, no jump)")
    except Exception as e:
        print("close failed:", e)

    print("RESULT:", "ALL PASS" if not FAILURES else "FAILED: %s" % FAILURES)
    return 0 if not FAILURES else 2


if __name__ == "__main__":
    sys.exit(main())
