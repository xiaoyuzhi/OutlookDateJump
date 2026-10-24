# -*- coding: utf-8 -*-
"""日历模块功能区端到端验证（沙箱版）。

流程：UIA 找左侧模块列表里的「日历」-> 点击（切模块）-> 选「开始」选项卡 ->
      找「跳转日期」按钮 -> Invoke -> EnumWindows 找弹窗 -> 量尺寸 -> 关窗。

用法（仓库根目录）：
    python tools/verify_calendar_ribbon.py

退出码：0=全部通过；2=按钮未找到；1=环境/交互失败。

环境结论（2026-10-24 实测）：
- Outlook 启动后默认停在邮件模块，必须先点左侧模块列表的「日历」（ListItemControl）
  才能看到本加载项的按钮；COM 的 ActiveExplorer.CurrentView 赋值在沙箱里无效。
- 功能区选项卡是 TabItemControl：「开始」/「发送/接收」/「文件夹」/「视图」/「开发工具」/「帮助」。
- 沙箱内 PowerShell 读 Process.Modules 看不到本加载项模块（不可靠），
  判定「是否加载」以本脚本能否找到按钮 + Outlook 事件 ID 45 为准。
- 截图三路全死，观感只能靠 UIA BoundingRectangle 实测尺寸。
"""
import sys
import time
import ctypes
import uiautomation as auto

auto.SetGlobalSearchTimeout(5)
TARGET = "\u8df3\u8f6c\u65e5\u671f"      # 跳转日期
DLG_TITLE = "\u8df3\u8f6c\u5230\u65e5\u671f"  # 跳转到日期
CALENDAR = "\u65e5\u5386"              # 日历
START_TAB = "\u5f00\u59cb"             # 开始

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


def snapshot(hwnd):
    return collect(auto.ControlFromHandle(hwnd))


def find_outlook():
    for h, p, c, t in enum_hwnd("rctrl_renwnd32"):
        return h, p, t
    return None, None, None


def click_control(ctrl):
    """优先 InvokePattern，回退 ClickablePoint + 鼠标点击。"""
    try:
        ctrl.GetPattern(auto.PatternId.InvokePattern).Invoke()
        return "invoke"
    except Exception:
        pass
    try:
        ctrl.GetPattern(auto.PatternId.SelectionItemPattern).Select()
        return "select"
    except Exception:
        pass
    try:
        x, y = ctrl.ClickablePoint
        ctrl.GetPattern(auto.PatternId.LegacyIAccessiblePattern).DoDefaultAction()
        return "legacy"
    except Exception as e:
        raise RuntimeError("cannot click control: %s" % e)


def switch_to_calendar(hwnd):
    """点左侧模块列表的「日历」，再点功能区「开始」选项卡。"""
    items = snapshot(hwnd)
    cal = [c for (d, c, n, t) in items if (n or "").strip() == CALENDAR and t == "ListItemControl"]
    if not cal:
        print("FAIL: module list item '日历' not found")
        return False
    print("clicking module '日历' ...")
    click_control(cal[0])
    time.sleep(4)

    items = snapshot(hwnd)
    tabs = [c for (d, c, n, t) in items if (n or "").strip() == START_TAB and t == "TabItemControl"]
    if tabs:
        print("clicking tab '开始' ...")
        try:
            click_control(tabs[0])
        except Exception as e:
            print("tab click failed (may already be active): %s" % e)
        time.sleep(2)
    return True


def find_button(hwnd):
    items = snapshot(hwnd)
    for d, c, n, t in items:
        if (n or "").strip() == TARGET and t == "ButtonControl":
            print("FOUND button %r (depth=%d)" % (n, d))
            return c
    return None


def find_dialog(pid, timeout_s=15):
    t0 = time.time()
    while time.time() - t0 < timeout_s:
        for h, wpid, cls, txt in enum_hwnd():
            if wpid == pid and cls.startswith("HwndWrapper") and DLG_TITLE in txt and IsWindowVisible(h):
                return h
        time.sleep(0.3)
    return None


def main():
    hwnd, pid, title = find_outlook()
    if not hwnd:
        print("FAIL: outlook main window not found")
        return 1
    print("outlook pid=%d %r" % (pid, title))

    if not switch_to_calendar(hwnd):
        return 1

    btn = find_button(hwnd)
    if not btn:
        print("button %r NOT found in calendar ribbon" % TARGET)
        names = sorted({n for (_d, _c, n, t) in snapshot(hwnd) if t == "ButtonControl" and n})
        print("buttons (%d): %s" % (len(names), names[:30]))
        return 2

    r = btn.BoundingRectangle
    print("button rect=(%d,%d)-(%d,%d)" % (r.left, r.top, r.right, r.bottom))
    try:
        btn.GetPattern(auto.PatternId.InvokePattern).Invoke()
        print("invoked")
    except Exception as e:
        print("invoke failed: %s" % e)
        return 1

    dh = find_dialog(pid)
    if dh is None:
        print("dialog not found")
        return 1
    time.sleep(0.4)
    c = auto.ControlFromHandle(dh)
    dr = c.BoundingRectangle
    print("dialog: %dx%d at (%d,%d)" % (dr.width(), dr.height(), dr.left, dr.top))
    try:
        c.GetWindowPattern().Close()
        print("closed (picked=NULL, no jump)")
    except Exception as e:
        print("close failed: %s" % e)
        return 1
    print("PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
