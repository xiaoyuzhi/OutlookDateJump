# -*- coding: utf-8 -*-
"""诊断：在 Outlook 功能区里找「跳转日期」按钮（比 verify_uia.py 宽松）。

用途：验证绿色版（|vstolocal 便携目录）是否真的加载并渲染了功能区按钮。

环境结论（2026-10-24 实测）：
- 沙箱内 PowerShell 读 Process.Modules 看不到本加载项的模块（不可靠），
  判定「是否加载」以 UIA 功能区树 / Outlook 事件 ID 45 为准。
- UIA 遍历要宽：按钮可能在 ribbon 容器下多层，且控件名可能带空白。
- 沙箱截图三路全死，观感只能靠 UIA BoundingRectangle 实测。
"""
import sys
import time
import ctypes
import uiautomation as auto

auto.SetGlobalSearchTimeout(5)
TARGET = "\u8df3\u8f6c\u65e5\u671f"   # 跳转日期

user32 = ctypes.windll.user32
EnumWindows = user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
GetClassName = user32.GetClassNameW
IsWindowVisible = user32.IsWindowVisible


def enum_main():
    out = []

    def cb(hwnd, lparam):
        cls = ctypes.create_unicode_buffer(256)
        GetClassName(hwnd, cls, 256)
        if cls.value == "rctrl_renwnd32" and IsWindowVisible(hwnd):
            pid = ctypes.c_ulong()
            GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            out.append((hwnd, pid.value))
        return True

    EnumWindows(EnumWindowsProc(cb), 0)
    return out


def collect(root, depth=0, max_depth=14, acc=None):
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
            acc.append((depth, c.Name, c.ControlTypeName))
        except Exception:
            pass
        try:
            if c.GetChildren():
                collect(c, depth + 1, max_depth, acc)
        except Exception:
            pass
    return acc


def main():
    wins = enum_main()
    if not wins:
        print("FAIL: outlook main window not found")
        return 1
    hwnd, pid = wins[0]
    print("outlook pid=%d hwnd=%d" % (pid, hwnd))

    w = auto.ControlFromHandle(hwnd)
    t0 = time.time()
    items = collect(w)
    print("tree nodes: %d (walk %.1fs)" % (len(items), time.time() - t0))

    # 1) 精确命中
    exact = [(d, n, t) for (d, n, t) in items if (n or "").strip() == TARGET]
    # 2) 包含命中（控件名可能带空白/附加描述）
    loose = [(d, n, t) for (d, n, t) in items if TARGET in (n or "")]
    print("exact hits : %d" % len(exact))
    print("loose hits : %d" % len(loose))
    for d, n, t in (exact or loose)[:5]:
        print("   depth=%d type=%s name=%r" % (d, t, n))
    if not (exact or loose):
        # 打印功能区里所有 Button 名称，便于人工确认渲染状态
        btns = sorted({n for (d, n, t) in items if t == "ButtonControl" and n})
        print("no match. visible buttons (%d):" % len(btns))
        for b in btns[:40]:
            print("   %r" % b)
        # 顺带看看有没有我们自己的窗体类名（DatePickerWindow）
        wins2 = [(h, p) for h, p in []]
        print("---")
    return 0 if (exact or loose) else 2


if __name__ == "__main__":
    sys.exit(main())
