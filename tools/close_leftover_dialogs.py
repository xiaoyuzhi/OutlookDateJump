# -*- coding: utf-8 -*-
"""关闭 Outlook 残留的模态对话框（测试被残留弹窗挡住时先跑这个）。"""
import ctypes, sys, time
import uiautomation as auto

auto.SetGlobalSearchTimeout(3)
u = ctypes.windll.user32
CB = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)


def find_dialogs():
    out = []

    def cb(h, l):
        cls = ctypes.create_unicode_buffer(256); u.GetClassNameW(h, cls, 256)
        txt = ctypes.create_unicode_buffer(256); u.GetWindowTextW(h, txt, 256)
        if u.IsWindowVisible(h) and cls.value in ("NUIDialog", "#32770"):
            out.append(h)
        return True

    u.EnumWindows(CB(cb), 0)
    return out


def main():
    closed = 0
    for _ in range(5):
        ds = find_dialogs()
        if not ds:
            break
        for h in ds:
            try:
                c = auto.ControlFromHandle(h)
                c.GetWindowPattern().Close()
                closed += 1
            except Exception as e:
                print("close failed: %s" % e)
            time.sleep(0.8)
    print("closed %d leftover dialog(s)" % closed)
    return 0


if __name__ == "__main__":
    sys.exit(main())
