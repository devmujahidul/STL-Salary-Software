"""Read the text of the modal message box that blocks the Excel automation."""
import win32gui

results = []


def cb(hwnd, _):
    if not win32gui.IsWindowVisible(hwnd):
        return
    text = win32gui.GetWindowText(hwnd)
    cls = win32gui.GetClassName(hwnd)
    if text:
        child_texts = []
        child = win32gui.FindWindowEx(hwnd, 0, None, None)
        while child:
            t = win32gui.GetWindowText(child)
            if t:
                child_texts.append(t)
            child = win32gui.FindWindowEx(hwnd, child, None, None)
        results.append((hwnd, cls, text, child_texts))


win32gui.EnumWindows(cb, None)
for hwnd, cls, text, children in results:
    if "Sub Section" in text or "Excel" in text or "Microsoft" in text:
        print(hwnd, cls, repr(text))
        for c in children:
            print("    ", repr(c))
