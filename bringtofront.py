import win32gui


def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))


def bring_window_to_front(window_title):
    results = []
    top_windows = []
    win32gui.EnumWindows(windowEnumerationHandler, top_windows)
    for i in top_windows:
        # print(i)
        print(i[1].lower())
        if window_title.lower() in i[1].lower():
            print("found window")
            win32gui.ShowWindow(i[0], 5)
            win32gui.SetForegroundWindow(i[0])
            break


bring_window_to_front('Document1 - Word')
