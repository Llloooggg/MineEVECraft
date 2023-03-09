import pygetwindow as gw
import pyautogui


def get_screenshot(win_name):
    eve_window = gw.getWindowsWithTitle(win_name)[0]

    if not eve_window.isMaximized:
        was_minimized = True
        eve_window.restore()
        eve_window = gw.getWindowsWithTitle(win_name)[0]

    screenshot = pyautogui.screenshot(
        "screenshot.png",
        region=(
            eve_window.box.left + 10,
            eve_window.box.top,
            eve_window.box.width - 20,
            eve_window.box.height - 10,
        ),
    )

    if was_minimized:
        eve_window.minimize()

    return screenshot


get_screenshot("EVE - Nostrom Stone")
