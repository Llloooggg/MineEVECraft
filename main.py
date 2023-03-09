import logging
import numpy as np
import pygetwindow as gw
import pyautogui
from pytesseract import pytesseract
from pytesseract import Output
import cv2

win_name = "EVE - Nostrom Stone"
pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(message)s",
    datefmt="%d-%b-%y %H:%M:%S",
    handlers=[logging.StreamHandler()],
)


def get_screenshot(save_result=False):
    eve_window = gw.getWindowsWithTitle(win_name)[0]

    was_minimized = False
    if eve_window.isMinimized:
        was_minimized = True

    eve_window.maximize()
    eve_window.activate()

    if save_result:
        filepath = "images/screenshot.png"
    else:
        filepath = None

    screenshot = pyautogui.screenshot(
        filepath,
        region=(
            eve_window.box.left + 10,
            eve_window.box.top,
            eve_window.box.width - 20,
            eve_window.box.height - 10,
        ),
    )

    if was_minimized:
        eve_window.minimize()

    logging.info("Скриншот получен")

    return screenshot


def get_boxed(screenshot, save_result=False):
    # img = cv2.imread("test.png")  # from file
    img = cv2.cvtColor(np.array(screenshot), cv2.COLOR_BGR2RGB)

    raw_df = pytesseract.image_to_data(
        img,
        lang="eng",
        output_type=Output.DATAFRAME,
        config="--psm 4 -c preserve_interword_spaces=1",
    )
    filtered_df = raw_df.loc[
        (raw_df["text"].str.isalnum()) & (raw_df["text"].str.len() > 3)
    ][["left", "top", "width", "height", "text"]]

    if save_result:
        for r in filtered_df.itertuples():
            (x, y, w, h) = (
                r.left,
                r.top,
                r.width,
                r.height,
            )
            cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)
            cv2.putText(
                img,
                f"'{r.text}': {r.left}.{r.top} {r.width}.{r.height}",
                (x, y - 10),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.4,
                (0, 255, 0),
                1,
            )

        cv2.imwrite("images/highlighted_screenshot.png", img)

    logging.info("Боксы получены")

    return filtered_df


screenshot = get_screenshot(True)
get_boxed(screenshot, True)
