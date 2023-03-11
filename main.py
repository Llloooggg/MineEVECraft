import time
import logging

import numpy as np
import pygetwindow as gw
import pyautogui
import cv2
import easyocr
import pandas as pd


win_name = "EVE - Nostrom Stone"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(message)s",
    datefmt="%d-%b-%y %H:%M:%S",
    handlers=[logging.StreamHandler()],
)

debug = True


def get_screenshot():
    eve_window = gw.getWindowsWithTitle(win_name)[0]

    was_minimized = False
    if eve_window.isMinimized:
        was_minimized = True

    eve_window.maximize()
    eve_window.activate()

    if debug:
        filepath = "images/0_screenshot.png"
    else:
        filepath = None

    time.sleep(0.5)
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

    return cv2.cvtColor(np.array(screenshot), cv2.COLOR_BGR2RGB)


def get_boxes(screenshot):
    reader = easyocr.Reader(["en"], gpu=True)
    results = reader.readtext(
        cv2.bitwise_not(cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)),
        width_ths=1,
    )

    results_frame = pd.DataFrame(
        [
            [
                int(result[0][0][0]),
                int(result[0][0][1]),
                int(result[0][1][0]),
                int(result[0][1][1]),
                int(result[0][2][0]),
                int(result[0][2][1]),
                int(result[0][3][0]),
                int(result[0][3][1]),
                result[1],
            ]
            for result in results
        ],
        columns=[
            "tl_x",
            "tl_y",
            "bl_x",
            "bl_y",
            "br_x",
            "br_y",
            "tr_x",
            "tr_y",
            "text",
        ],
    )

    results_frame = results_frame.loc[results_frame["text"].str.len() > 2]

    if debug:
        results_frame.to_excel("xlsx/1_boxes.xlsx", index=False)
        for result in results_frame.itertuples(index=False):
            tl = (result.tl_x, result.tl_y)
            br = (result.br_x, result.br_y)

            cv2.rectangle(screenshot, tl, br, (0, 255, 0), 1)
            cv2.putText(
                screenshot,
                result.text,
                (tl[0], tl[1] - 5),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.4,
                (0, 255, 0),
                1,
            )
        cv2.imwrite("images/1_highlited_screenshot.png", screenshot)

    logging.info("Боксы получены")

    return results_frame


screenshot = get_screenshot()
get_boxes(screenshot)
