import time
import logging

import numpy as np
import pygetwindow as gw
import pyautogui
import cv2
import easyocr
import pandas as pd

debug = True

win_name = "EVE - Nostrom Stone"

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(message)s",
    datefmt="%d-%b-%y %H:%M:%S",
    handlers=[logging.StreamHandler()],
)

logging.info("Бот: запущен")

reader = easyocr.Reader(["en"], gpu=True)
logging.info("Бот: модели загружены")


def get_screenshot():
    eve_window = gw.getWindowsWithTitle(win_name)[0]

    was_minimized = False
    if eve_window.isMinimized:
        was_minimized = True

    eve_window.maximize()
    eve_window.activate()

    if debug:
        filepath = "images/screenshot.png"
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
    logging.info("Скриншот: получен")
    if debug:
        logging.debug("Скриншот: сохранен")

    if was_minimized:
        eve_window.minimize()

    return cv2.cvtColor(np.array(screenshot), cv2.COLOR_BGR2RGB)


def get_boxes(screenshot):
    results = reader.readtext(
        cv2.bitwise_not(cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)),
        low_text=0.4,
        width_ths=1.5,
    )
    logging.debug("Боксы: получены")

    results_frame = pd.DataFrame(
        [
            [
                int(result[0][0][0]),  # tl_x
                int(result[0][0][1]),  # tl_y
                int(result[0][1][0]),  # tr_x
                int(result[0][1][1]),  # tr_y
                int(result[0][2][0]),  # br_x
                int(result[0][2][1]),  # br_y
                int(result[0][3][0]),  # bl_x
                int(result[0][3][1]),  # bl_y
                int((result[0][0][0] + result[0][1][0]) / 2),  # cent_x
                int((result[0][0][1] + result[0][3][1]) / 2),  # cent_y
                result[1].lower(),  # text
            ]
            for result in results
        ],
        columns=[
            "tl_x",
            "tl_y",
            "tr_x",
            "tr_y",
            "br_x",
            "br_y",
            "bl_x",
            "bl_y",
            "cent_x",
            "cent_y",
            "text",
        ],
    )

    results_frame = results_frame.loc[results_frame["text"].str.len() > 2]
    logging.debug("Боксы: переведены во фрейм")
    logging.info("Боксы: готовы")

    if debug:
        results_frame.to_excel("xlsx/boxes.xlsx", index=False)
        logging.debug("Боксы: документ сохранен")

        for result in results_frame.itertuples(index=False):
            tl = (result.tl_x, result.tl_y)
            br = (result.br_x, result.br_y)
            cv2.rectangle(screenshot, tl, br, (0, 255, 0), 1)

            # cent = (result.cent_x, result.cent_y)
            # cv2.circle(screenshot, cent, 0, (0, 0, 255), 3) # отрисовка центра бокса

            cv2.putText(
                screenshot,
                result.text,
                (tl[0], tl[1] - 5),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.4,
                (0, 255, 0),
                1,
            )

        cv2.imwrite("images/screenshot_highlited.png", screenshot)
        logging.debug("Боксы: изображение сохранено")

    return results_frame


def get_targets(boxes_frame, name):
    anchor_y = boxes_frame.loc[boxes_frame["text"] == "name", "cent_y"].values[
        0
    ]

    cor_delta = 100
    targets = boxes_frame.loc[
        (boxes_frame["text"].str.contains(name))
        & (
            boxes_frame["cent_y"].between(
                anchor_y - cor_delta, anchor_y + cor_delta
            )
        )
    ]
    logging.info("Цели: получены")

    if debug:
        targets.to_excel("xlsx/targets.xlsx", index=False)
        logging.debug("Цели: документ сохранен")

    return targets


while True:
    screenshot = get_screenshot()
    boxes_frame = get_boxes(screenshot)
    get_targets(boxes_frame, "asteroid")
    input("Следущий скриншот - enter")
