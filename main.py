import time
import random
import logging

import numpy as np
import cv2
import pygetwindow as gw
import pyautogui as pg
from pyclick import HumanClicker
from pytesseract import pytesseract as pt
import pandas as pd

debug = True

win_name = "EVE - Nostrom Stone"

pt.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(message)s",
    datefmt="%d-%b-%y %H:%M:%S",
    handlers=[logging.StreamHandler()],
)

logging.info("Бот: запущен")

hc = HumanClicker()


def move_mouse(x, y):
    hc.move(
        (x, y - 3),
        random.uniform(0.1, 0.4),
    )


def click_mouse(x, y, right=False, runaway=True):
    move_mouse(x + random.randrange(-5, 5), y + random.randrange(-3, 3))
    if right:
        pg.click(button="right")
    else:
        pg.click()
        time.sleep(random.uniform(0.1, 0.5))
        if runaway:
            move_mouse(
                x - random.randrange(500, 700), y - random.randrange(-300, 300)
            )
    time.sleep(random.uniform(0.1, 0.5))


def get_screenshot():
    eve_window = gw.getWindowsWithTitle(win_name)[0]

    # was_minimized = False
    # if eve_window.isMinimized:
    #     was_minimized = True

    eve_window.maximize()
    eve_window.activate()

    if debug:
        filepath = "images/screenshot.png"
    else:
        filepath = None

    time.sleep(0.5)
    screenshot = pg.screenshot(
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

    # if was_minimized:
    #     eve_window.minimize()

    return cv2.cvtColor(np.array(screenshot), cv2.COLOR_BGR2RGB)


def highlite_boxes(boxes, module_name, file_name):
    canvas = screenshot.copy()
    for result in boxes.itertuples(index=False):
        tl = (result.tl_x, result.tl_y)
        br = (result.br_x, result.br_y)
        cv2.rectangle(canvas, tl, br, (0, 255, 0), 1)

        cent = (result.cent_x, result.cent_y)
        cv2.circle(canvas, cent, 0, (0, 0, 255), 3)  # отрисовка центра бокса

        cv2.putText(
            canvas,
            f"{result.text} {cent}",
            (tl[0], tl[1] - 10),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.4,
            (0, 255, 0),
            1,
        )
    cv2.imwrite(f"images/{file_name}.png", canvas)
    logging.debug(f"{module_name}: изображение сохранено")


def get_boxes(screenshot):
    results = pt.image_to_data(
        cv2.bitwise_not(cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)),
        lang="eng",
        output_type=pt.Output.DATAFRAME,
        config="--psm 3",
    )
    logging.debug("Боксы: получены")
    results = results.loc[
        (results["conf"] > 30)
        & (results["text"].notnull())
        & (len(results["text"].str.strip()) > 0)
    ]

    results_frame = pd.DataFrame(
        [
            [
                result.left,  # tl_x
                result.top,  # tl_y
                result.left + result.width,  # tr_x
                result.top,  # tr_y
                result.left + result.width,  # br_x
                result.top + result.height,  # br_y
                result.left,  # bl_x
                result.top + result.height,  # bl_y
                int(result.left + result.width / 2),  # cent_x
                int(result.top + result.height / 2),  # cent_y
                result.text.lower(),  # text
            ]
            for result in results.itertuples(index=False)
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
    logging.debug("Боксы: переведены в удобный фрейм")
    logging.info("Боксы: готовы")

    if debug:
        results_frame.to_excel("xlsx/boxes.xlsx", index=False)
        logging.debug("Боксы: документ сохранен")

        highlite_boxes(results_frame, "Боксы", "boxes")

    return results_frame


def get_targets(boxes_frame, text=False, x_delta=100):
    anchor_top_x, anchor_top_y = boxes_frame.loc[
        boxes_frame["text"].str.contains("nam"), ["cent_x", "cent_y"]
    ].values[0]
    anchor_bot_y = boxes_frame.loc[
        boxes_frame["text"] == "hobgoblin", "cent_y"
    ].values[0]

    targets = boxes_frame.loc[
        (boxes_frame["cent_y"].between(anchor_top_y + 10, anchor_bot_y - 10))
        & (
            boxes_frame["cent_x"].between(
                anchor_top_x - 100, anchor_top_x + x_delta
            )
        )
    ]
    if text:
        targets = targets.loc[targets["text"].str.contains(text)]
    logging.info("Цели: получены")

    if debug:
        targets.to_excel("xlsx/targets.xlsx", index=False)
        logging.debug("Цели: документ сохранен")

        highlite_boxes(targets, "Цели", "targets")

    return targets


def get_cors_by_unique_name(boxes_frame, name):
    sub_ftame = boxes_frame.loc[
        boxes_frame["text"] == name, ["cent_x", "cent_y"]
    ]
    if not sub_ftame.empty:
        x, y = sub_ftame.values[0]
        return (x, y)
    else:
        return None


def go_to_minefield():
    global screenshot
    logging.info("Перемещение к месту добычи: начато")
    while True:
        screenshot = get_screenshot()
        boxes_frame = get_boxes(screenshot)
        target = get_targets(boxes_frame, "belt", x_delta=150)
        if target.empty:
            logging.warning(
                "Перемещение к месту добычи: потенциальные цели не найдены"
            )
            continue

        target = target.sample().iloc[0]

        click_mouse(target.cent_x, target.cent_y, right=True, runaway=False)

        screenshot = get_screenshot()
        boxes_frame = get_boxes(screenshot)

        if get_cors_by_unique_name(
            boxes_frame, "look"
        ):  # защита при ошибочном состоянии
            logging.warning(
                "Перемещение к месту добычи: цель добычи уже зафиксирована"
            )
            return

        warp_cor = get_cors_by_unique_name(boxes_frame, "warp")
        if warp_cor:
            click_mouse(warp_cor[0], warp_cor[1])
            logging.info("Перемещение к месту добычи: перемещене корабля")
            return


def start_mine():
    global screenshot
    logging.info("Добыча: начата")
    while True:
        screenshot = get_screenshot()
        boxes_frame = get_boxes(screenshot)
        target = get_targets(boxes_frame, "(veldspar)|(scordite)").iloc[0]
        if target.empty:
            logging.warning("Добыча: потенциальные цели не найдены")
            continue

        click_mouse(target.cent_x, target.cent_y, True)

        screenshot = get_screenshot()
        boxes_frame = get_boxes(screenshot)

        target_lock_cor = get_cors_by_unique_name(
            boxes_frame, "lock"
        )  # свдиг примерно 40, 90
        if target_lock_cor:
            click_mouse(target_lock_cor[0], target_lock_cor[1])
            time.sleep(random.uniform(4.4, 5.8))
            pg.press("f1")
            time.sleep(random.uniform(0.1, 1))
            pg.press("f2")
            logging.info("Добыча: цель добычи найдена")
            return
        approach_cor = get_cors_by_unique_name(boxes_frame, "approach")
        if approach_cor:
            click_mouse(approach_cor[0], approach_cor[1])
            logging.info("Добыча: цель сближения найдена")
            time.sleep(random.uniform(20, 25))
        # решение через сдвиг
        # click_mouse(target.cent_x + 40, target.cent_y + 20, runaway=False)
        # move_mouse(
        #     target.cent_x - random.randrange(500, 700),
        #     target.cent_y - random.randrange(50, 400),
        # )


"""
while True:
    screenshot = get_screenshot()
    boxes_frame = get_boxes(screenshot)
    targets = get_targets(boxes_frame, "belt")

    input("Следущий скриншот - enter")
"""


def main(current_state="EMPTY"):
    if current_state == "UNDOCKED":
        go_to_minefield()
        current_state = "ON_MINEFILD"
        time.sleep(random.uniform(20, 25))
    if current_state == "ON_MINEFILD":
        start_mine()
        current_state = "MINING"


main("UNDOCKED")
