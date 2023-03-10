import time
import random
import math
import logging
import numpy as np
import pandas as pd
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

save_result = True


def save_highlighted_screenshot(screenshot, boxes, filename):
    new_image = screenshot.copy()

    for block in boxes["block_num"].unique():
        boxes_in_block = boxes.loc[boxes["block_num"] == block]
        # Повторяемая генерация rgb-цвета для числа
        col = (
            round(math.sin(0.024 * block * 255 / 3 + 0) * 127 + 128),
            round(math.sin(0.024 * block * 255 / 3 + 2) * 127 + 128),
            round(math.sin(0.024 * block * 255 / 3 + 4) * 127 + 128),
        )
        for r in boxes_in_block.itertuples():
            (x, y, w, h) = (
                r.left,
                r.top,
                r.width,
                r.height,
            )
            cv2.rectangle(new_image, (x, y), (x + w, y + h), col, 2)
            cv2.putText(
                new_image,
                f"'{r.text}' x:{r.left} y:{r.top} l:{r.line_num}",
                (x, y - 10),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.4,
                col,
                1,
            )
    cv2.imwrite(f"images/{filename}.png", new_image)


def get_screenshot():
    eve_window = gw.getWindowsWithTitle(win_name)[0]

    was_minimized = False
    if eve_window.isMinimized:
        was_minimized = True

    eve_window.maximize()
    eve_window.activate()

    if save_result:
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
    raw_boxes = pytesseract.image_to_data(
        screenshot,
        lang="eng",
        output_type=Output.DATAFRAME,
        config="--psm 3 -c preserve_interword_spaces=1",
    )

    if save_result:
        raw_boxes.to_excel("xlsx/0_raw_boxes.xlsx", index=False)

    base_boxes = raw_boxes.loc[raw_boxes["text"].str.len() > 3]

    if save_result:
        save_highlighted_screenshot(
            screenshot, base_boxes, "1_base_highlighted_screenshot"
        )
        base_boxes.to_excel("xlsx/1_base_boxes.xlsx", index=False)

    logging.info("Боксы получены")

    return base_boxes


def union_boxes(base_boxes):
    result_phrases = pd.DataFrame(
        columns=[
            "left",
            "top",
            "width",
            "height",
            "text",
            "block_num",
            "line_num",
        ]
    )
    for box in base_boxes["block_num"].unique():
        words_in_blocks = base_boxes.loc[base_boxes["block_num"] == box]

        grouped_words = words_in_blocks.groupby("line_num", as_index=False)

        box_phrases = grouped_words["width"].sum()
        box_phrases = box_phrases.merge(
            grouped_words["height"].max(), on="line_num", how="left"
        )
        box_phrases = box_phrases.merge(
            grouped_words["left"].min(), on="line_num", how="left"
        )
        box_phrases = box_phrases.merge(
            grouped_words["top"].min(), on="line_num", how="left"
        )
        box_phrases = box_phrases.merge(
            grouped_words["text"].apply(" ".join),
            on="line_num",
            how="left",
        )
        box_phrases["block_num"] = box

        rightest_box = words_in_blocks.loc[
            words_in_blocks["left"] == words_in_blocks["left"].max()
        ]
        leftest_box = words_in_blocks.loc[
            words_in_blocks["left"] == words_in_blocks["left"].min()
        ]
        box_phrases["width"] = (
            rightest_box.iloc[0].left
            + rightest_box.iloc[0].width
            - leftest_box.iloc[0].left
        )

        result_phrases = pd.concat([result_phrases, box_phrases])

    if save_result:
        result_phrases.to_excel("xlsx/2_unioned_boxes.xlsx", index=False)

    logging.info("Боксы объединены")

    return result_phrases


screenshot = get_screenshot()
base_boxes = get_boxes(screenshot)
unioned_boxes = union_boxes(base_boxes)

save_highlighted_screenshot(
    screenshot, unioned_boxes, "2_unioned_highlighted_screenshot"
)
