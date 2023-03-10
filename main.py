import time
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

save_result = True


def save_highlighted_screenshot(screenshot, boxes, filename):
    new_image = screenshot.copy()
    for r in boxes.itertuples():
        (x, y, w, h) = (
            r.left,
            r.top,
            r.width,
            r.height,
        )
        cv2.rectangle(new_image, (x, y), (x + w, y + h), (0, 255, 0), 2)
        cv2.putText(
            new_image,
            f"'{r.text}': {r.left}.{r.top} {r.width}.{r.height}",
            (x, y - 10),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.4,
            (0, 255, 0),
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
        config="--psm 4 -c preserve_interword_spaces=1",
    )
    base_boxes = raw_boxes.loc[(raw_boxes["text"].str.len() > 3)][
        ["left", "top", "width", "height", "text", "line_num", "word_num"]
    ]

    if save_result:
        save_highlighted_screenshot(
            screenshot, base_boxes, "1_base_highlighted_screenshot"
        )
        base_boxes.to_excel("xlsx/base_boxes.xlsx", index=False)

    logging.info("Боксы получены")

    return base_boxes


def union_boxes(base_boxes):
    grouped_words = base_boxes.groupby("line_num", as_index=False)
    result_boxes = grouped_words["width"].sum()
    result_boxes = result_boxes.merge(
        grouped_words["height"].max(), on="line_num", how="left"
    )
    result_boxes = result_boxes.merge(
        grouped_words["left"].min(), on="line_num", how="left"
    )
    result_boxes = result_boxes.merge(
        grouped_words["top"].min(), on="line_num", how="left"
    )
    result_boxes = result_boxes.merge(
        grouped_words["text"].apply(" ".join),
        on="line_num",
        how="left",
    )

    if save_result:
        result_boxes.to_excel("xlsx/unioned_boxes.xlsx", index=False)

    logging.info("Боксы объединены")

    return result_boxes


screenshot = get_screenshot()
base_boxes = get_boxes(screenshot)
unioned_boxes = union_boxes(base_boxes)

save_highlighted_screenshot(
    screenshot, unioned_boxes, "2_unioned_highlighted_screenshot"
)
