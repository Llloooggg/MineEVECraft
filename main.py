import time
import math
import logging

import numpy as np
import pygetwindow as gw
import pyautogui
import cv2
import easyocr


win_name = "EVE - Nostrom Stone"

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


screenshot = get_screenshot()

reader = easyocr.Reader(["en"], gpu=True)
result = reader.readtext(screenshot)


for bbox, text, prob in result:
    (tl, tr, br, bl) = bbox
    tl = (int(tl[0]), int(tl[1]))
    tr = (int(tr[0]), int(tr[1]))
    br = (int(br[0]), int(br[1]))
    bl = (int(bl[0]), int(bl[1]))

    text = "".join([c if ord(c) < 128 else "" for c in text]).strip()
    cv2.rectangle(screenshot, tl, br, (0, 255, 0), 2)
    cv2.putText(
        screenshot,
        text,
        (tl[0], tl[1] - 10),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.8,
        (0, 255, 0),
        2,
    )


print(result)

cv2.imshow("screenshot", screenshot)
cv2.waitKey(0)
