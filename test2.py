"""
QR Excel Processor (рефакторинг)

Требуемые библиотеки:
    pip install openpyxl pillow numpy opencv-python zxing-cpp pyzbar

Если pyzbar не работает в Windows, можно удалить его из кода —
ZXing будет использоваться как основной движок.
"""

import io
import logging
from openpyxl import load_workbook
from PIL import Image, ImageOps, ImageEnhance
import numpy as np
import cv2
import zxingcpp

try:
    from pyzbar.pyzbar import decode as zbar_decode
    HAS_ZBAR = True
except Exception:
    HAS_ZBAR = False

logging.basicConfig(level=logging.INFO, format="%(message)s")


# ---------- Подготовка изображения ----------

def prepare_variants(image_bytes):
    """Создает несколько вариантов изображения для повышения
    вероятности успешного распознавания QR."""
    img = Image.open(io.BytesIO(image_bytes)).convert("RGBA")

    bg = Image.new("RGBA", img.size, "WHITE")
    bg.alpha_composite(img)
    rgb = bg.convert("RGB")

    variants = []

    # оригинал
    variants.append(rgb)

    # оттенки серого
    gray = ImageOps.grayscale(rgb)
    variants.append(gray)

    # увеличенный
    variants.append(gray.resize((gray.width * 2, gray.height * 2)))

    # высокий контраст
    contrast = ImageEnhance.Contrast(gray).enhance(2.5)
    variants.append(contrast)

    # бинаризация
    thr = contrast.point(lambda p: 255 if p > 140 else 0)
    variants.append(thr)

    # OpenCV адаптивная бинаризация
    cv = np.array(gray)
    cv = cv2.adaptiveThreshold(
        cv,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31,
        5,
    )
    variants.append(Image.fromarray(cv))

    return variants


# ---------- ZXing ----------

def decode_zxing(img):
    """Пытается прочитать QR через ZXing."""
    try:
        result = zxingcpp.read_barcode(img)
        if result:
            return result.text
    except Exception:
        pass
    return None


# ---------- pyzbar ----------

def decode_pyzbar(img):
    """Резервное распознавание через pyzbar."""
    if not HAS_ZBAR:
        return None

    try:
        codes = zbar_decode(img)
        if codes:
            return codes[0].data.decode("utf-8", errors="ignore")
    except Exception:
        pass

    return None


# ---------- Основное распознавание ----------

def decode_image_bytes(image_bytes):
    """Перебирает варианты изображения и движки распознавания."""
    for variant in prepare_variants(image_bytes):

        text = decode_zxing(variant)
        if text:
            break

        text = decode_pyzbar(variant)
        if text:
            break
    else:
        return None

    # исправление cp1251
    try:
        text = text.encode("latin1").decode("cp1251")
    except Exception:
        pass

    return text


# ---------- Разбор ГОСТ QR ----------

def parse_gost_qr(text):
    if not text:
        return {}

    data = {}

    parts = text.split("|")

    if parts:
        data["Format"] = parts[0]

    for part in parts[1:]:
        if "=" in part:
            k, v = part.split("=", 1)
            data[k] = v

    return data


# ---------- Обработка ----------

def process_all_qrs(source_excel, target_excel):

    source = load_workbook(source_excel)
    sheet = source.active

    target = load_workbook(target_excel)
    ws = target.active

    # строим индекс ЛС → строка
    ls_index = {}

    for row in range(2, ws.max_row + 1):
        val = ws.cell(row=row, column=1).value
        if val:
            ls_index[str(val).strip()] = row

    logging.info(f"Изображений: {len(sheet._images)}")

    ok = 0
    fail = 0

    for i, img in enumerate(sheet._images, 1):

        logging.info(f"\n[{i}] обработка...")

        qr = decode_image_bytes(img._data())

        if not qr:
            logging.info("QR не найден.")
            fail += 1
            continue

        info = parse_gost_qr(qr)

        ls = info.get("persAcc", "").strip()

        if ls not in ls_index:
            logging.info(f"ЛС {ls} отсутствует.")
            fail += 1
            continue

        row = ls_index[ls]

        purpose = info.get("Purpose", "")
        group = "гр. 1"

        if "гр. " in purpose:
            try:
                group = "гр. " + purpose.split("гр. ")[1].split()[0]
            except Exception:
                pass

        fio = info.get("childFio", "")

        parts = fio.split()
        if len(parts) >= 2:
            fio = parts[0] + " " + parts[1]

        ws.cell(row=row, column=3).value = group
        ws.cell(row=row, column=4).value = fio
        ws.cell(row=row, column=5).value = group
        ws.cell(row=row, column=6).value = "Устиновский район"
        ws.cell(row=row, column=7).value = info.get("PayeeINN", "") + "_1"

        cbc = info.get("CBC", "")
        ws.cell(row=row, column=8).value = cbc[-4:]

        ws.cell(row=row, column=9).value = info.get("OKTMO", "")

        s = info.get("Sum", "0")
        value = int(s) / 100 if s.isdigit() else 0

        cell = ws.cell(row=row, column=10)
        cell.value = value
        cell.number_format = "0.00"

        ok += 1
        logging.info(f"OK: {ls}")

    target.save(target_excel)

    logging.info("\nГотово")
    logging.info(f"Успешно: {ok}")
    logging.info(f"Ошибок: {fail}")


if __name__ == "__main__":

    SOURCE = "Форма № ПД-4 с QR-кодом.xlsx"
    TARGET = "ИТОГ.xlsx"

    process_all_qrs(SOURCE, TARGET)
