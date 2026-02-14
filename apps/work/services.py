import os
import zipfile
import shutil
import re
import math
import hashlib
import fitz # PyMuPDF
import easyocr
import cv2 # OpenCV для обработки изображений
import numpy as np
from PIL import Image
from decimal import Decimal, ROUND_HALF_UP, ROUND_DOWN
from datetime import datetime
import openpyxl
import warnings
warnings.filterwarnings("ignore", category=UserWarning) # Suppress torch/easyocr warnings
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import PatternFill, Font
import requests 
from bs4 import BeautifulSoup
from django.conf import settings
import difflib # For fuzzy matching

class NetworkError(Exception):
    def __init__(self, user_message, technical_details):
        self.user_message = user_message
        self.technical_details = technical_details
        super().__init__(user_message)

NBKR_URL = "https://www.nbkr.kg/index1.jsp?item=1562&lang=RUS&valuta_id=15"

def selected_usa_dollar(soup):
    option = soup.find("option", value="15")
    # print(f"[get_current_dollar_rate.selected_usa_dollar] option found: {option}")
    value_text = '<option selected="" value="15">1 Доллар США'
    if value_text in str(option):
        return True
    else:
        return False

def get_curs(soup, date):
    table = soup.find_all("tr")

    for tr in table:
        if f"{date}" in str(tr):
            td = tr.find("td", class_="stat-right")
            value = td.get_text()
            rate_decimal = Decimal(value.replace(",", "."))
            rate_truncated = (rate_decimal * Decimal("100")).quantize(Decimal("1"), rounding=ROUND_DOWN) / Decimal("100")
            print("Курс:", rate_truncated)
            return rate_truncated

def get_current_dollar_rate(date_str=None):
    try:
        resp = requests.get(NBKR_URL, timeout=10)
        html = resp.text
        # print(f"[get_current_dollar_rate] Fetched NBKR_URL status={resp.status_code} length={len(html)} for date={date_str}")
    except requests.RequestException as e:
        # print(f"[get_current_dollar_rate] RequestException: {e} (url={NBKR_URL})")
        user_message = "Проверьте подключение к интернету"
        technical_details = f"Ошибка при подключении к сайту НБКР: {str(e)}"
        raise NetworkError(user_message, technical_details)

    soup = BeautifulSoup(html, "html.parser")
    
    if not selected_usa_dollar(soup):
        raise Exception("На странице НБКР не выбран доллар США")
    
    if not date_str:
        print("[get_current_dollar_rate] date_str is empty or None")
        raise Exception("Дата не указана")
    
    rate = get_curs(soup, date_str)
    
    if rate:
        print(f"Получен курс доллара США на {date_str}: {rate}")
        return rate
    else:
        print(f"[get_current_dollar_rate] rate not found for date {date_str}")
        raise Exception(f"Не удалось найти курс на дату {date_str}")

FIELDS_MAP_TYPE_1 = {
    "Дата (1)": (880, 2520, 390, 140),
    "ФИО Водит. (4)": (850, 2450, 500, 150),
    "Кол.тон (7)": (1500, 1390, 300, 80),
    "Марка": (380, 2730, 350, 70),
    "Гос_номер ()": (-80, 2730, 700, 170), 
    "Якорь (1)": (0, 500, 500, 1000), # Anchor for alignment
}

FIELDS_MAP_TYPE_2 = {
    "Цена (8)": (1450, 2150, 250, 160),
    "№ счет факт (Инвойс) (16)": (500, 200, 500, 60),
}

FIELDS_MAP_TYPE_3 = {
    "№ сопров.накл. KZ (15)": (360, 250, 330, 200),
    "Дата сопр.накл (13)": (360, 250, 330, 200)
}

FIELDS_MAP_TYPE_2_PAGE_2 = {
    "Цена (8) Alt": (1500, 100, 150, 120),
}

MIN_HEIGHT_CONFIG = {
    "ФИО Водит. (4)": 20,
    "Марка_Гос_номер ()": 38,
}

try:
    reader = easyocr.Reader(["ru", "en"], gpu=False)
except Exception as e:
    print(f"Error initializing EasyOCR: {e}")
    reader = None

def deskew_image(img_cv):
    try:
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)

        edges = cv2.Canny(gray, 50, 150, apertureSize=3)

        lines = cv2.HoughLinesP(edges, 1, np.pi / 180, threshold=100, minLineLength=100, maxLineGap=10)

        angles = []
        if lines is not None:
            for line in lines:
                x1, y1, x2, y2 = line[0]

                angle_rad = math.atan2(y2 - y1, x2 - x1)
                angle_deg = math.degrees(angle_rad)

                if abs(angle_deg) > 45:
                    if angle_deg > 0:
                        deviation = angle_deg - 90
                    else:
                        deviation = angle_deg + 90

                    angles.append(deviation)

        if not angles:
            return img_cv

        median_angle = np.median(angles)

        if abs(median_angle) < 0.1:
            return img_cv

        print(f" [Deskew] Обнаружен перекос: {median_angle:.2f} градусов. Исправляем...")

        (h, w) = img_cv.shape[:2]
        center = (w // 2, h // 2)

        M = cv2.getRotationMatrix2D(center, median_angle, 1.0)

        rotated = cv2.warpAffine(
            img_cv, M, (w, h),
            flags=cv2.INTER_CUBIC,
            borderMode=cv2.BORDER_CONSTANT,
            borderValue=(255, 255, 255)
        )

        return rotated
    except Exception as e:
        print(f" [Deskew Error] Не удалось выровнять: {e}")
        return img_cv

class DataCleaner:
    @staticmethod
    def replace_ruble(text):
        if not text:
            return text
        return text.replace('₽', 'Р')

    @staticmethod
    def clean_1(text, context):
        match = re.search(r'\b(\d{2}\.\d{2}\.\d{4})\b', text)
        if match:
            return match.group(1)

        surname = context.get('surname')
        if surname:
            all_files = context.get('type_2_files', []) + context.get('type_3_files', [])
            for fpath in all_files:
                fname = os.path.basename(fpath)
                if surname.lower() in fname.lower():
                    match_file = re.search(r'(\d{2}\.\d{2}\.\d{4})', fname)
                    if match_file:
                        return match_file.group(1)

        zip_name = context.get('zip_filename', '')
        match_zip = re.search(r'(\d{2}-\d{2}-\d{4})', zip_name)
        if match_zip:
            return match_zip.group(1).replace('-', '.')

        return text

    @staticmethod
    def get_cleaned_big_3_list(data):
        if not isinstance(data, list):
            return []

        filtered = []
        for text, h in data:
            if 35 < h < 46:
                filtered.append(text.strip())

        while filtered and filtered[0] in ["25", "26", "27", "28", "29", "30"]:
            filtered.pop(0)

        return filtered

    @staticmethod
    def clean_plate_text(text):
        t = text.strip().upper()
        
        cyr_to_lat = {
            'А': 'A', 'В': 'B', 'Е': 'E', 'К': 'K', 'М': 'M', 'Н': 'H',
            'О': 'O', 'Р': 'P', 'С': 'C', 'Т': 'T', 'У': 'Y', 'Х': 'X'
        }
        for cyr, lat in cyr_to_lat.items():
            t = t.replace(cyr, lat)

        t = t.replace('L', 'I') 
        t = t.replace('|', 'I')
        t = t.replace('I', 'I') 
        
        t = t.replace('O', '0')
        t = t.replace('S', '5')
        
        t = re.sub(r'[^A-Z0-9/]', '', t)
        
        return t

    @staticmethod
    def clean_2(data, context): return ""
    @staticmethod
    def clean_3(data, context): return ""
    
    @staticmethod
    def clean_fio_raw(data, context):
        if not isinstance(data, list):
            return DataCleaner.replace_ruble(str(data).strip()), DataCleaner.replace_ruble(str(data).strip())

        filtered_items = []
        for text, h in data:
            if re.search(r'\d{2}\.\d{2}\.\d{4}', text):
                continue
                
            if 35 < h < 500:
                filtered_items.append((text.strip(), h))
        
        while filtered_items and filtered_items[0][0] in ["21", "22", "23", "24"]:
            filtered_items.pop(0)
        
        if filtered_items:
            first_item = filtered_items[0]
            t, h = first_item
            
            if t.endswith(':'):
                t = t[:-1] + '.'
            if t.endswith('-'):
                t = t[:-1] + '.'
            
            t = DataCleaner.replace_ruble(t)
            
            if not t.endswith('.'):
                t += '.'
            
            return t, t
            
        return "", ""

    @staticmethod
    def clean_4(text, context): return text.strip("'") if text else text
    @staticmethod
    def clean_5(text, context): return str(text).strip("'") if text else text
    @staticmethod
    def clean_6(text, context): return text
    
    @staticmethod
    def clean_7(text, context): 
        match = re.search(r'(\d{2}\s?\d{3})\s*нетто', text, re.IGNORECASE)
        if match:
            return match.group(1).replace(' ', '')
        return text

    @staticmethod
    def clean_8(text, context): 
        return DataCleaner.replace_ruble(text)
    @staticmethod
    def clean_9(text, context): return text
    @staticmethod
    def clean_10(text, context): return text
    @staticmethod
    def clean_11(text, context): return text
    @staticmethod
    def clean_12(text, context): return text
    @staticmethod
    def clean_13(text, context): return str(text).strip("'") if text else text
    
    @staticmethod
    def clean_14(text, context): 
        filename = context.get('filename', '')
        # Убираем расширение
        base_name = os.path.splitext(filename)[0]
        # Убираем точки (мусор)
        base_name = base_name.replace('.', '')
        # Убираем суффиксы CMP, СМП, СМР и т.д.
        base_name = re.sub(r'\s*(cmp|смп|смр|cmr)', '', base_name, flags=re.IGNORECASE)
        
        return base_name.strip("'").strip()
    
    @staticmethod
    def clean_15(text, context):
        match = re.search(r'(KZ-SNT-[\w-]+(?:\s+[\w-]+)*)', text)
        if match:
            return match.group(1).replace(" ", "")
        return text

    @staticmethod
    def clean_16(text, context): 
        return DataCleaner.replace_ruble(text)

    @staticmethod
    def clean_marka_gos_number(data, context):
        cleaned_list = DataCleaner.get_cleaned_big_3_list(data)
        return " / ".join(cleaned_list)

def normalize_surname(surname):
    if not surname:
        return []
    variants = {surname}
    replacements = {
        'i': 'и', 'o': 'о', 'a': 'а', 'e': 'е', 'c': 'с', 'p': 'р', 
        'y': 'у', 'x': 'х', 'H': 'Н', 'K': 'К', 'M': 'М', 'B': 'В', 'T': 'Т'
    }
    new_surname = surname
    for lat, cyr in replacements.items():
        new_surname = new_surname.replace(lat, cyr)
    variants.add(new_surname)

    # Добавляем латинскую версию фамилии для поиска по файлам (для казахских/русских имен)
    latin_variant = []
    for char in surname:
        if char in CYRILLIC_TO_LATIN:
            latin_variant.append(CYRILLIC_TO_LATIN[char])
        else:
            latin_variant.append(char)
    variants.add("".join(latin_variant))

    return list(variants)

def safe_decimal(value, field_name):
    if not value:
        return Decimal("0")
    cleaned = ""
    for ch in value:
        if ch.isdigit() or ch in ".,":
            cleaned += ch
    cleaned = cleaned.replace(",", ".")
    if cleaned == "":
        return Decimal("0")
    try:
        return Decimal(cleaned)
    except Exception as e:
        print(f"Ошибка Decimal для '{field_name}': '{value}' в '{cleaned}': {e}")
        return Decimal("0")

CYRILLIC_TO_LATIN = {
    'А': 'A', 'Б': 'B', 'В': 'V', 'Г': 'G', 'Д': 'D', 'Е': 'E', 'Ё': 'E',
    'Ж': 'Zh', 'З': 'Z', 'И': 'I', 'Й': 'Y', 'К': 'K', 'Л': 'L', 'М': 'M',
    'Н': 'N', 'О': 'O', 'П': 'P', 'Р': 'R', 'С': 'S', 'Т': 'T', 'У': 'U',
    'Ф': 'F', 'Х': 'H', 'Ц': 'Ts', 'Ч': 'Ch', 'Ш': 'Sh', 'Щ': 'Sch',
    'Ъ': '', 'Ы': 'Y', 'Ь': '', 'Э': 'E', 'Ю': 'Yu', 'Я': 'Ya',
    'Ә': 'A', 'Ғ': 'G', 'Қ': 'Q', 'Ң': 'N', 'Ө': 'O', 'Ұ': 'U', 'Ү': 'U', 'H': 'H', 'І': 'I',
    'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'e',
    'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm',
    'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u',
    'ф': 'f', 'х': 'h', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'sch',
    'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya',
    'ә': 'a', 'ғ': 'g', 'қ': 'q', 'ң': 'n', 'ө': 'o', 'ұ': 'u', 'ү': 'u', 'h': 'h', 'і': 'i'
}

def get_safe_filename(original_name, field_name):
    """
    Создает безопасное ASCII-имя файла на основе MD5-хеша исходного имени.
    Это решает проблему с кириллицей в именах файлов в Windows.
    """
    
    def transliterate(text):
        result = []
        for char in text:
            if char in CYRILLIC_TO_LATIN:
                result.append(CYRILLIC_TO_LATIN[char])
            elif char.isalnum() or char in '_-':
                result.append(char)
            else:
                result.append('_')
        return ''.join(result)
    
    combined = f"{original_name}_{field_name}"
    hash_obj = hashlib.md5(combined.encode('utf-8'))
    hash_hex = hash_obj.hexdigest()[:8]
    safe_base = os.path.splitext(os.path.basename(original_name))[0]
    safe_base = transliterate(safe_base)
    safe_base = re.sub(r'[^a-zA-Z0-9_-]', '_', safe_base)
    if len(safe_base) > 20:
        safe_base = safe_base[:20]
    
    safe_field = transliterate(field_name)
    safe_field = re.sub(r'[^a-zA-Z0-9_-]', '_', safe_field)
    safe_field = re.sub(r'_+', '_', safe_field).strip('_')
    
    return f"{safe_base}_{hash_hex}_{safe_field}.png"

def extract_text_from_pdf(pdf_path, coords_map, save_dir, apply_deskew=False, page_num=0):
    extracted_data = {}
    try:
        # print(f"[extract_text_from_pdf] Processing {pdf_path} with coords_map keys: {list(coords_map.keys())}, apply_deskew={apply_deskew}")
        doc = fitz.open(pdf_path)
        if page_num >= len(doc):
            print(f"[extract_text_from_pdf] Page {page_num} does not exist in {pdf_path}")
            return {}
        page = doc.load_page(page_num)
        pix = page.get_pixmap(dpi=300)
        
        img_np = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        
        if pix.n == 3:
            img_cv = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)
        elif pix.n == 4:
            img_cv = cv2.cvtColor(img_np, cv2.COLOR_RGBA2BGR)
        else:
            img_cv = cv2.cvtColor(img_np, cv2.COLOR_GRAY2BGR)

        if apply_deskew:
            img_cv = deskew_image(img_cv)

        img_rgb = cv2.cvtColor(img_cv, cv2.COLOR_BGR2RGB)
        img_full = Image.fromarray(img_rgb)

        os.makedirs(save_dir, exist_ok=True)

        os.makedirs(save_dir, exist_ok=True)

        offset_x = 0
        offset_y = 0

        # Check for explicit Anchor definition in the map
        anchor_rect = None
        anchor_key = None
        for k, v in coords_map.items():
            if "Якорь" in k or "Anchor" in k:
                anchor_rect = v
                anchor_key = k
                break
        
        if anchor_rect:
            # ONLY run anchor logic if the map has an anchor key
            ax, ay, aw, ah = anchor_rect
            # Safety checks
            ax = max(0, ax)
            ay = max(0, ay)
            anchor_crop = img_full.crop((ax, ay, ax + aw, ay + ah))
            
            # Save debug image
            base_fname = os.path.basename(pdf_path)
            anchor_img_name = f"debug_anchor_{base_fname}.png"
            anchor_img_path = os.path.join(save_dir, anchor_img_name)
            anchor_crop.save(anchor_img_path)
            
            if reader is not None:
                try:
                    print("=" * 35)
                    print(f"данные читаетсья вот из этого фото: {anchor_img_path}")
                    print("=" * 35)

                    # Fix: Pass numpy array to EasyOCR to avoid OpenCV 'can't open/read file' error with Cyrillic paths
                    # Convert PIL to BGR numpy array
                    anchor_np = cv2.cvtColor(np.array(anchor_crop), cv2.COLOR_RGB2BGR)
                    anchor_results = reader.readtext(anchor_np, detail=1)
                    
                    print("сырые данные из этого фото которые были взяты")
                    print("=" * 35)
                    
                    # DEBUG: Print all raw findings in anchor zone
                    # print(f"\n[ANCHOR DEBUG RAW] File: {base_fname}")
                    # print("  RAW (Все найденное в зоне якоря):")
                    if not anchor_results:
                        print("    (Пусто, OCR ничего не увидел)")
                    else:
                        for (bbox, text, prob) in anchor_results:
                             # bbox=[[x1,y1],[x2,y1],[x2,y2],[x1,y2]]
                             h = int(bbox[2][1] - bbox[0][1]) 
                             print(f"    - '{text}' (H: {h}, Prob: {prob:.2f})")

                    target_anchor_text = "ИНН" # ТО ЧТО МЫ ИЩЕМ (можно менять на "1" или "CMR" и т.д.)
                    
                    found_anchor = False
                    for (bbox, text, prob) in anchor_results:
                        if target_anchor_text in text:
                            # Found anchor. Offset is relative to the anchor box top-left
                            # The bbox is local to the crop.
                            local_x = int(bbox[0][0])
                            local_y = int(bbox[0][1])
                            
                            # Global shift:
                            # We expected '1' at (0,0) inside the crop (ideal case).
                            # Found at (local_x, local_y).
                            # Shift = local_x, local_y
                            
                            offset_x = local_x
                            offset_y = local_y
                            
                            print(f"[ANCHOR DEBUG] Якорь мы искали '{target_anchor_text}' нашли: '{text}'")
                            print(f"[ANCHOR DEBUG] Координаты которые мы ожидаем: х=0, у=0")
                            print(f"[ANCHOR DEBUG] Координаты найденного якоря: x={offset_x}, y={offset_y}")
                            print(f"[ANCHOR DEBUG] Расчет смещения: сдвиг по х={offset_x}, сдвиг по у={offset_y}")
                            
                            found_anchor = True
                            break
                    
                    if not found_anchor:
                         print(f"[ANCHOR DEBUG] Якорь '{target_anchor_text}' не найден в {anchor_rect}. Смещение (0,0).")
                         print(f"[ANCHOR DEBUG] Сохранено фото области поиска для проверки: {anchor_img_path}")
                except Exception as e:
                    print(f"[ANCHOR] Error: {e}")

        for field_name, (x0, y0, w, h) in coords_map.items():
            # Skip the anchor field itself if it shouldn't be extracted as data
            if field_name == anchor_key:
                continue
                
            # Apply anchor offset
            if offset_x != 0 or offset_y != 0:
                 x0 = max(0, x0 + offset_x)
                 y0 = max(0, y0 + offset_y)
                 if anchor_key: # Only log if we actually used an anchor
                     print(f"[ANCHOR DEBUG] Применяем к полю '{field_name}'... New coords: ({x0}, {y0})")

            x1 = min(x0 + w, img_full.width)
            y1 = min(y0 + h, img_full.height)
            
            crop_img = img_full.crop((x0, y0, x1, y1))
            
            if any(x in field_name for x in ["ФИО Водит.", "Марка", "Гос_номер"]):
                r, g, b = crop_img.split()
                crop_img = b 
                
                if "ФИО Водит." in field_name:
                    threshold = 170
                    crop_img = crop_img.point(lambda p: 255 if p > threshold else 0)
                elif "Гос_номер" in field_name:
                    threshold = 165
                    crop_img = crop_img.point(lambda p: 255 if p > threshold else 0)
            
            img_filename = get_safe_filename(pdf_path, field_name)
            img_path = os.path.join(save_dir, img_filename)
            crop_img.save(img_path)
            # print(f"[extract_text_from_pdf] Saved image: {img_filename} (original: {os.path.basename(pdf_path)}_{field_name})")

            if reader is None:
                print(f"[extract_text_from_pdf] OCR reader is not initialized. Skipping OCR for {img_path}")
                results = []
            else:
                try:
                    if not os.path.exists(img_path):
                        print(f"[extract_text_from_pdf] ERROR: Image file does not exist: {img_path}")
                        results = []
                    else:
                        test_img = cv2.imread(img_path)
                        if test_img is None:
                            print(f"[extract_text_from_pdf] ERROR: cv2.imread returned None for {img_path}. File may have encoding issues.")
                            results = []
                        else:
                            results = reader.readtext(img_path, detail=1)
                            # print(f"[extract_text_from_pdf] OCR successful for {img_filename}, found {len(results)} text regions")
                except Exception as e:
                    print(f"[extract_text_from_pdf] OCR error for {img_path}: {e}")
                    import traceback
                    traceback.print_exc()
                    results = []
            text_parts = []
            raw_items = []
            
            min_height = MIN_HEIGHT_CONFIG.get(field_name, 0)
            
            for (bbox, text, prob) in results:
                height = int(((bbox[3][1] - bbox[0][1]) + (bbox[2][1] - bbox[1][1])) / 2)
                
                raw_items.append((text, height))

                if height >= min_height:
                    text_parts.append(text)
                else:
                    pass
                    # print(f"[extract_text_from_pdf] Filtered out text '{text}' in '{field_name}' due to height {height} < {min_height}")
            
            if any(x in field_name for x in ["ФИО Водит.", "Марка", "Гос_номер"]):
                extracted_data[field_name] = raw_items
            else:
                extracted_data[field_name] = " ".join(text_parts)
                extracted_data[field_name] = " ".join(text_parts).strip()
        
        doc.close()
    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")
    
    return extracted_data
def extract_data_from_xlsx(xlsx_path):
    extracted_data = {}
    try:
        print(f"[extract_data_from_xlsx] Loading xlsx {xlsx_path}")
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        sheet = wb.active
        
        def get_val(cell_ref):
            val = sheet[cell_ref].value
            return str(val).strip() if val is not None else ""

        date_val = get_val("K75")
        if not date_val:
            date_val = get_val("K76")
        extracted_data["Дата (1)"] = date_val

        extracted_data["Марка_XLSX"] = get_val("G89")

        plate_1 = get_val("B89")
        plate_2 = get_val("B90")
        plate_1_clean = DataCleaner.clean_plate_text(plate_1) if plate_1 else ""
        plate_2_clean = DataCleaner.clean_plate_text(plate_2) if plate_2 else ""
        extracted_data["Гос.номер_XLSX"] = f"{plate_1_clean} / {plate_2_clean}"

        extracted_data["ФИО Водит. (4)"] = get_val("M80")

        extracted_data["Кол.тон (7)"] = get_val("U43")

        wb.close()
    except Exception as e:
        print(f"[extract_data_from_xlsx] Error processing XLSX {xlsx_path}: {e}")
    
    source_map = {
        1: "K75/K76",
        2: "G89",
        3: "B89/B90",
        4: "M80",
        7: "U43"
    }
    
    return extracted_data, source_map

def process_zip_file(zip_file, dollar_rate, selected_date, tn_ved_code, bnd_code, nds_percent, save_photos=False):
    base_temp_dir = os.path.join(settings.MEDIA_ROOT, "temp_ocr")
    upload_dir = os.path.join(base_temp_dir, "upload")
    extract_dir = os.path.join(base_temp_dir, "extracted")
    imgs_root_dir = os.path.join(settings.MEDIA_ROOT, "imgs")
    preview_imgs_dir = os.path.join(base_temp_dir, "preview_imgs")

    if os.path.exists(base_temp_dir):
        print(f"[process_zip_file] Cleaning up old temp dir: {base_temp_dir}")
        shutil.rmtree(base_temp_dir)
    else:
        print(f"[process_zip_file] No old temp dir tokens to clean: {base_temp_dir}")
    
    os.makedirs(imgs_root_dir, exist_ok=True)
    os.makedirs(upload_dir, exist_ok=True)
    os.makedirs(extract_dir, exist_ok=True)
    os.makedirs(preview_imgs_dir, exist_ok=True)

    zip_path = os.path.join(upload_dir, "upload.zip")
    with open(zip_path, 'wb+') as destination:
        total_written = 0
        for chunk in zip_file.chunks():
            try:
                destination.write(chunk)
                total_written += len(chunk)
            except Exception as e:
                print(f"[process_zip_file] Error writing chunk to {zip_path}: {e}")
        print(f"[process_zip_file] Wrote zip to {zip_path}, bytes={total_written}")

    print(f"Extracting ZIP to: {extract_dir}")
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

    type_1_files = []
    type_2_files = []
    type_3_files = []

    print("Scanning extracted files...")
    for root, dirs, files in os.walk(extract_dir):
        for file in files:
            full_path = os.path.join(root, file)
            
            if re.match(r'^[\d\.]+(\s*(cmp|смп|смр|cmr))?\s*\.(pdf|xlsx)$', file, re.IGNORECASE):
                type_1_files.append(full_path)
                print(f" [SCAN] Found Type 1 (Main): {file}")
            elif file.lower().endswith('.pdf'):
                if re.match(r'^(эсф|электронный\s*(-)?\s*счет\s*(-)?\s*фактура)', file.lower()):
                    type_2_files.append(full_path)
                    print(f" [SCAN] Found Type 2 (ESF): {file}")
                elif re.match(r'^(снт|сопроводительная\s*накладная\s*(на)?\s*товары)', file.lower()):
                    type_3_files.append(full_path)
                    print(f" [SCAN] Found Type 3 (SNT): {file}")
                else:
                    # print(f" [SCAN] Ignored PDF: {file}")
                    pass
            else:
                # print(f" [SCAN] Ignored file: {file}")
                pass


    # print(f"[process_zip_file] Found {len(type_1_files)} type_1, {len(type_2_files)} type_2, {len(type_3_files)} type_3 files")

    if not type_1_files:
        print(f"[process_zip_file] No type_1 files found in {extract_dir}")
        raise Exception("В архиве не найдены файлы основных документов (например, '1.pdf' или '1.xlsx'). Проверьте структуру архива.")

    final_results = []
    
    used_type_2 = set()
    used_type_3 = set()
    
    driver_debug_info = [] # Store debug string for each driver

    for obj_idx, t1_path in enumerate(type_1_files):
        print(f"Processing Type 1 file: {t1_path} (Basename: {os.path.basename(t1_path)})")
        
        temp_img_path = os.path.join(base_temp_dir, f"temp_imgs_processing_obj_{obj_idx}")
        if os.path.exists(temp_img_path):
            shutil.rmtree(temp_img_path)
        os.makedirs(temp_img_path, exist_ok=True)
        
        is_xlsx = t1_path.lower().endswith('.xlsx')
        t1_data = {}
        
        if is_xlsx:
            t1_data, source_map = extract_data_from_xlsx(t1_path)
            if isinstance(source_map, dict):
                source_map.pop(1, None)
        else:
            t1_data = extract_text_from_pdf(t1_path, FIELDS_MAP_TYPE_1, temp_img_path, apply_deskew=True)
            source_map = {}
        
        fio_raw_data = t1_data.get("ФИО Водит. (4)", [])
        
        if is_xlsx:
            fio_str = str(fio_raw_data).strip()
            fio_formatted = fio_str
            fio_clean = fio_str
        else:
            fio_formatted, fio_clean = DataCleaner.clean_fio_raw(fio_raw_data, {})
        
        surname_full = fio_clean
        
        context = {
            'surname': surname_full.split()[0].strip() if surname_full else "",
            'zip_filename': zip_file.name,
            'type_2_files': type_2_files,
            'type_3_files': type_3_files,
            'filename': os.path.basename(t1_path)
        }
        
        surname_clean = surname_full.split()[0].strip() if surname_full else "Unknown"
        print(f"[MATCH DEBUG] File: {os.path.basename(t1_path)}")
        print(f" [MATCH DEBUG] Extracted Raw FIO: {fio_raw_data}")
        print(f" [MATCH DEBUG] Cleaned Surname: '{surname_clean}'")
        
        # DEBUG: Print raw FIO data and selected name
        # if not is_xlsx:
        #     print(f"\n[DEBUG FIO] File: {os.path.basename(t1_path)}")
            
        #     print("  RAW (Все найденное):")
        #     for t, h in fio_raw_data:
        #         print(f"    - '{t}' (H: {h})")
            
        #     print("  FILTERED (Прошло фильтр > 35):")
        #     passed_filter = []
        #     for t, h in fio_raw_data:
        #         if 35 < h < 500 and not re.search(r'\d{2}\.\d{2}\.\d{4}', t):
        #             passed_filter.append(f"    - '{t}' (H: {h})")
        #     if passed_filter:
        #         print("\n".join(passed_filter))
        #     else:
        #         print("    (Пусто)")
                
        #     print(f"  SELECTED (Итого): '{surname_clean}'\n")
        #     d_info = f"\n[FILE: {os.path.basename(t1_path)}]\n  RAW (Все найденное):\n"
        #     if not fio_raw_data:
        #         d_info += "    (Нет данных)\n"
        #     else:
        #         for t, h in fio_raw_data: d_info += f"    - '{t}' (H: {h})\n"
            
        #     d_info += "  FILTERED (Прошло фильтр > 35):\n"
        #     if passed_filter: 
        #         d_info += "\n".join(passed_filter) + "\n"
        #     else: 
        #         d_info += "    (Пусто)\n"
            
        #     d_info += f"  SELECTED (Итого): '{surname_clean}'\n"
        #     driver_debug_info.append(d_info)
        # else:
        #      print(f"\n[DEBUG FIO] File: {os.path.basename(t1_path)}")
        #      print(f"  Raw Data (XLSX): {fio_raw_data}")
        #      print(f"  Selected: {surname_clean}\n")
        #      driver_debug_info.append(f"[XLSX] {os.path.basename(t1_path)}: Selected='{surname_clean}'")

        found_t2 = False
        found_t3 = False

        raw_date = t1_data.get("Дата (1)")
        date_clean = DataCleaner.clean_1(raw_date, context)
        if not date_clean:
            date_clean = "Unknown_Date"
        
        date_folder = date_clean.replace("/", "-").replace("\\", "-")

        person_img_dir = os.path.join(imgs_root_dir, date_folder, surname_clean)
        
        preview_obj_dir = os.path.join(preview_imgs_dir, f"obj_{len(final_results)}")
        os.makedirs(preview_obj_dir, exist_ok=True)
        preview_image_paths = []
        
        if os.path.exists(temp_img_path):
            for img_file in os.listdir(temp_img_path):
                src_path = os.path.join(temp_img_path, img_file)
                dst_path = os.path.join(preview_obj_dir, img_file)
                shutil.copy2(src_path, dst_path)
                rel_path = os.path.relpath(dst_path, settings.MEDIA_ROOT)
                preview_image_paths.append(rel_path.replace("\\", "/"))
        
        if save_photos:
            os.makedirs(person_img_dir, exist_ok=True)
            if os.path.exists(temp_img_path):
                for img_file in os.listdir(temp_img_path):
                    shutil.move(os.path.join(temp_img_path, img_file), os.path.join(person_img_dir, img_file))
        
        plate_val = ""
        car_val = ""
        big_3_cleaned_list = []

        if is_xlsx:
            plate_val = t1_data.get("Гос.номер_XLSX", "")
            car_val = t1_data.get("Марка_XLSX", "")
            big_3_cleaned_list = [car_val, plate_val]
        else:
            # --- Process BRAND ---
            brand_raw = t1_data.get("Марка", [])
            # For brand, we just take all text found in the box
            car_val = " ".join([t.strip() for t, h in brand_raw if t.strip()])
            
            # --- Process PLATE ---
            plate_raw = t1_data.get("Гос_номер ()", [])
            plate_cleaned_list = DataCleaner.get_cleaned_big_3_list(plate_raw)
            
            # Debug: Capture filtered items with heights for Plate
            filtered_debug = []
            for text, h in plate_raw:
                if 35 < h < 46:
                     filtered_debug.append(f"{text} (H: {h})")
            filtered_debug_str = " | ".join(filtered_debug)
            
            if plate_cleaned_list:
                if len(plate_cleaned_list) >= 2:
                    # Assume first is number, last is region, or vice versa. 
                    # Usually: [Number, Region]
                    p1 = DataCleaner.clean_plate_text(plate_cleaned_list[0])
                    p2 = DataCleaner.clean_plate_text(plate_cleaned_list[-1])
                    plate_val = f"{p1} / {p2}"
                else:
                    plate_val = DataCleaner.clean_plate_text(plate_cleaned_list[0])
        
        if selected_date:
            user_date_str = selected_date.strftime('%d.%m.%Y') if hasattr(selected_date, 'strftime') else str(selected_date)
        else:
            user_date_str = ""
        
        field_images = {}
        
        if os.path.exists(preview_obj_dir):
            print(f"[process_zip_file] Scanning preview_obj_dir: {preview_obj_dir}")
            for img_file in os.listdir(preview_obj_dir):
                img_path = os.path.join(preview_obj_dir, img_file)
                if os.path.isfile(img_path):
                    rel_path = os.path.relpath(img_path, settings.MEDIA_ROOT)
                    rel_path = rel_path.replace("\\", "/")
                    
                    img_file_lower = img_file.lower()
                    print(f"[process_zip_file] Checking file: {img_file}")
                    
                    if (("date" in img_file_lower or "дата" in img_file_lower) and 
                        ("_1" in img_file_lower or "(1)" in img_file or img_file_lower.endswith("_1.png"))):
                        if 1 not in field_images:
                            field_images[1] = []
                        field_images[1].append(rel_path)
                        print(f"[process_zip_file] Added to field_images[1]: {img_file}")
                    elif (("fio" in img_file_lower or "фио" in img_file_lower or "vodit" in img_file_lower or "водит" in img_file_lower) and 
                          ("_4" in img_file_lower or "(4)" in img_file or "4" in img_file_lower)):
                        if 4 not in field_images:
                            field_images[4] = []
                        field_images[4].append(rel_path)
                        print(f"[process_zip_file] Added to field_images[4]: {img_file}")
                    elif (("kol" in img_file_lower or "кол" in img_file_lower) and 
                          ("ton" in img_file_lower or "тон" in img_file_lower or "_7" in img_file_lower or "(7)" in img_file)):
                        if 7 not in field_images:
                            field_images[7] = []
                        field_images[7].append(rel_path)
                        print(f"[process_zip_file] Added to field_images[7]: {img_file}")
                    elif (("marka" in img_file_lower or "марка" in img_file_lower) and 
                          not ("gos" in img_file_lower or "гос" in img_file_lower or 
                               "nomer" in img_file_lower or "номер" in img_file_lower)):
                        if 2 not in field_images:
                            field_images[2] = []
                        field_images[2].append(rel_path)
                        print(f"[process_zip_file] Added to field_images[2]: {img_file}")
                    elif (("gos" in img_file_lower or "гос" in img_file_lower or 
                           "nomer" in img_file_lower or "номер" in img_file_lower) and
                          not ("marka" in img_file_lower or "марка" in img_file_lower)):
                        if 3 not in field_images:
                            field_images[3] = []
                        field_images[3].append(rel_path)
                        print(f"[process_zip_file] Added to field_images[3]: {img_file}")
            
            print(f"[process_zip_file] Final field_images keys: {list(field_images.keys())}")
        
        row_data = {
            1: user_date_str,
            2: car_val,
            3: plate_val,
            4: fio_formatted,
            5: tn_ved_code,
            6: bnd_code,
            7: DataCleaner.clean_7(t1_data.get("Кол.тон (7)"), context),
            8: None, 9: None, 10: dollar_rate,
            11: None,
            12: None, 13: None,
            14: DataCleaner.clean_14(None, context),
            15: None, 16: None,
            15: None, 16: None,
            17: "",
            18: filtered_debug_str if not is_xlsx else "",
            'preview_images': preview_image_paths,
            'field_images': field_images,
            'sources': source_map,
            'errors': []
        }

        if not is_xlsx:
            brand_raw = t1_data.get("Марка", [])
            plate_raw = t1_data.get("Гос_номер ()", [])
            
            raw_details_parts = []
            if brand_raw:
                 raw_details_parts.append(f"Brand: {' '.join([t for t, h in brand_raw])}")
            if plate_raw:
                 plate_debug = " | ".join([f"{t} (H: {h})" for t, h in plate_raw])
                 raw_details_parts.append(f"Plate: {plate_debug}")
            
            row_data[17] = " | ".join(raw_details_parts)

        if surname_clean and surname_clean != "Unknown":
            surname_variants = normalize_surname(surname_clean)
            print(f" [MATCH DEBUG] Generated variants for '{surname_clean}': {surname_variants}")
            # print(f"[process_zip_file] Searching for ESF/SNT files with surname variants: {surname_variants}")
            
            for t2_path in type_2_files:
                fname = os.path.basename(t2_path).lower()
                
                match_found = False
                # 1. Exact match
                for variant in surname_variants:
                    if variant.lower() in fname:
                        match_found = True
                        print(f" [MATCH] Exact match found Type 2: {t2_path}")
                        break
                
                # 2. Fuzzy match
                if not match_found:
                    fname_words = re.findall(r'\w+', fname)
                    for variant in surname_variants:
                        v_lower = variant.lower()
                        for word in fname_words:
                            ratio = difflib.SequenceMatcher(None, v_lower, word).ratio()
                            if ratio > 0.80: # Threshold 80%
                                match_found = True
                                print(f" [MATCH] Fuzzy match ({ratio:.2f}) found Type 2: {t2_path} (variant: {variant}, word: {word})")
                                break
                        if match_found: break
                
                if match_found:
                    print(f" Match confirmed for Type 2: {t2_path}")
                    t2_preview_dir = os.path.join(preview_obj_dir, "type2")
                    os.makedirs(t2_preview_dir, exist_ok=True)
                    
                    t2_data = extract_text_from_pdf(t2_path, FIELDS_MAP_TYPE_2, t2_preview_dir)
                    
                    price_raw = t2_data.get("Цена (8)")
                    price_val_str = DataCleaner.clean_8(price_raw, context)
                    
                    check_price = safe_decimal(price_val_str, "Check Price")
                    
                    if check_price == Decimal("7") or check_price <= Decimal("1"):
                        print(f" [Price Check] Price is {check_price}, checking 2nd page of ESF...")
                        t2_data_p2 = extract_text_from_pdf(t2_path, FIELDS_MAP_TYPE_2_PAGE_2, t2_preview_dir, page_num=1)
                        price_alt_raw = t2_data_p2.get("Цена (8) Alt")
                        if price_alt_raw:
                            print(f" [Price Check] Found price on 2nd page: {price_alt_raw}")
                            t2_data["Цена (8)"] = price_alt_raw
                        else:
                            print(" [Price Check] No price found on 2nd page.")

                    row_data[8] = DataCleaner.clean_8(t2_data.get("Цена (8)"), context)
                    row_data[16] = DataCleaner.clean_16(t2_data.get("№ счет факт (Инвойс) (16)"), context)
                    
                    if os.path.exists(t2_preview_dir):
                        print(f"[process_zip_file] Scanning t2_preview_dir: {t2_preview_dir}")
                        for img_file in os.listdir(t2_preview_dir):
                            img_path = os.path.join(t2_preview_dir, img_file)
                            if os.path.isfile(img_path):
                                rel_path = os.path.relpath(img_path, settings.MEDIA_ROOT)
                                rel_path = rel_path.replace("\\", "/")
                                preview_image_paths.append(rel_path)
                                
                                img_file_lower = img_file.lower()
                                print(f"[process_zip_file] Checking t2 file: {img_file}")
                                
                                has_16 = ("_16" in img_file_lower or "(16)" in img_file or "16.png" in img_file_lower or
                                          "schet" in img_file_lower or "счет" in img_file_lower or 
                                          "fakt" in img_file_lower or "факт" in img_file_lower or
                                          "invoice" in img_file_lower or "инвойс" in img_file_lower or
                                          "invois" in img_file_lower)
                                
                                has_8 = ("_8" in img_file_lower or "(8)" in img_file or 
                                         "8.png" in img_file_lower or "_8_" in img_file_lower or 
                                         ("cena" in img_file_lower or "цена" in img_file_lower or 
                                          "price" in img_file_lower or "tsena" in img_file_lower))
                                
                                if has_16:
                                    if 16 not in field_images:
                                        field_images[16] = []
                                    field_images[16].append(rel_path)
                                    print(f"[process_zip_file] Added to field_images[16]: {img_file}")
                                elif has_8 or not has_16:
                                    if 8 not in field_images:
                                        field_images[8] = []
                                    field_images[8].append(rel_path)
                                    print(f"[process_zip_file] Added to field_images[8]: {img_file}")
                    
                    if save_photos:
                        os.makedirs(person_img_dir, exist_ok=True)
                        if os.path.exists(t2_preview_dir):
                            for img_file in os.listdir(t2_preview_dir):
                                src_path = os.path.join(t2_preview_dir, img_file)
                                dst_path = os.path.join(person_img_dir, f"type2_{img_file}")
                                shutil.copy2(src_path, dst_path)
                    
                    found_t2 = True
                    used_type_2.add(t2_path)
                    break 
                if found_t2: break
            
            if not found_t2:
                row_data['errors'].append("Не найден файл ЭСФ (Счет-фактура) для этого водителя.")
                print(f"[process_zip_file] Warning: no Type2 (ЭСФ) match for surname {surname_clean} in object {len(final_results)}")
            
            for t3_path in type_3_files:
                fname = os.path.basename(t3_path).lower()
                
                match_found = False
                # 1. Exact match
                for variant in surname_variants:
                    if variant.lower() in fname:
                        match_found = True
                        print(f" [MATCH] Exact match found Type 3: {t3_path}")
                        break
                
                # 2. Fuzzy match
                if not match_found:
                    fname_words = re.findall(r'\w+', fname)
                    for variant in surname_variants:
                        v_lower = variant.lower()
                        for word in fname_words:
                            ratio = difflib.SequenceMatcher(None, v_lower, word).ratio()
                            if ratio > 0.80: # Threshold 80%
                                match_found = True
                                print(f" [MATCH] Fuzzy match ({ratio:.2f}) found Type 3: {t3_path} (variant: {variant}, word: {word})")
                                break
                        if match_found: break

                if match_found:
                    print(f" Match confirmed for Type 3: {t3_path}")
                    t3_preview_dir = os.path.join(preview_obj_dir, "type3")
                    os.makedirs(t3_preview_dir, exist_ok=True)
                    
                    t3_data = extract_text_from_pdf(t3_path, FIELDS_MAP_TYPE_3, t3_preview_dir)
                    row_data[15] = DataCleaner.clean_15(t3_data.get("№ сопров.накл. KZ (15)"), context)
                    
                    date_sopr = DataCleaner.clean_1(t3_data.get("Дата сопр.накл (13)"), context)
                    row_data[13] = date_sopr
                        
                    if os.path.exists(t3_preview_dir):
                        print(f"[process_zip_file] Scanning t3_preview_dir: {t3_preview_dir}")
                        t3_files_list = []
                        for img_file in os.listdir(t3_preview_dir):
                            img_path = os.path.join(t3_preview_dir, img_file)
                            if os.path.isfile(img_path) and img_file.lower().endswith('.png'):
                                rel_path = os.path.relpath(img_path, settings.MEDIA_ROOT)
                                rel_path = rel_path.replace("\\", "/")
                                preview_image_paths.append(rel_path)
                                t3_files_list.append((img_file, rel_path))
                        
                        for img_file, rel_path in t3_files_list:
                            img_file_lower = img_file.lower()
                            print(f"[process_zip_file] Checking t3 file: {img_file}")
                            
                            has_kz_15 = ("kz" in img_file_lower or "_15" in img_file_lower or "(15)" in img_file or 
                                            "15.png" in img_file_lower or "_15_" in img_file_lower or
                                            "soprovozhdenie" in img_file_lower or "soprovozhd" in img_file_lower)
                            
                            has_13 = ("_13" in img_file_lower or "(13)" in img_file or 
                                        "13.png" in img_file_lower or "_13_" in img_file_lower)
                            has_date_keyword = ("date" in img_file_lower or "data" in img_file_lower or "дата" in img_file_lower or
                                                "datar" in img_file_lower) 
                            has_sopr_nakl = ("sopr" in img_file_lower or "сопров" in img_file_lower or "сопр" in img_file_lower or
                                                "nakl" in img_file_lower or "накл" in img_file_lower)
                            
                            if has_kz_15 and not has_13:
                                if 15 not in field_images:
                                    field_images[15] = []
                                field_images[15].append(rel_path)
                                print(f"[process_zip_file] Added to field_images[15]: {img_file}")
                            elif (has_13 or (has_date_keyword and has_sopr_nakl)) and not has_kz_15:
                                if 13 not in field_images:
                                    field_images[13] = []
                                field_images[13].append(rel_path)
                                print(f"[process_zip_file] Added to field_images[13]: {img_file}")
                            elif 13 not in field_images and 15 not in field_images:
                                file_index = [i for i, (f, _) in enumerate(t3_files_list) if f == img_file][0]
                                if file_index == 0:
                                    if 15 not in field_images:
                                        field_images[15] = []
                                    field_images[15].append(rel_path)
                                    print(f"[process_zip_file] Added to field_images[15] by order (first): {img_file}")
                                else:
                                    if 13 not in field_images:
                                        field_images[13] = []
                                    field_images[13].append(rel_path)
                                    print(f"[process_zip_file] Added to field_images[13] by order (second): {img_file}")
                            elif 15 in field_images and 13 not in field_images:
                                if 13 not in field_images:
                                    field_images[13] = []
                                field_images[13].append(rel_path)
                                print(f"[process_zip_file] Added to field_images[13] (15 already filled): {img_file}")
                            elif 13 in field_images and 15 not in field_images:
                                if 15 not in field_images:
                                    field_images[15] = []
                                field_images[15].append(rel_path)
                                print(f"[process_zip_file] Added to field_images[15] (13 already filled): {img_file}")
                    
                    if save_photos:
                        os.makedirs(person_img_dir, exist_ok=True)
                        if os.path.exists(t3_preview_dir):
                            for img_file in os.listdir(t3_preview_dir):
                                src_path = os.path.join(t3_preview_dir, img_file)
                                dst_path = os.path.join(person_img_dir, f"type3_{img_file}")
                                shutil.copy2(src_path, dst_path)
                    
                    found_t3 = True
                    used_type_3.add(t3_path)
                    break 
                if found_t3: break
            
            if not found_t3:
                row_data['errors'].append("Не найден файл СНТ (Накладная) для этого водителя.")
                print(f"[process_zip_file] Warning: no Type3 (СНТ) match for surname {surname_clean} in object {len(final_results)}")
        else:
            print(f"[process_zip_file] Surname not found or empty ('{surname_clean}'), skipping ESF/SNT matching.")

        try:
            kol_ton = safe_decimal(row_data[7], "Кол.тон (7)")
            kol_ton = kol_ton / Decimal("1000")
            row_data[7] = kol_ton
            
            cena = safe_decimal(row_data[8], "Цена (8)")
            row_data[8] = cena
            
            sum_dollar = (kol_ton * cena).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            row_data[9] = sum_dollar
            
            sum_som = (sum_dollar * dollar_rate).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            row_data[11] = sum_som
            
            nds_percent_value = nds_percent if isinstance(nds_percent, Decimal) else Decimal(str(nds_percent))
            nds_sum = (sum_som * nds_percent_value / Decimal("100")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            row_data[12] = nds_sum
            
            if found_t3 and not row_data.get(15):
                row_data['errors'].append("Не удалось найти '№ сопров.накл. KZ'. Проверьте файл СНТ.")
            if found_t2 and not row_data.get(16):
                row_data['errors'].append("Не удалось найти '№ счет факт'. Проверьте файл ЭСФ.")
            if found_t3 and not row_data.get(13):
                row_data['errors'].append("Не удалось найти 'Дата сопр.накл'. Проверьте файл СНТ.")
            
        except Exception as e:
            print(f"[process_zip_file] Calculation error for object {len(final_results)}: {e}")

        final_results.append(row_data)

    unused_t2 = len(type_2_files) - len(used_type_2)
    unused_t3 = len(type_3_files) - len(used_type_3)

    # --- Force Match Logic (1-1-1 Rule) ---
    incomplete_rows = []
    for i, res in enumerate(final_results):
        has_esf = res.get(16) is not None
        has_snt = res.get(15) is not None
        if not has_esf or not has_snt:
             incomplete_rows.append(i)

    if len(incomplete_rows) == 1 and unused_t2 == 1 and unused_t3 == 1:
        row_idx = incomplete_rows[0]
        row_data = final_results[row_idx]
        
        print(f"[Force Match] Triggered! 1 incomplete row (Index {row_idx}), 1 unused ESF, 1 unused SNT.")

        # Находим неиспользованные файлы
        leftover_t2 = list(set(type_2_files) - used_type_2)[0]
        leftover_t3 = list(set(type_3_files) - used_type_3)[0]
        
        print(f"[Force Match] Force linking ESF: {os.path.basename(leftover_t2)}")
        print(f"[Force Match] Force linking SNT: {os.path.basename(leftover_t3)}")
        
        # Reconstruct context
        t1_path_for_row = type_1_files[row_idx]
        fio_fmt = row_data.get(4, "")
        surname_for_ctx = fio_fmt.split()[0].strip() if fio_fmt else "Unknown"
        
        context = {
            'surname': surname_for_ctx,
            'zip_filename': zip_file.name,
            'type_2_files': type_2_files,
            'type_3_files': type_3_files,
            'filename': os.path.basename(t1_path_for_row)
        }

        # Directories
        target_preview_dir = os.path.join(preview_imgs_dir, f"obj_{row_idx}")
        t2_preview_dir = os.path.join(target_preview_dir, "type2")
        t3_preview_dir = os.path.join(target_preview_dir, "type3")
        os.makedirs(t2_preview_dir, exist_ok=True)
        os.makedirs(t3_preview_dir, exist_ok=True)

        # 1. Extract ESF (Type 2)
        t2_data = extract_text_from_pdf(leftover_t2, FIELDS_MAP_TYPE_2, t2_preview_dir)
        
        price_raw = t2_data.get("Цена (8)")
        price_val_str = DataCleaner.clean_8(price_raw, context)
        check_price = safe_decimal(price_val_str, "Check Price")
        
        if check_price == Decimal("7") or check_price <= Decimal("1"):
            print(f" [Force Match] checking 2nd page of ESF... (Price={check_price})")
            t2_data_p2 = extract_text_from_pdf(leftover_t2, FIELDS_MAP_TYPE_2_PAGE_2, t2_preview_dir, page_num=1)
            price_alt_raw = t2_data_p2.get("Цена (8) Alt")
            if price_alt_raw:
                t2_data["Цена (8)"] = price_alt_raw

        row_data[8] = DataCleaner.clean_8(t2_data.get("Цена (8)"), context)
        row_data[16] = DataCleaner.clean_16(t2_data.get("№ счет факт (Инвойс) (16)"), context)
        used_type_2.add(leftover_t2)

        # 2. Extract SNT (Type 3)
        t3_data = extract_text_from_pdf(leftover_t3, FIELDS_MAP_TYPE_3, t3_preview_dir)
        row_data[15] = DataCleaner.clean_15(t3_data.get("№ сопров.накл. KZ (15)"), context)
        date_sopr = DataCleaner.clean_1(t3_data.get("Дата сопр.накл (13)"), context)
        row_data[13] = date_sopr
        used_type_3.add(leftover_t3)
        
        # 3. Update Images in row_data
        def scan_and_add_images(scan_dir, r_data):
            if os.path.exists(scan_dir):
                for img_file in os.listdir(scan_dir):
                    img_path = os.path.join(scan_dir, img_file)
                    if os.path.isfile(img_path):
                        rel_path = os.path.relpath(img_path, settings.MEDIA_ROOT).replace("\\", "/")
                        r_data['preview_images'].append(rel_path)
                        
                        img_lower = img_file.lower()
                        has_16 = ("_16" in img_lower or "(16)" in img_lower or "schet" in img_lower or "invoice" in img_lower)
                        has_8 = ("_8" in img_lower or "(8)" in img_lower or "cena" in img_lower or "price" in img_lower)
                        if has_16:
                             if 16 not in r_data['field_images']: r_data['field_images'][16] = []
                             r_data['field_images'][16].append(rel_path)
                        elif has_8:
                             if 8 not in r_data['field_images']: r_data['field_images'][8] = []
                             r_data['field_images'][8].append(rel_path)
                        
                        has_kz_15 = ("kz" in img_lower or "_15" in img_lower or "(15)" in img_lower)
                        has_13 = ("_13" in img_lower or "(13)" in img_lower or "date" in img_lower)
                        if has_kz_15:
                             if 15 not in r_data['field_images']: r_data['field_images'][15] = []
                             r_data['field_images'][15].append(rel_path)
                        elif has_13:
                             if 13 not in r_data['field_images']: r_data['field_images'][13] = []
                             r_data['field_images'][13].append(rel_path)

        scan_and_add_images(t2_preview_dir, row_data)
        scan_and_add_images(t3_preview_dir, row_data)

        # 4. Save photos if needed
        if save_photos:
            sname = context['surname']
            r_date = row_data.get(1, "Unknown_Date")
            r_date_folder = str(r_date).replace("/", "-").replace("\\", "-")
            person_img_dir = os.path.join(imgs_root_dir, r_date_folder, sname)
            os.makedirs(person_img_dir, exist_ok=True)
            for d in [t2_preview_dir, t3_preview_dir]:
                if os.path.exists(d):
                    for img_file in os.listdir(d):
                         shutil.copy2(os.path.join(d, img_file), os.path.join(person_img_dir, img_file))

        # 5. Clear missing file errors
        new_errors = []
        for err in row_data['errors']:
             if "Не найден файл" in err or "Не удалось найти" in err:
                 continue
             new_errors.append(err)
        row_data['errors'] = new_errors
        
        # 6. Recalculate Totals
        try:
             # Ensure kol_ton is Decimal
             kt = row_data.get(7)
             if not isinstance(kt, Decimal):
                 kt = safe_decimal(kt, "Кол.тон (7)")
                 # Note: in loop, row_data[7] was result of / 1000 if successful.
                 # If it failed before, it might be string. 
                 # If it was successful, it is Decimal (tons).
                 # We assume if it was 20000kg -> 20t. 
                 # If it is > 500, likely it is still in kg?
                 if kt > 500: 
                     kt = kt / Decimal("1000")
                 row_data[7] = kt
             
             price = safe_decimal(row_data.get(8), "Цена (8)")
             row_data[8] = price
             
             sum_dollar = (kt * price).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
             row_data[9] = sum_dollar
             
             sum_som = (sum_dollar * dollar_rate).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
             row_data[11] = sum_som
             
             n_val = nds_percent if isinstance(nds_percent, Decimal) else Decimal(str(nds_percent))
             nds_sum = (sum_som * n_val / Decimal("100")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
             row_data[12] = nds_sum
             
        except Exception as e:
            print(f"[Force Match] Re-calculation error: {e}")
            row_data['errors'].append(f"Ошибка пересчета после Force Match: {e}")

        unused_t2 = 0
        unused_t3 = 0

    if unused_t2 > 0 or unused_t3 > 0:
        print(f"[process_zip_file] Unused files: unused_t2={unused_t2}, unused_t3={unused_t3}")
        
        # Build detailed diagnostics
        debug_msg = "\n--- ДЕТАЛИЗАЦИЯ ОБРАБОТКИ ---\n"
        for i, res in enumerate(final_results):
            surname = res.get(4, "Unknown")
            has_esf = "OK" if res.get(16) else "MISSING"
            has_snt = "OK" if res.get(15) or res.get(13) else "MISSING" # 15=KZ num, 13=Date
            debug_msg += f"#{i+1}: {surname} | ЭСФ: {has_esf} | СНТ: {has_snt}\n"
        
        unused_t2_files = sorted(list(set(type_2_files) - used_type_2))
        unused_t3_files = sorted(list(set(type_3_files) - used_type_3))
        
        if unused_t2_files:
            debug_msg += "\nНеиспользованные файлы ЭСФ:\n" + "\n".join([os.path.basename(f) for f in unused_t2_files])
        if unused_t3_files:
            debug_msg += "\nНеиспользованные файлы СНТ:\n" + "\n".join([os.path.basename(f) for f in unused_t3_files])

        print(debug_msg) # Ensure it goes to console logs

        error_msg = "Обнаружено несоответствие количества файлов:\n"
        if unused_t2 > 0:
            error_msg += f"- Лишних файлов ЭСФ (Счет-фактура): {unused_t2} шт.\n"
        if unused_t3 > 0:
            error_msg += f"- Лишних файлов СНТ (Накладная): {unused_t3} шт.\n"
        error_msg += "Убедитесь, что для каждого ЭСФ/СНТ есть соответствующий основной документ (PDF/XLSX).\n"
        error_msg += "\nСм. детали в логах ниже (какие файлы остались, какие водители обработаны):\n" + debug_msg
        
        if driver_debug_info:
             error_msg += "\n\n--- ДЕТАЛИЗАЦИЯ ПО ВОДИТЕЛЯМ (РАСПОЗНАВАНИЕ) ---\n" + "\n".join(driver_debug_info)

        raise Exception(error_msg)

    return final_results


def generate_excel(data, existing_excel_file=None, nds_percent=2):
    has_numbering_column = False
    next_row_number = 1
    
    if existing_excel_file:
        wb = openpyxl.load_workbook(existing_excel_file)
        ws = wb.active
        
        first_cell_value = ws.cell(row=1, column=1).value
        if first_cell_value:
            first_cell_str = str(first_cell_value).lower()
            if "дата" not in first_cell_str:
                has_numbering_column = True
                print(f"[generate_excel] Detected numbering column. First header: '{first_cell_value}'")
                
                max_number = 0
                for row_idx in range(2, ws.max_row + 1):
                    cell_value = ws.cell(row=row_idx, column=1).value
                    if cell_value is not None:
                        try:
                            num = int(cell_value)
                            if num > max_number:
                                max_number = num
                        except (ValueError, TypeError):
                            pass
                next_row_number = max_number + 1
                print(f"[generate_excel] Will continue numbering from {next_row_number}")
            else:
                print(f"[generate_excel] First column contains 'дата', no numbering column detected")
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "OCR Results"
        headers = [
            "Дата", "Марка АТС", "Гос.номер АТС", "ФИО Водит.", "Код ТН ВЭД",
            "БНД", "Кол.тон", "Цена", "Сумма в $", "Курс", "Сумма в сомах",
            "НДС ЕАЭС", "Дата сопр.накл", "Номер СМР", "№ сопров.накл. KZ", "№ счет факт"
        ]
        ws.append(headers)

    bold_font = Font(bold=True)
    header_row = 1
    max_col = ws.max_column
    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=header_row, column=col_idx)
        cell.font = bold_font

    fill_light_green = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    fill_dark_green = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")

    for row in data:
        if has_numbering_column:
            row_values = [
                next_row_number,
                row.get(1), row.get(2), row.get(3), row.get(4), row.get(5),
                row.get(6), row.get(7), row.get(8), None,
                row.get(10), None,
                None,
                row.get(13), row.get(14), row.get(15), row.get(16)
            ]
            next_row_number += 1
        else:
            row_values = [
                row.get(1), row.get(2), row.get(3), row.get(4), row.get(5),
                row.get(6), row.get(7), row.get(8), None,
                row.get(10), None,
                None,
                row.get(13), row.get(14), row.get(15), row.get(16)
            ]
        ws.append(row_values)
        
        current_row = ws.max_row
        
        col_offset = 1 if has_numbering_column else 0
        
        from openpyxl.utils import get_column_letter
        
        col_kol_ton = get_column_letter(7 + col_offset)
        col_price = get_column_letter(8 + col_offset)
        col_sum_dollar = get_column_letter(9 + col_offset)
        col_rate = get_column_letter(10 + col_offset)
        col_sum_som = get_column_letter(11 + col_offset)
        col_nds = get_column_letter(12 + col_offset)
        
        ws[f'{col_sum_dollar}{current_row}'] = f'={col_kol_ton}{current_row}*{col_price}{current_row}'
        
        ws[f'{col_sum_som}{current_row}'] = f'={col_sum_dollar}{current_row}*{col_rate}{current_row}'
        
        ws[f'{col_nds}{current_row}'] = f'={col_sum_som}{current_row}*{nds_percent}%'
        
        for col_idx in range(1, 19 + col_offset):
            cell = ws.cell(row=current_row, column=col_idx)
            
            val = cell.value
            if val:
                if has_numbering_column and col_idx == 1:
                    continue
                
                field_idx = col_idx - col_offset
                
                if field_idx in [1, 13] and isinstance(val, str):
                    try:
                        dt = datetime.strptime(val, "%d.%m.%Y").date()
                        cell.value = dt
                        cell.number_format = 'DD.MM.YYYY'
                    except ValueError:
                        pass
                    
                if field_idx in [5, 14] and isinstance(val, str) and val.isdigit():
                    try:
                        cell.value = int(val)
                        cell.number_format = '0'
                    except:
                        pass
                        
        ws.cell(row=current_row, column=1 + col_offset).fill = fill_light_green
        ws.cell(row=current_row, column=10 + col_offset).fill = fill_light_green
        
        ws.cell(row=current_row, column=7 + col_offset).fill = fill_dark_green

    col_offset = 1 if has_numbering_column else 0
    numeric_cols = [7 + col_offset, 8 + col_offset, 9 + col_offset, 10 + col_offset, 11 + col_offset, 12 + col_offset]
    
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for col_idx in numeric_cols:
            cell = row[col_idx - 1]
            cell.number_format = '#,##0.00'

    existing_tables = list(ws.tables.values())
    for existing_table in existing_tables:
        del ws.tables[existing_table.name]
    
    ws.auto_filter = None

    return wb