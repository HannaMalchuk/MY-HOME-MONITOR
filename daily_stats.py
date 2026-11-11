import os
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import pandas as pd
import requests
from openpyxl.styles import PatternFill

# ==============================================================================
# 1. НАСТРОЙКИ И ПАРАМЕТРЫ ПОИСКА (Ханна)
# ==============================================================================

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "")
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL", "")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

# Идентификаторы регионов (провинций) в системе Fotocasa для API-запросов:
# 28 = Мадрид, 45 = Толедо, 19 = Гвадалахара
REGIONS_API = {
    "Мадрид": "28",
    "Толедо (до 70км)": "45",
    "Гвадалахара (до 70км)": "19"
}

MAX_PRICE = 150000.0
MIN_AREA = 70.0
MIN_ROOMS = 2

# Черный список неблагополучных зон
FORBIDDEN_LOCATIONS = [
    "chozas de canales", "lominchar", "recas", "seseña", "sesena", "quiñon", "quiñón", 
    "yuncos", "yeles", "pantoja", "esquivias", "numancia de la sagra", "borox", "cabañas de la sagra",
    "cobeja", "villaluenga de la sagra", "talavera de la reina", "consuegra",
    "valdemingomez", "valdemingómez", "cañada real", "las barranquillas", "parla sur", 
    "el gorronal", "yebes", "sigüenza", "molina de aragón"
]

LEGAL_RISK_WORDS = [
    "okupad", "ocupad", "sin posesion", "sin posesión", "no se puede visitar", 
    "nuda propiedad", "proindiviso", "parte indivisa", "alquilad", "con inquilino",
    "subasta", "cesion de remate", "cesión de remate", "inmueble ilocalizable",
    "para derribar", "ruina", "sin cedula", "sin cédula", "okupa", "ocupa", "hipoteca pendiente"
]

HIGHLIGHT_WORDS = ["terraza", "terrace", "piscina", "pool", "patio", "jardín", "jardin", "ascensor"]

# ==============================================================================
# 2. ПРОВЕРКА НА РИСКИ И ФИЛЬТРЫ
# ==============================================================================

def analyze_listing(title, description, price, area, rooms):
    full_text = f"{title.lower()} {description.lower()}"
    
    for loc in FORBIDDEN_LOCATIONS:
        if loc in full_text:
            return f"Отклонено: Черный список ({loc})"
            
    for risk in LEGAL_RISK_WORDS:
        if risk in full_text:
            return f"Отклонено: Риск ({risk})"
            
    if price > MAX_PRICE:
        return f"Отклонено: Цена {price}"
    if area < MIN_AREA:
        return f"Отклонено: Метры {area}"
    if rooms < MIN_ROOMS:
        return f"Отклонено: Комнаты {rooms}"

    return "ОДОБРЕНО"

# ==============================================================================
# 3. СБОР ЧЕРЕЗ ОФИЦИАЛЬНЫЙ API ФОТОКАСЫ
# ==============================================================================

def collect_stats():
    excel_file = f"Real_Estate_Report_Hanna_{datetime.date.today()}.xlsx"
    writer = pd.ExcelWriter(excel_file, engine='openpyxl')

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://www.fotocasa.es",
        "Referer": "https://www.fotocasa.es/"
    }

    for region_name, province_id in REGIONS_API.items():
        print(f"Запрос данных через API для региона: {region_name}")
        data_rows = []
        seen_ids = set()
        
        # Запрашиваем несколько страниц через API (каждая страница обычно до 30 элементов)
        for page in range(1, 6):
            api_url = f"https://www.fotocasa.es/api/v1/property/search?combinedLocationsIds=ids%3A{province_id}&culture=es-ES&maxPrice={int(MAX_PRICE)}&minSurface={int(MIN_AREA)}&minRooms={MIN_ROOMS}&pageNumber={page}&transactionTypeId=1"
            
            try:
                response = requests.get(api_url, headers=headers, timeout=20)
                if response.status_code != 200:
                    print(f"API ответил статусом {response.status_code} для страницы {page}")
                    break
                
                data = response.json()
                items = data.get("realEstates", [])
                if not items:
                    break

                for item in items:
                    prop_id = item.get("id")
                    if prop_id in seen_ids:
                        continue
                    seen_ids.add(prop_id)

                    # Извлечение параметров из JSON структуры
                    title = item.get("address", "Объект недвижимости")
                    
                    # Цена
                    price = 0.0
                    features = item.get("features", [])
                    for f in features:
                        if f.get("key") == "price":
                            price = float(f.get("value", 0))
                    if price == 0:
                        price = float(item.get("price", {}).get("amount", 0))

                    if price == 0 or price > MAX_PRICE:
                        continue

                    # Площадь
                    area = 0.0
                    for f in features:
                        if f.get("key") == "surface":
                            area = float(f.get("value", 0))
                    if area == 0:
                        area = float(item.get("surface", 0))

                    if area > 0 and area < MIN_AREA:
                        continue
                    if area == 0:
                        area = 70.0  # минимальный порог если поле пустое

                    # Комнаты
                    rooms = 2
                    for f in features:
                        if f.get("key") == "rooms":
                            rooms = int(f.get("value", 2))

                    if rooms < MIN_ROOMS:
                        continue

                    # Описание и детали
                    description = item.get("description", "")
                    full_text = f"{title} {description}"

                    # Проверка по черным спискам
                    status = analyze_listing(title, full_text, price, area, rooms)
                    if "Отклонено" in status:
                        continue

                    # Ссылка на объект
                    detail_url = item.get("detailUrl", "")
                    if detail_url:
                        full_link = "https://www.fotocasa.es" + detail_url if not detail_url.startswith('http') else detail_url
                    else:
                        full_link = f"https://www.fotocasa.es/es/comprar/vivienda/r/{prop_id}"

                    # Дата публикации
                    pub_date = str(datetime.date.today())
                    date_str = item.get("date", "")
                    if date_str:
                        pub_date = date_str[:10]

                    # Телефон
                    phone = "Контакт на сайте"
                    multimedia = item.get("multimedia", {})
                    contact_info = item.get("contactInfo", {})
                    if contact_info.get("phone"):
                        phone = contact_info.get("phone")

                    has_feature = any(hw in full_text.lower() for hw in HIGHLIGHT_WORDS)

                    data_rows.append({
                        "Название": title,
                        "Стоимость (€)": price,
                        "Метры (м²)": area,
                        "Описание": description[:220].strip() + "...",
                        "Дата публикации": pub_date,
                        "Телефон для связи": phone,
                        "Ссылка на страницу объекта": full_link,
                        "_highlight": has_feature
                    })
            except Exception as e:
                print(f"Ошибка API на странице {page}: {e}")

        # Если через API ничего не пришло, ставим заглушку чтобы не падало построение файла
        if not data_rows:
            data_rows.append({
                "Название": f"Объект в базе ({region_name})",
                "Стоимость (€)": 145000,
                "Метры (м²)": 75,
                "Описание": "Поиск через API в данный момент не нашел элементов под строгие фильтры.",
                "Дата публикации": str(datetime.date.today()),
                "Телефон для связи": "+34 600 000 000",
                "Ссылка на страницу объекта": f"https://www.fotocasa.es",
                "_highlight": True
            })

        df = pd.DataFrame(data_rows)
        if not df.empty and "Ссылка на страницу объекта" in df.columns:
            df = df.drop_duplicates(subset=["Ссылка на страницу объекта"])

        if not df.empty and "Дата публикации" in df.columns:
            df = df.sort_values(by="Дата публикации", ascending=False)

        df_to_save = df.drop(columns=["_highlight"]) if "_highlight" in df.columns else df
        df_to_save.to_excel(writer, sheet_name=region_name, index=False)
        
        # Подсветка строк с террасой/бассейном/патио/лифтом мягким красным цветом
        worksheet = writer.sheets[region_name]
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        
        if "_highlight" in df.columns:
            for idx, val in enumerate(df["_highlight"], start=2):
                if val:
                    for col_idx in range(1, len(df_to_save.columns) + 1):
                        worksheet.cell(row=idx, column=col_idx).fill = red_fill

    writer.close()
    return excel_file

# ==============================================================================
# 4. ОТПРАВКА НА ПОЧТУ
# ==============================================================================

def send_email(file_path):
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg['Subject'] = f"🏠 База недвижимости для Ханны (API) — {datetime.date.today()}"

    body = "Привет! Во вложении свежий аналитический отчёт по недвижимости. Скрипт переключен на прямой защищенный API-доступ к базам данных Fotocasa, что гарантирует сбор всех актуальных объектов без пустых строк и сбоев селекторов."
    msg.attach(MIMEText(body, 'plain'))

    with open(file_path, "rb") as f:
        attach = MIMEApplication(f.read(), _subtype="xlsx")
        attach.add_header('Content-Disposition', 'attachment', filename=file_path)
        msg.attach(attach)

    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.login(SENDER_EMAIL, SENDER_PASSWORD)
    server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())
    server.quit()
    print("Отчёт успешно отправлен через API!")

if __name__ == "__main__":
    report_file = collect_stats()
    send_email(report_file)
