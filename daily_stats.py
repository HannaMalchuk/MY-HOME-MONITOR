import os
import re
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import PatternFill

# ==============================================================================
# 1. НАСТРОЙКИ И ПАРАМЕТРЫ ПОИСКА (Ханна)
# ==============================================================================

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "")
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL", "")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

# Ссылки на поиск Fotocasa (цена до 150к, от 70м², от 2 комнат)
REGIONS = {
    "Мадрид": "https://www.fotocasa.es/es/comprar/viviendas/madrid-provincia/todas-las-zonas/l?maxPrice=150000&minSurface=70&minRooms=2",
    "Толедо (до 70км)": "https://www.fotocasa.es/es/comprar/viviendas/toledo-provincia/todas-las-zonas/l?maxPrice=150000&minSurface=70&minRooms=2",
    "Гвадалахара (до 70км)": "https://www.fotocasa.es/es/comprar/viviendas/guadalajara-provincia/todas-las-zonas/l?maxPrice=150000&minSurface=70&minRooms=2"
}

MAX_PRICE = 150000.0
MIN_AREA = 70.0
MIN_ROOMS = 2

# Скорректированный черный список (исключены пересечения с названиями улиц)
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

# Ключевые слова для подсветки красным цветом
HIGHLIGHT_WORDS = ["terraza", "terrace", "piscina", "pool", "patio", "jardín", "jardin", "ascensor"]

# ==============================================================================
# 2. ФИЛЬТРАЦИЯ И ПРОВЕРКА ОБЪЕКТОВ
# ==============================================================================

def analyze_listing(title, description, price, area, rooms):
    full_text = f"{title.lower()} {description.lower()}"
    
    for loc in FORBIDDEN_LOCATIONS:
        if loc in full_text:
            return f"Отклонено: Черный список зоны ({loc.title()})"
            
    for risk in LEGAL_RISK_WORDS:
        if risk in full_text:
            return f"Отклонено: Юридический риск / Обременение ({risk})"
            
    if price > MAX_PRICE:
        return f"Отклонено: Цена выше лимита ({price:,.0f} €)"
    if area < MIN_AREA:
        return f"Отклонено: Маленькая площадь ({area:.0f} m² < 70м²)"
    if rooms < MIN_ROOMS:
        return f"Отклонено: Мало комнат ({rooms} < 2)"

    return "ОДОБРЕНО"

# ==============================================================================
# 3. СБОР ВСЕЙ БАЗЫ И СОРТИРОВКА
# ==============================================================================

def collect_stats():
    excel_file = f"Real_Estate_Report_Hanna_{datetime.date.today()}.xlsx"
    writer = pd.ExcelWriter(excel_file, engine='openpyxl')

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Language": "es-ES,es;q=0.9"
    }

    for region_name, base_url in REGIONS.items():
        print(f"Сбор данных для региона: {region_name}")
        data_rows = []
        seen_links = set()
        
        # Увеличиваем сбор до 10 страниц для максимального охвата базы
        for page_num in range(1, 11):
            target_url = f"{base_url}&p={page_num}" if page_num > 1 else base_url
            try:
                response = requests.get(target_url, headers=headers, timeout=20)
                if response.status_code != 200:
                    continue
                
                soup = BeautifulSoup(response.text, 'html.parser')
                links = soup.find_all('a', href=True)

                for link_tag in links:
                    href = link_tag.get('href')
                    if not href or '/comprar/vivienda/' not in href:
                        continue
                        
                    full_link = "https://www.fotocasa.es" + href if not href.startswith('http') else href
                    if full_link in seen_links:
                        continue
                    seen_links.add(full_link)

                    card = link_tag.find_parent('article') or link_tag.find_parent('div', class_=lambda x: x and 'card' in x.lower())
                    card_text = card.get_text() if card else link_tag.get_text()

                    title = link_tag.get_text(strip=True)
                    if len(title) < 5:
                        continue

                    # Цена
                    price = 0.0
                    m_price = re.search(r"([\d\.]+)\s*€", card_text)
                    if m_price:
                        price_str = m_price.group(1).replace('.', '')
                        if price_str.isdigit():
                            price = float(price_str)

                    if price == 0 or price > MAX_PRICE:
                        continue

                    # Площадь
                    area = 0.0
                    m_area = re.search(r"(\d+)\s*m[²2]", card_text)
                    if m_area:
                        area = float(m_area.group(1))

                    # Комнаты
                    rooms = 0
                    m_rooms = re.search(r"(\d+)\s*(?:hab|dorm)", card_text, re.IGNORECASE)
                    if m_rooms:
                        rooms = int(m_rooms.group(1))

                    if area < MIN_AREA or rooms < MIN_ROOMS:
                        continue

                    # Проверка по критериям
                    status = analyze_listing(title, card_text, price, area, rooms)
                    if "Отклонено" in status:
                        continue

                    pub_date = str(datetime.date.today())
                    if "hace" in card_text.lower() or "ayer" in card_text.lower():
                        pub_date = str(datetime.date.today() - datetime.timedelta(days=1))
                    
                    phone = "Контакт на сайте"
                    m_phone = re.search(r"(\+34[\s\d]{9,})", card_text)
                    if m_phone:
                        phone = m_phone.group(1).strip()

                    has_feature = any(hw in card_text.lower() for hw in HIGHLIGHT_WORDS)

                    data_rows.append({
                        "Название": title,
                        "Стоимость (€)": price,
                        "Метры (м²)": area,
                        "Описание": card_text[:200] + "...",
                        "Дата публикации": pub_date,
                        "Телефон для связи": phone,
                        "Ссылка на страницу объекта": full_link,
                        "_highlight": has_feature
                    })
            except Exception as e:
                print(f"Ошибка на странице {page_num}: {e}")

        df = pd.DataFrame(data_rows)
        if not df.empty and "Ссылка на страницу объекта" in df.columns:
            df = df.drop_duplicates(subset=["Ссылка на страницу объекта"])

        if not df.empty and "Дата публикации" in df.columns:
            df = df.sort_values(by="Дата публикации", ascending=False)

        df_to_save = df.drop(columns=["_highlight"]) if "_highlight" in df.columns else df
        df_to_save.to_excel(writer, sheet_name=region_name, index=False)
        
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
    msg['Subject'] = f"🏠 Полная база недвижимости для Ханны — {datetime.date.today()}"

    body = "Привет! Во вложении твой полный аналитический отчёт по недвижимости (Мадрид, Толедо до 70км, Гвадалахара до 70км).\n\nЧерные списки скорректированы, глубина сбора увеличена до 10 страниц. Объекты с террасой, бассейном, патио или лифтом подсвечены красным."
    msg.attach(MIMEText(body, 'plain'))

    with open(file_path, "rb") as f:
        attach = MIMEApplication(f.read(), _subtype="xlsx")
        attach.add_header('Content-Disposition', 'attachment', filename=file_path)
        msg.attach(attach)

    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.login(SENDER_EMAIL, SENDER_PASSWORD)
    server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())
    server.quit()
    print("Отчёт успешно отправлен!")

if __name__ == "__main__":
    report_file = collect_stats()
    send_email(report_file)
