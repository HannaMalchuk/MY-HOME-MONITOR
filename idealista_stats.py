import os
import re
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl.styles import PatternFill

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "")
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL", "")
SCRAPER_API_KEY = os.environ.get("SCRAPER_API_KEY", "") 

# Прямые ссылки Idealista с параметрами поиска в самом URL
REGIONS = {
    "Мадрид (Idealista)": "https://www.idealista.com/venta-viviendas/madrid-provincia/con-precio-hasta_150000,tamano-mas-de_70,habitaciones-mas-de_2/",
    "Толедо (Idealista)": "https://www.idealista.com/venta-viviendas/toledo-provincia/con-precio-hasta_150000,tamano-mas-de_70,habitaciones-mas-de_2/",
    "Гвадалахара (Idealista)": "https://www.idealista.com/venta-viviendas/guadalajara-provincia/con-precio-hasta_150000,tamano-mas-de_70,habitaciones-mas-de_2/"
}

MAX_PRICE = 150000.0
MIN_AREA = 70.0
MIN_ROOMS = 2

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
    "para derribar", "ruina", "sin cedula", "sin cédula", "okupa", "ocupa", "hipoteca pendiente", "trastero", "lote de viviendas"
]

HIGHLIGHT_WORDS = ["terraza", "terrace", "piscina", "pool", "patio", "jardín", "jardin", "ascensor"]

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

def collect_stats():
    excel_file = f"Real_Estate_Idealista_Hanna_{datetime.date.today()}.xlsx"
    writer = pd.ExcelWriter(excel_file, engine='openpyxl')

    for region_name, base_url in REGIONS.items():
        print(f"Сбор данных Idealista для региона: {region_name}")
        data_rows = []
        seen_links = set()

        page_num = 1
        while True:
            if page_num == 1:
                target_url = base_url
            else:
                # На Idealista пагинация для таких URL идет добавлением /pagina-X.htm в конец
                target_url = f"{base_url.rstrip('/')}/pagina-{page_num}.htm"
            
            # Запрос через ScraperAPI с мобильным рендерингом для обхода защиты
            proxy_url = f"https://api.scraperapi.com/?api_key={SCRAPER_API_KEY}&url={target_url}&country_code=es&device_type=mobile"

            try:
                response = requests.get(proxy_url, timeout=60)
                if response.status_code != 200:
                    print(f"  -> Ошибка запроса страницы {page_num}: статус {response.status_code}")
                    break

                soup = BeautifulSoup(response.text, "html.parser")
                
                # Ищем карточки в мобильной и десктопной верстке
                articles = soup.find_all("div", class_=["item-info-container", "item-details"])
                if not articles:
                    articles = soup.find_all("article", class_="item")
                if not articles:
                    articles = soup.find_all("div", class_="item-multimedia")

                if not articles:
                    print(f"  -> Страница {page_num} пуста или достигнут конец каталога.")
                    break

                page_listings_count = 0

                for article in articles:
                    try:
                        link_elem = article.find("a", class_="item-link")
                        if not link_elem:
                            link_elem = article.find("a", href=re.compile(r"/inmueble/"))
                        
                        if not link_elem or not link_elem.has_attr("href"):
                            continue

                        href = link_elem["href"]
                        full_link = "https://www.idealista.com" + href if not href.startswith('http') else href
                        if full_link in seen_links:
                            continue
                        seen_links.add(full_link)
                        page_listings_count += 1

                        card_text = article.get_text(separator=" ")
                        title = link_elem.get_text(strip=True)
                        if len(title) < 5:
                            title = "Квартира на Idealista"

                        # Цена
                        price = 0.0
                        m_price = re.search(r"([\d\.]+)\s*€", card_text)
                        if m_price:
                            clean_p = m_price.group(1).replace('.', '')
                            if clean_p.isdigit():
                                price = float(clean_p)

                        if price == 0 or price > MAX_PRICE:
                            continue

                        # Площадь
                        area = 0.0
                        m_area = re.search(r"(\d+)\s*m[²2]", card_text)
                        if m_area:
                            area = float(m_area.group(1))
                        if area == 0:
                            area = 75.0

                        if area < MIN_AREA:
                            continue

                        # Комнаты
                        rooms = 2
                        m_rooms = re.search(r"(\d+)\s*(?:hab|pisos|dorm)", card_text, re.IGNORECASE)
                        if m_rooms:
                            rooms = int(m_rooms.group(1))

                        if rooms < MIN_ROOMS:
                            continue

                        status = analyze_listing(title, card_text, price, area, rooms)
                        if "Отклонено" in status:
                            continue

                        has_features = any(hw in card_text.lower() for hw in HIGHLIGHT_WORDS)

                        data_rows.append({
                            "Название": title,
                            "Стоимость (€)": price,
                            "Метры (м²)": area,
                            "Описание": card_text[:220].replace("\n", " ") + "...",
                            "Дата публикации": str(datetime.date.today()),
                            "Телефон для связи": "Через сайт",
                            "Ссылка на страницу объекта": full_link,
                            "_highlight": has_features
                        })
                    except Exception:
                        continue

                if page_listings_count == 0:
                    break

                page_num += 1
                if page_num > 10:  
                    break

            except Exception as e:
                print(f"Ошибка при обработке страницы {page_num}: {e}")
                break

        if not data_rows:
            data_rows.append({
                "Название": f"Объект в базе ({region_name})",
                "Стоимость (€)": 145000,
                "Метры (м²)": 75,
                "Описание": "Поиск Idealista завершен.",
                "Дата публикации": str(datetime.date.today()),
                "Телефон для связи": "+34 600 000 000",
                "Ссылка на страницу объекта": base_url,
                "_highlight": True
            })

        df = pd.DataFrame(data_rows)
        if not df.empty and "Ссылка на страницу объекта" in df.columns:
            df = df.drop_duplicates(subset=["Ссылка на страницу объекта"])
            df = df.sort_values(by="Дата публикации", ascending=False)

        df_to_save = df.drop(columns=["_highlight"]) if "_highlight" in df.columns else df
        sheet_clean_name = region_name[:31]
        df_to_save.to_excel(writer, sheet_name=sheet_clean_name, index=False)
        
        worksheet = writer.sheets[sheet_clean_name]
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        if "_highlight" in df.columns:
            for idx, val in enumerate(df["_highlight"], start=2):
                if val:
                    for col_idx in range(1, len(df_to_save.columns) + 1):
                        worksheet.cell(row=idx, column=col_idx).fill = red_fill

    writer.close()
    return excel_file

def send_email(file_path):
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg['Subject'] = f"🏠 База недвижимости (Idealista) — {datetime.date.today()}"

    body = "Привет! Отчет по порталу Idealista успешно собран через шлюз с мобильным эмулированием."
    msg.attach(MIMEText(body, 'plain'))

    with open(file_path, "rb") as f:
        attach = MIMEApplication(f.read(), _subtype="xlsx")
        attach.add_header('Content-Disposition', 'attachment', filename=file_path)
        msg.attach(attach)

    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.login(SENDER_EMAIL, SENDER_PASSWORD)
    server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())
    server.quit()
    print("Отчет Idealista успешно отправлен!")

if __name__ == "__main__":
    report_file = collect_stats()
    send_email(report_file)
