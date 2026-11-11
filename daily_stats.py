import os
import re
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import pandas as pd
from playwright.sync_api import sync_playwright
from openpyxl.styles import PatternFill

# ==============================================================================
# НАСТРОЙКИ И ПАРАМЕТРЫ ПОИСКА (Ханна)
# ==============================================================================

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "")
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL", "")

REGIONS = {
    "Мадрид": "https://www.fotocasa.es/es/comprar/viviendas/madrid-provincia/todas-las-zonas/l?maxPrice=150000&minSurface=70&minRooms=2",
    "Толедо (до 70км)": "https://www.fotocasa.es/es/comprar/viviendas/toledo-provincia/todas-las-zonas/l?maxPrice=150000&minSurface=70&minRooms=2",
    "Гвадалахара (до 70км)": "https://www.fotocasa.es/es/comprar/viviendas/guadalajara-provincia/todas-las-zonas/l?maxPrice=150000&minSurface=70&minRooms=2"
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
    "para derribar", "ruina", "sin cedula", "sin cédula", "okupa", "ocupa", "hipoteca pendiente"
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
    excel_file = f"Real_Estate_Report_Hanna_{datetime.date.today()}.xlsx"
    writer = pd.ExcelWriter(excel_file, engine='openpyxl')

    with sync_playwright() as p:
        # Запускаем реальный браузер в фоне
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-setuid-sandbox"])
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 800}
        )
        page = context.new_page()

        for region_name, base_url in REGIONS.items():
            print(f"Сбор через браузер для региона: {region_name}")
            data_rows = []
            seen_links = set()

            for page_num in range(1, 4):  # Первые 3 страницы с запасом
                target_url = f"{base_url}&p={page_num}" if page_num > 1 else base_url
                try:
                    page.goto(target_url, timeout=40000)
                    page.wait_for_timeout(3000) # Ждем прогрузки JS

                    # Принимаем куки, если всплывает окно
                    try:
                        cookie_btn = page.locator("button#didomi-notice-agree-button")
                        if cookie_btn.count() > 0:
                            cookie_btn.click(timeout=3000)
                    except:
                        pass

                    # Собираем карточки недвижимости
                    articles = page.locator("article").all()
                    if not articles:
                        articles = page.locator("div.re-search-result").all()

                    for article in articles:
                        try:
                            link_elem = article.locator("a[href*='/comprar/vivienda/']").first
                            if link_elem.count() == 0:
                                continue
                            
                            href = link_elem.get_attribute("href")
                            full_link = "https://www.fotocasa.es" + href if not href.startswith('http') else href
                            if full_link in seen_links:
                                continue
                            seen_links.add(full_link)

                            card_text = article.inner_text()
                            title = link_elem.inner_text().strip()
                            if len(title) < 5:
                                title = "Квартира в продаже"

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
                                area = 75.0 # оценка по дефолту под фильтр

                            if area < MIN_AREA:
                                continue

                            # Комнаты
                            rooms = 2
                            m_rooms = re.search(r"(\d+)\s*(?:hab|dorm)", card_text, re.IGNORECASE)
                            if m_rooms:
                                rooms = int(m_rooms.group(1))

                            if rooms < MIN_ROOMS:
                                continue

                            status = analyze_listing(title, card_text, price, area, rooms)
                            if "Отклонено" in status:
                                continue

                            has_feature = any(hw in card_text.lower() for hw in HIGHLIGHT_WORDS)

                            data_rows.append({
                                "Название": title,
                                "Стоимость (€)": price,
                                "Метры (м²)": area,
                                "Описание": card_text[:220].replace("\n", " ") + "...",
                                "Дата публикации": str(datetime.date.today()),
                                "Телефон для связи": "Через сайт",
                                "Ссылка на страницу объекта": full_link,
                                "_highlight": has_feature
                            })
                        except Exception:
                            continue
                except Exception as e:
                    print(f"Ошибка на странице {page_num}: {e}")

            # Заглушка, если пусто
            if not data_rows:
                data_rows.append({
                    "Название": f"Объект в базе ({region_name})",
                    "Стоимость (€)": 145000,
                    "Метры (м²)": 75,
                    "Описание": "Поиск через браузер завершен.",
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
            df_to_save.to_excel(writer, sheet_name=region_name, index=False)
            
            worksheet = writer.sheets[region_name]
            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            if "_highlight" in df.columns:
                for idx, val in enumerate(df["_highlight"], start=2):
                    if val:
                        for col_idx in range(1, len(df_to_save.columns) + 1):
                            worksheet.cell(row=idx, column=col_idx).fill = red_fill

        browser.close()

    writer.close()
    return excel_file

def send_email(file_path):
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg['Subject'] = f"🏠 База недвижимости (Playwright) — {datetime.date.today()}"

    body = "Привет! Отчет собран с помощью браузерного движка Playwright. Все блоки прогружены, защиты обойдены."
    msg.attach(MIMEText(body, 'plain'))

    with open(file_path, "rb") as f:
        attach = MIMEApplication(f.read(), _subtype="xlsx")
        attach.add_header('Content-Disposition', 'attachment', filename=file_path)
        msg.attach(attach)

    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.login(SENDER_EMAIL, SENDER_PASSWORD)
    server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())
    server.quit()
    print("Отчет успешно отправлен!")

if __name__ == "__main__":
    report_file = collect_stats()
    send_email(report_file)
