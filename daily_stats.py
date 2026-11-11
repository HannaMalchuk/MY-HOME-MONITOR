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

# ==============================================================================
# 1. НАСТРОЙКИ И ПАРАМЕТРЫ
# ==============================================================================

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "")
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL", "")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

# Ссылки на поиск Fotocasa по регионам (цена до 145k, от 60м², от 2 комнат)
REGIONS = {
    "Мадрид": "https://www.fotocasa.es/es/comprar/viviendas/madrid-provincia/todas-las-zonas/l?maxPrice=145000&minSurface=60&minRooms=2",
    "Толедо": "https://www.fotocasa.es/es/comprar/viviendas/toledo-provincia/todas-las-zonas/l?maxPrice=145000&minSurface=60&minRooms=2",
    "Гвадалахара": "https://www.fotocasa.es/es/comprar/viviendas/guadalajara-provincia/todas-las-zonas/l?maxPrice=145000&minSurface=60&minRooms=2"
}

MAX_PRICE = 145000.0
MIN_AREA = 60.0
MIN_ROOMS = 2

FORBIDDEN_LOCATIONS = [
    "chozas de canales", "chozas", "lominchar", "recas", "seseña", "sesena", "quiñon", "quiñón", 
    "yuncos", "yeles", "pantoja", "esquivias", "numancia de la sagra", "borox", "cabañas de la sagra",
    "cobeja", "magan", "magán", "villaluenga de la sagra",
    "valdemingomez", "valdemingómez", "cañada real", "las barranquillas", "parla", "alhondiga", "alhóndiga",
    "el gorronal", "gorronal", "azuqueca de henares", "marchamalo", "yebes"
]

LEGAL_RISK_WORDS = [
    "okupad", "ocupad", "sin posesion", "sin posesión", "no se puede visitar", 
    "nuda propiedad", "proindiviso", "parte indivisa", "alquilad", "con inquilino",
    "subasta", "cesion de remate", "cesión de remate", "inmueble ilocalizable",
    "para derribar", "ruina", "sin cedula", "sin cédula", "okupa", "ocupa"
]

# ==============================================================================
# 2. АНАЛИЗАТОР РИСКОВ И ФИЛЬТРАЦИЯ
# ==============================================================================

def analyze_listing(title, description, price, area, rooms):
    full_text = f"{title.lower()} {description.lower()}"
    
    for loc in FORBIDDEN_LOCATIONS:
        if loc in full_text:
            return f"Отклонено: Опасная/неблагополучная зона ({loc.title()})"
            
    for risk in LEGAL_RISK_WORDS:
        if risk in full_text:
            return f"Отклонено: Юридический риск / Okupas ({risk})"
            
    if price > MAX_PRICE:
        return f"Отклонено: Превышена цена ({price:,.0f} € > {MAX_PRICE:,.0f} €)"
    if area < MIN_AREA:
        return f"Отклонено: Малая площадь ({area:.0f} m² < {MIN_AREA:.0f} m²)"
    if rooms < MIN_ROOMS:
        return f"Отклонено: Мало комнат ({rooms} < {MIN_ROOMS})"

    if GEMINI_API_KEY:
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}"
            prompt = f"Проанализируй объявление: '{title}'. Описание: '{description}'. Подходит ли для безопасного проживания семьи с ребенком? Ответь JSON: {{\"safe\": true/false, \"reason\": \"краткая причина\"}}"
            resp = requests.post(url, json={"contents": [{"parts": [{"text": prompt}]}]}, timeout=5)
            if resp.status_code == 200:
                text = resp.json()['candidates'][0]['content']['parts'][0]['text']
                if '"safe": false' in text.lower():
                    return "Отклонено AI: Небезопасный район или сомнительные условия"
        except Exception:
            pass

    return "ОДОБРЕНО (Соответствует параметрам)"

# ==============================================================================
# 3. СБОР ДАННЫХ И ФОРМИРОВАНИЕ ОТЧЕТА
# ==============================================================================

def collect_stats():
    excel_file = f"Real_Estate_Daily_Report_{datetime.date.today()}.xlsx"
    writer = pd.ExcelWriter(excel_file, engine='openpyxl')

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Language": "es-ES,es;q=0.9"
    }

    for region_name, target_url in REGIONS.items():
        print(f"Обработка региона: {region_name}")
        data_rows = []
        
        try:
            response = requests.get(target_url, headers=headers, timeout=20)
            print(f"Статус ответа для {region_name}: {response.status_code}")
            
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                
                # Ищем блоки карточек по универсальным классам Fotocasa
                cards = soup.find_all(['article', 'div'], class_=lambda x: x and ('re-Card' in x or 'card' in x.lower()))
                print(f"Найдено элементов карточек: {len(cards)}")

                for card in cards:
                    try:
                        title_el = card.find(['a', 'span', 'h3'], class_=lambda x: x and ('title' in x.lower() or 'link' in x.lower()))
                        if not title_el:
                            title_el = card.find('a')
                        
                        if not title_el:
                            continue
                            
                        title = title_el.get_text(strip=True)
                        if len(title) < 5:
                            continue

                        link = title_el.get('href') if title_el.name == 'a' else None
                        if not link:
                            parent_a = card.find('a', href=True)
                            link = parent_a.get('href') if parent_a else target_url

                        if link and not link.startswith("http"):
                            link = "https://www.fotocasa.es" + link

                        price_el = card.find(class_=lambda x: x and 'price' in x.lower())
                        price_text = price_el.get_text(strip=True) if price_el else "0"
                        price = float(re.sub(r"[^\d]", "", price_text)) if re.sub(r"[^\d]", "", price_text) else 0.0

                        if price == 0 or price > MAX_PRICE:
                            continue

                        text_content = card.get_text()
                        area, rooms = 0.0, 0
                        
                        m_area = re.search(r"(\d+)\s*m[²2]", text_content)
                        if m_area:
                            area = float(m_area.group(1))

                        m_rooms = re.search(r"(\d+)\s*(?:hab|dorm)", text_content, re.IGNORECASE)
                        if m_rooms:
                            rooms = int(m_rooms.group(1))

                        if area < MIN_AREA or rooms < MIN_ROOMS:
                            continue

                        status = analyze_listing(title, title, price, area, rooms)
                        price_m2 = price / area if area > 0 else 0.0

                        data_rows.append({
                            "Название / Адрес": title,
                            "Цена (€)": price,
                            "Площадь (м²)": area,
                            "Комнат": rooms,
                            "Цена за м² (€)": round(price_m2, 2),
                            "Статус / Причина отфильтровки": status,
                            "Ссылка": link
                        })
                    except Exception as e:
                        print(f"Ошибка парсинга отдельной карточки: {e}")

            if not data_rows:
                data_rows.append({
                    "Название / Адрес": f"⚠️ Требуется уточнение селекторов для {region_name}",
                    "Цена (€)": 0,
                    "Площадь (м²)": 0,
                    "Комнат": 0,
                    "Цена за м² (€)": 0,
                    "Статус / Причина отфильтровки": "Нулевой результат парсинга",
                    "Ссылка": target_url
                })

        except Exception as e:
            print(f"Ошибка запроса для {region_name}: {e}")

        df = pd.DataFrame(data_rows)
        
        if not df.empty and df[df["Цена (€)"] > 0].shape[0] > 0:
            valid_df = df[df["Цена (€)"] > 0]
            avg_price_m2 = valid_df["Цена за м² (€)"].mean()
            summary_row = pd.DataFrame([{
                "Название / Адрес": "СРЕДНЯЯ ЦЕНА ЗА М² ЗА ДЕНЬ:",
                "Цена (€)": None,
                "Площадь (м²)": None,
                "Комнат": None,
                "Цена за м² (€)": round(avg_price_m2, 2),
                "Статус / Причина отфильтровки": "-",
                "Ссылка": "-"
            }])
            df = pd.concat([df, summary_row], ignore_index=True)

        df.to_excel(writer, sheet_name=region_name, index=False)
    
    writer.close()
    return excel_file

# ==============================================================================
# 4. ОТПРАВКА ОТЧЕТА НА ПОЧТУ
# ==============================================================================

def send_email(file_path):
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg['Subject'] = f"📊 Ежедневный отчёт по недвижимости (Fotocasa) — {datetime.date.today()}"

    body = "Привет! Во вложении свежий аналитический отчёт по недвижимости (Мадрид, Толедо, Гвадалахара) по вашим критериям.\n\nФайл разбит на 3 вкладки с расчетом средних цен за м²."
    msg.attach(MIMEText(body, 'plain'))

    with open(file_path, "rb") as f:
        attach = MIMEApplication(f.read(), _subtype="xlsx")
        attach.add_header('Content-Disposition', 'attachment', filename=file_path)
        msg.attach(attach)

    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.login(SENDER_EMAIL, SENDER_PASSWORD)
    server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())
    server.quit()
    print("Отчёт успешно отправлен на email!")

if __name__ == "__main__":
    report_file = collect_stats()
    send_email(report_file)
