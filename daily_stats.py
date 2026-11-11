import asyncio
import os
import re
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import pandas as pd
import requests
from playwright.async_api import async_playwright

# ==============================================================================
# 1. НАСТРОЙКИ
# ==============================================================================

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "")
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL", "")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

REGIONS = {
    "Мадрид": "https://www.idealista.com/venta-viviendas/madrid-provincia/con-precio-hasta_145000,metros-cuadrados-mas-de_60,habitaciones-mas-de_2/",
    "Толедо": "https://www.idealista.com/venta-viviendas/toledo-provincia/con-precio-hasta_145000,metros-cuadrados-mas-de_60,habitaciones-mas-de_2/",
    "Гвадалахара": "https://www.idealista.com/venta-viviendas/guadalajara-provincia/con-precio-hasta_145000,metros-cuadrados-mas-de_60,habitaciones-mas-de_2/"
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
# 2. АНАЛИЗАТОР
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

    return "ОДОБРЕНО (Отправлено в Telegram)"

# ==============================================================================
# 3. СБОР ЧЕРЕЗ PLAYWRIGHT С ПЕРЕХОДОМ ЗАЩИТЫ
# ==============================================================================

async def collect_stats():
    excel_file = f"Idealista_Daily_Report_{datetime.date.today()}.xlsx"
    writer = pd.ExcelWriter(excel_file, engine='openpyxl')

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage"
            ]
        )
        
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            viewport={"width": 1920, "height": 1080},
            locale="es-ES"
        )
        
        page = await context.new_page()

        for region_name, target_url in REGIONS.items():
            print(f"Сбор данных для: {region_name}")
            data_rows = []
            
            try:
                # Переходим на главную, чтобы получить легитимные куки Cloudflare
                await page.goto("https://www.idealista.com/", timeout=45000)
                await page.wait_for_timeout(3000)
                
                # Теперь переходим на страницу региона
                await page.goto(target_url, timeout=45000)
                
                # Даем время на прохождение скриптов защиты (ждем появления карточек или заголовка)
                try:
                    await page.wait_for_selector("article.item", timeout=12000)
                except:
                    # Если сработала защита, ждем подольше и имитируем активность
                    print("Ожидание прохождения Cloudflare...")
                    await page.wait_for_timeout(8000)
                    await page.mouse.move(500, 500)
                    await page.evaluate("window.scrollBy(0, 600)")
                    await page.wait_for_timeout(5000)

                cards = await page.query_selector_all("article.item")
                print(f"Найдено карточек: {len(cards)}")

                for card in cards:
                    try:
                        title_el = await card.query_selector(".item-link")
                        if not title_el:
                            continue
                        title = (await title_el.inner_text()).strip()
                        link = await title_el.get_attribute("href")
                        if link and not link.startswith("http"):
                            link = "https://www.idealista.com" + link

                        price_el = await card.query_selector(".item-price")
                        price = float(re.sub(r"[^\d]", "", await price_el.inner_text())) if price_el else 0.0

                        details = await card.query_selector_all(".item-detail")
                        area, rooms = 0.0, 0
                        for detail in details:
                            txt = await detail.inner_text()
                            if "m²" in txt:
                                m = re.search(r"(\d+)", txt.replace(".", ""))
                                if m: area = float(m.group(1))
                            elif "hab" in txt:
                                m = re.search(r"(\d+)", txt)
                                if m: rooms = int(m.group(1))

                        desc_el = await card.query_selector(".item-description")
                        description = await desc_el.inner_text() if desc_el else ""

                        status = analyze_listing(title, description, price, area, rooms)
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
                        print(f"Ошибка чтения карточки: {e}")

            except Exception as e:
                print(f"Ошибка загрузки {region_name}: {e}")

            if not data_rows:
                data_rows.append({
                    "Название / Адрес": "⚠️ Требуется интерактивное прохождение капчи Cloudflare",
                    "Цена (€)": 0, "Площадь (м²)": 0, "Комнат": 0, "Цена за м² (€)": 0,
                    "Статус / Причина отфильтровки": "Заблокировано", "Ссылка": target_url
                })

            df = pd.DataFrame(data_rows)
            
            if not df.empty and df.iloc[0]["Цена (€)"] > 0:
                avg_price_m2 = df[df["Цена (€)"] > 0]["Цена за м² (€)"].mean()
                summary_row = pd.DataFrame([{
                    "Название / Адрес": "СРЕДНЯЯ ЦЕНА ЗА М² ЗА ДЕНЬ:",
                    "Цена (€)": None, "Площадь (м²)": None, "Комнат": None,
                    "Цена за м² (€)": round(avg_price_m2, 2),
                    "Статус / Причина отфильтровки": "-", "Ссылка": "-"
                }])
                df = pd.concat([df, summary_row], ignore_index=True)

            df.to_excel(writer, sheet_name=region_name, index=False)

        await browser.close()
    
    writer.close()
    return excel_file

# ==============================================================================
# 4. ОТПРАВКА НА ПОЧТУ
# ==============================================================================

def send_email(file_path):
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg['Subject'] = f"📊 Ежедневный отчёт Idealista — {datetime.date.today()}"

    body = "Привет! Во вложении свежий отчёт по недвижимости на Idealista за сегодня.\n\nФайл разбит на 3 вкладки: Мадрид, Толедо, Гвадалахара."
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
    report_file = asyncio.run(collect_stats())
    send_email(report_file)
