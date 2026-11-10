import asyncio
import os
import re
import sqlite3
import requests
from playwright.async_api import async_playwright

# ==============================================================================
# 1. НАСТРОЙКИ И КОНФИГУРАЦИЯ
# ==============================================================================

# Токен бота и Chat ID из Telegram
TELEGRAM_BOT_TOKEN = "9124072670:AAFRQvxf_pIFhdGyNLYJrvLGyYVTWlR7Sw0"
TELEGRAM_CHAT_ID = "775383357"
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

# ИЗМЕНЕННЫЕ ССЫЛКИ: Сам Мадрид (город) + Провинция Толедо (Ла-Сагра)
TARGET_URLS = [
    # Мадрид (сам город) с фильтрами: цена до 145k, от 60м², от 2 комнат
    "https://www.idealista.com/venta-viviendas/madrid/con-precio-hasta_145000,metros-cuadrados-mas-de_60,habitaciones-mas-de_2/",
    # Провинция Толедо (Ла-Сагра и окрестности)
    "https://www.idealista.com/venta-viviendas/toledo-provincia/con-precio-hasta_145000,metros-cuadrados-mas-de_60,habitaciones-mas-de_2/"
]

# Параметры фильтрации в коде
MAX_PRICE = 145000.0
MIN_AREA = 60.0
MIN_ROOMS = 2

# Опасные и неблагополучные города/районы (Черный список)
FORBIDDEN_LOCATIONS = [
    # Опасные городки Ла-Сагры (Okupas / Криминал)
    "chozas de canales", "chozas", "lominchar", "recas", 
    "seseña", "sesena", "quiñon", "quiñón", "yuncos", "yeles",
    # Зоны риска паводков и затоплений (DANA)
    "cobeja", "magan", "magán",
    # Неблагополучные зоны Мадрида и пригородов
    "parla", "valdemingomez", "valdemingómez", "cañada real", "alhondiga", "alhóndiga",
    "el gorronal", "gorronal"
]

# Юридические проблемы, оккупасы, аукционы и аварийное жилье
LEGAL_RISK_WORDS = [
    "okupad", "ocupad", "sin posesion", "sin posesión", "no se puede visitar", 
    "nuda propiedad", "proindiviso", "parte indivisa", "alquilad", "con inquilino",
    "subasta", "cesion de remate", "cesión de remate", "inmueble ilocalizable",
    "para derribar", "ruina", "sin cedula", "sin cédula"
]

# Ключевые слова для "Изюминки" (🔴 Красный маркер)
HIGHLIGHT_KEYWORDS = [
    "terraza", "piscina", "jardín", "jardin", "patio", 
    "porche", "atico", "ático", "chalet", "adosado", "pareado"
]

DB_FILE = "real_estate_agent.db"

# ==============================================================================
# 2. РАБОТА С БАЗОЙ ДАННЫХ (Исключение дублей)
# ==============================================================================

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS seen_properties (
            id TEXT PRIMARY KEY,
            price REAL,
            title TEXT,
            status TEXT
        )
    """)
    conn.commit()
    conn.close()

def is_already_seen(prop_id):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT 1 FROM seen_properties WHERE id = ?", (prop_id,))
    res = cursor.fetchone()
    conn.close()
    return res is not None

def save_property(prop_id, price, title, status="PASSED"):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("INSERT OR IGNORE INTO seen_properties VALUES (?, ?, ?, ?)", (prop_id, price, title, status))
    conn.commit()
    conn.close()

# ==============================================================================
# 3. ИНТЕЛЛЕКТУАЛЬНЫЙ АНАЛИЗАТОР (GEMINI LLM)
# ==============================================================================

def analyze_with_ai(title, description):
    if not GEMINI_API_KEY:
        return True, "AI проверка пропущена (нет API ключа)"
        
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}"
    
    prompt = f"""
    Проанализируй текст объявления о продаже недвижимости в Испании для семьи с маленьким ребенком.
    Заголовок: {title}
    Описание: {description}
    
    Ответь строго в формате JSON с двумя полями:
    1. "safe_for_living" (boolean): true - если объект пригоден для нормальной покупки и жизни. false - если есть проблемы (оккупасы, арендаторы, продается только доля/nuda propiedad, невозможность визита, аварийное состояние, дом под снос).
    2. "reason" (string): Краткое резюме на русском языке (1 предложение), почему объект подходит или почему отклонен.
    """
    
    try:
        response = requests.post(
            url, 
            json={"contents": [{"parts": [{"text": prompt}]}]},
            headers={"Content-Type": "application/json"},
            timeout=10
        )
        if response.status_code == 200:
            result_text = response.json()['candidates'][0]['content']['parts'][0]['text']
            is_safe = '"safe_for_living": true' in result_text.lower()
            return is_safe, result_text
    except Exception as e:
        print(f"Ошибка Gemini API: {e}")
        
    return True, "Одобрено (базовый фильтр)"

# ==============================================================================
# 4. ОТПРАВКА УВЕДОМЛЕНИЙ В TELEGRAM
# ==============================================================================

def send_telegram_msg(title, price, area, rooms, location, link, is_highlight=False, ai_note=""):
    header = "🔴 **ИЗЮМИНКА! (Дворик / Терраса / Бассейн)**" if is_highlight else "🟡 **Новое предложение**"
    price_per_m2 = price / area if area > 0 else 0
    
    message = (
        f"{header}\n\n"
        f"🏠 **{title}**\n"
        f"📍 Локация: {location}\n"
        f"💰 Цена: **{price:,.0f} €** ({price_per_m2:.0f} €/м²)\n"
        f"📐 Площадь: **{area:.0f} м²** | 🛏 Комнат: **{rooms}**\n"
    )
    
    if ai_note:
        message += f"🤖 AI-Примечание: _{ai_note}_\n"
        
    message += f"\n🔗 [Смотреть объявление на Idealista]({link})"
    
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "parse_mode": "Markdown",
        "disable_web_page_preview": False
    }
    try:
        requests.post(url, json=payload, timeout=10)
    except Exception as e:
        print(f"Ошибка отправки в Telegram: {e}")

# ==============================================================================
# 5. ОСНОВНОЙ СКРИПТ (ПАРСИНГ)
# ==============================================================================

async def run_agent():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 800}
        )
        page = await context.new_page()

        for url in TARGET_URLS:
            print(f"Сканирование: {url}")
            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=60000)
                await page.wait_for_timeout(3000)
            except Exception as e:
                print(f"Ошибка загрузки страницы: {e}")
                continue

            cards = await page.query_selector_all("article.item")

            for card in cards:
                try:
                    prop_id = await card.get_attribute("data-element-id")
                    if not prop_id or is_already_seen(prop_id):
                        continue

                    title_el = await card.query_selector(".item-link")
                    if not title_el:
                        continue
                    title = (await title_el.inner_text()).strip()
                    link = await title_el.get_attribute("href")
                    if link and not link.startswith("http"):
                        link = "https://www.idealista.com" + link

                    title_lower = title.lower()

                    # 1. Проверка черного списка локаций
                    if any(forbidden in title_lower for forbidden in FORBIDDEN_LOCATIONS):
                        save_property(prop_id, 0, title, status="REJECTED_LOCATION")
                        continue

                    # 2. Быстрая проверка юридических рисков и оккупасов
                    desc_el = await card.query_selector(".item-description")
                    description = (await desc_el.inner_text()).lower() if desc_el else ""
                    
                    full_text = f"{title_lower} {description}"
                    if any(risk in full_text for risk in LEGAL_RISK_WORDS):
                        print(f"Пропущено (Юридический риск / Okupas): {title}")
                        save_property(prop_id, 0, title, status="REJECTED_LEGAL_RISK")
                        continue

                    # Извлечение цены
                    price_el = await card.query_selector(".item-price")
                    if not price_el:
                        continue
                    price_text = await price_el.inner_text()
                    price = float(re.sub(r"[^\d]", "", price_text))

                    if price > MAX_PRICE:
                        save_property(prop_id, price, title, status="REJECTED_PRICE")
                        continue

                    # Площадь и комнаты
                    details = await card.query_selector_all(".item-detail")
                    area, rooms = 0.0, 0
                    for detail in details:
                        text = await detail.inner_text()
                        if "m²" in text:
                            m = re.search(r"(\d+)", text.replace(".", ""))
                            if m: area = float(m.group(1))
                        elif "hab" in text:
                            m = re.search(r"(\d+)", text)
                            if m: rooms = int(m.group(1))

                    if area < MIN_AREA or rooms < MIN_ROOMS:
                        save_property(prop_id, price, title, status="REJECTED_SPECS")
                        continue

                    # 3. Дополнительная проверка через Gemini AI
                    is_safe_ai, ai_note = analyze_with_ai(title, description)
                    if not is_safe_ai:
                        print(f"Пропущено AI: {title}")
                        save_property(prop_id, price, title, status="REJECTED_AI")
                        continue

                    # Определение "Изюминки"
                    is_highlight = any(kw in full_text for kw in HIGHLIGHT_KEYWORDS)
                    location = title.split("en")[-1].strip() if "en" in title else "Мадрид / Толедо"

                    # Отправка в Telegram
                    send_telegram_msg(
                        title=title,
                        price=price,
                        area=area,
                        rooms=rooms,
                        location=location,
                        link=link,
                        is_highlight=is_highlight
                    )
                    
                    save_property(prop_id, price, title, status="PASSED")
                    print(f"Успешно отправлено: {title}")

                except Exception as e:
                    print(f"Ошибка обработки: {e}")

        await browser.close()

if __name__ == "__main__":
    init_db()
    asyncio.run(run_agent())
