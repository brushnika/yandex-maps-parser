# -*- coding: utf-8 -*-
"""
ГИБКИЙ И БЫСТРЫЙ ПАРСЕР КОМПАНИЙ С ЯНДЕКС.КАРТ

УЛУЧШЕНИЯ:
1. Ускоренный сбор ссылок (оптимизированная прокрутка)
2. Ускоренный парсинг карточек (меньше ожиданий)
3. Автоматическая проверка дубликатов названий
4. Сбор большего количества ссылок для фильтрации дубликатов

Как использовать:
1. Измените настройки в блоке НАСТРОЙКИ ниже
2. Запустите скрипт
3. Повторите для других регионов/типов предприятий
"""

import time
import random
import re
import os
from datetime import datetime
from typing import Dict, List, Optional, Set

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    StaleElementReferenceException,
    WebDriverException,
)
from webdriver_manager.chrome import ChromeDriverManager

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# =====================================================================
#                        НАСТРОЙКИ (МЕНЯЙТЕ ЗДЕСЬ)
# =====================================================================

SEARCH_QUERY = "Кожевенные и обувные производства"
CITY = "Брянская область"
REGION_FULL = "Брянская область"
FEDERAL_DISTRICT = "ЦФО"
MAX_COMPANIES = 50

# Паузы (уменьшены для скорости)
PAUSE_MIN = 0.8
PAUSE_MAX = 1.5

# Таймауты (уменьшены)
TIMEOUT_SHORT = 2
TIMEOUT_MEDIUM = 4

OUTPUT_DIR = "."


# =====================================================================
#                    ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =====================================================================

def random_pause(min_sec: float = PAUSE_MIN, max_sec: float = PAUSE_MAX) -> None:
    time.sleep(random.uniform(min_sec, max_sec))


def create_driver() -> webdriver.Chrome:
    """Создаёт и настраивает браузер."""
    options = Options()
    
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-plugins")
    options.add_argument("--disable-images")  # Отключаем загрузку картинок для скорости
    options.add_argument("--disable-javascript")  # ВАЖНО: НЕ ОТКЛЮЧАТЬ!
    
    prefs = {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        "profile.default_content_setting_values.notifications": 2,
        "profile.default_content_setting_values.images": 2,  # Отключаем картинки
    }
    options.add_experimental_option("prefs", prefs)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    
    user_agents = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
    ]
    options.add_argument(f"user-agent={random.choice(user_agents)}")
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    
    # Устанавливаем таймауты
    driver.set_page_load_timeout(15)
    driver.set_script_timeout(10)
    
    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {
                "source": """
                    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                """
            },
        )
    except WebDriverException:
        pass
    
    return driver


def safe_get_text(element, default: str = "Н/Д") -> str:
    if element is None:
        return default
    try:
        text = element.text.strip()
        return text if text else default
    except Exception:
        return default


def find_element_fast(driver, selectors: List[str], timeout: int = TIMEOUT_SHORT):
    """Быстрый поиск элемента без лишних ожиданий."""
    for selector in selectors:
        try:
            element = driver.find_element(By.CSS_SELECTOR, selector)
            if element:
                return element
        except Exception:
            continue
    return None


def find_elements_fast(driver, selectors: List[str]):
    """Быстрый поиск всех элементов."""
    for selector in selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)
            if elements:
                return elements
        except Exception:
            continue
    return []


def clean_phone(phone: str) -> str:
    if not phone or phone == "Н/Д":
        return phone
    phone = re.sub(r'\s+', ' ', phone.strip())
    phone = re.sub(r'[()\-]', '', phone)
    return phone


# =====================================================================
#                     ПАРСИНГ ЯНДЕКС.КАРТ (УСКОРЕННЫЙ)
# =====================================================================

def perform_search(driver) -> bool:
    """Выполняет поиск на Яндекс.Картах (ускоренный)."""
    search_text = f"{SEARCH_QUERY} {CITY}"
    
    print(f"🔍 Поиск: {search_text}")
    driver.get("https://yandex.ru/maps/")
    time.sleep(1.5)  # Уменьшено ожидание
    
    search_selectors = [
        "input.input__control",
        "input#text",
        "input[type='text']",
        "input[name='text']",
        "input[placeholder*='Поиск']",
    ]
    
    search_input = None
    for _ in range(5):
        search_input = find_element_fast(driver, search_selectors, timeout=2)
        if search_input:
            break
        time.sleep(0.3)
    
    if search_input is None:
        print("❌ Не найдено поле поиска")
        return False
    
    search_input.click()
    time.sleep(0.2)
    search_input.clear()
    
    # Быстрый ввод
    for char in search_text:
        search_input.send_keys(char)
        time.sleep(0.02)
    
    time.sleep(0.3)
    
    button_selectors = [
        "button[aria-label='Найти']",
        "button.input__button",
        "button[type='submit']",
    ]
    
    search_button = find_element_fast(driver, button_selectors)
    if search_button:
        try:
            search_button.click()
        except Exception:
            search_input.send_keys("\n")
    else:
        search_input.send_keys("\n")
    
    time.sleep(2.0)  # Ждём загрузки результатов
    return True


def collect_card_links_fast(driver) -> List[str]:
    """
    БЫСТРЫЙ сбор ссылок с оптимизированной прокруткой.
    Собираем больше ссылок, чтобы потом отфильтровать дубликаты.
    """
    print("📋 Сбор ссылок на карточки...")
    time.sleep(1.5)
    
    # Ищем контейнер
    container_selectors = [
        ".scroll__container",
        "[class*='scroll__container']",
        "[class*='search-list']",
    ]
    container = find_element_fast(driver, container_selectors, timeout=3)
    
    # Собираем ссылки
    link_selectors = [
        "a[href*='/maps/org/']",
        "a[class*='snippet-view__link']",
    ]
    
    all_links = []
    collected_links = []
    seen_names = set()
    prev_count = 0
    no_change_count = 0
    scroll_attempts = 0
    
    # Прокручиваем и собираем
    while len(collected_links) < MAX_COMPANIES + 20:  # С запасом
        # Собираем ссылки
        for selector in link_selectors:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector)
                for el in elements:
                    href = el.get_attribute("href")
                    if href and href not in all_links and "/maps/org/" in href:
                        all_links.append(href)
            except Exception:
                continue
        
        # Проверяем уникальные ссылки
        for link in all_links:
            if link not in collected_links:
                collected_links.append(link)
        
        if len(collected_links) > 0:
            print(f"  Собрано уникальных ссылок: {len(collected_links)}")
        
        # Проверяем изменения
        if len(collected_links) == prev_count:
            no_change_count += 1
            if no_change_count >= 3:  # Быстрее выходим
                break
        else:
            no_change_count = 0
            prev_count = len(collected_links)
        
        # Прокрутка с оптимизацией
        try:
            if container:
                driver.execute_script(
                    "arguments[0].scrollTop = arguments[0].scrollTop + 600;",
                    container
                )
            else:
                driver.execute_script("window.scrollBy(0, 600);")
        except Exception:
            driver.execute_script("window.scrollBy(0, 600);")
        
        scroll_attempts += 1
        time.sleep(0.5)  # Минимальная пауза между прокрутками
        
        # Если много прокруток, но мало ссылок - выходим
        if scroll_attempts > 20 and len(collected_links) < 5:
            break
    
    return collected_links[:MAX_COMPANIES + 20]  # Возвращаем с запасом


def extract_company_data_fast(driver, url: str, existing_names: Set[str]) -> Optional[Dict]:
    """
    БЫСТРОЕ извлечение данных с проверкой дубликатов названий.
    Возвращает None, если название уже есть (дубликат).
    """
    data = {
        "ФО": FEDERAL_DISTRICT,
        "РЕГИОН": REGION_FULL,
        "ГОРОД": "Н/Д",
        "НАИМ": "Н/Д",
        "КРАТКНАИМ": "Н/Д",
        "ФИО": "Н/Д",
        "ИО": "Н/Д",
        "ИНФО": "Н/Д",
        "КОМУ": "Н/Д",
        "АДРЕС": "Н/Д",
        "ТЕЛ": "Н/Д",
        "ПОЧТА1": "Н/Д",
        "ССЫЛКИ": "Н/Д",
        "Ссылка на карточку": url,
    }
    
    try:
        driver.get(url)
        time.sleep(0.8)  # Минимальное ожидание
        
        # НАЗВАНИЕ (быстро)
        title_selectors = [
            "h1[class*='title']",
            "h1[class*='header']",
            "h1.card-title-view__title",
            "h1",
        ]
        title_elem = find_element_fast(driver, title_selectors, timeout=2)
        company_name = safe_get_text(title_elem)
        
        # ПРОВЕРКА НА ДУБЛИКАТ
        if company_name in existing_names:
            return None  # Пропускаем дубликат
        
        data["НАИМ"] = company_name
        
        # Краткое наименование
        if "Акционерное общество" in company_name:
            data["КРАТКНАИМ"] = company_name.replace("Акционерное общество", "АО")
        elif "Общество с ограниченной ответственностью" in company_name:
            data["КРАТКНАИМ"] = company_name.replace("Общество с ограниченной ответственностью", "ООО")
        elif "Публичное акционерное общество" in company_name:
            data["КРАТКНАИМ"] = company_name.replace("Публичное акционерное общество", "ПАО")
        else:
            data["КРАТКНАИМ"] = company_name
        
        # АДРЕС
        address_selectors = [
            "[class*='address']",
            "[class*='orgpage-header-view__address']",
            "span[class*='address']",
        ]
        address_elem = find_element_fast(driver, address_selectors)
        data["АДРЕС"] = safe_get_text(address_elem)
        
        if data["АДРЕС"] != "Н/Д":
            city_match = re.search(r'г\.\s*([А-Яа-я\-]+)', data["АДРЕС"])
            if city_match:
                data["ГОРОД"] = city_match.group(1)
        
        # ТЕЛЕФОН (быстро)
        phone_button_selectors = [
            "[class*='phone-button']",
            "button[class*='show-phone']",
        ]
        phone_button = find_element_fast(driver, phone_button_selectors, timeout=1)
        if phone_button:
            try:
                phone_button.click()
                time.sleep(0.3)
            except Exception:
                pass
        
        phone_selectors = [
            "a[href^='tel:']",
            "[class*='phone-number']",
            "[class*='phone']",
        ]
        phone_elem = find_element_fast(driver, phone_selectors)
        if phone_elem:
            href = phone_elem.get_attribute("href")
            if href and href.startswith("tel:"):
                data["ТЕЛ"] = clean_phone(href.replace("tel:", ""))
            else:
                data["ТЕЛ"] = clean_phone(safe_get_text(phone_elem))
        
        if data["ТЕЛ"] == "Н/Д":
            try:
                body_text = driver.find_element(By.TAG_NAME, "body").text[:3000]
                phone_match = re.search(
                    r'(\+7|8)[\s\-]?\(?\d{3,4}\)?[\s\-]?\d{2,3}[\s\-]?\d{2}[\s\-]?\d{2}',
                    body_text
                )
                if phone_match:
                    data["ТЕЛ"] = clean_phone(phone_match.group(0))
            except Exception:
                pass
        
        # САЙТ
        website_selectors = [
            "a[class*='website']",
            "a[class*='url']",
            "[class*='website'] a",
        ]
        website_elem = find_element_fast(driver, website_selectors)
        if website_elem:
            href = website_elem.get_attribute("href")
            if href:
                data["ССЫЛКИ"] = href
            else:
                data["ССЫЛКИ"] = safe_get_text(website_elem)
        
        # EMAIL
        try:
            body_text = driver.find_element(By.TAG_NAME, "body").text
            email_match = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', body_text)
            if email_match:
                data["ПОЧТА1"] = email_match.group(0)
        except Exception:
            pass
        
        # РУКОВОДИТЕЛЬ
        try:
            body_text = driver.find_element(By.TAG_NAME, "body").text
            name_patterns = [
                r'Генеральный директор\s*[:\-]?\s*([А-Я][а-я]+\s+[А-Я][а-я]+\s+[А-Я][а-я]+)',
                r'Директор\s*[:\-]?\s*([А-Я][а-я]+\s+[А-Я][а-я]+\s+[А-Я][а-я]+)',
            ]
            for pattern in name_patterns:
                match = re.search(pattern, body_text)
                if match:
                    full_name = match.group(1)
                    data["ФИО"] = full_name
                    name_parts = full_name.split()
                    if len(name_parts) >= 2:
                        data["ИО"] = f"{name_parts[1]} {name_parts[2] if len(name_parts) > 2 else ''}".strip()
                    data["ИНФО"] = "Генеральный директор"
                    data["КОМУ"] = f"Генеральному директору {full_name}"
                    break
        except Exception:
            pass
        
        # Добавляем название в множество (уникальные названия)
        existing_names.add(company_name)
        
        print(f"  ✓ {data['НАИМ'][:40]}")
        return data
        
    except Exception as e:
        print(f"  ✗ Ошибка: {str(e)[:50]}")
        return None


# =====================================================================
#                     СОХРАНЕНИЕ РЕЗУЛЬТАТОВ
# =====================================================================

def save_to_excel(companies: List[Dict]) -> str:
    """Сохраняет данные в Excel."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    def clean_filename(text: str) -> str:
        text = text.strip().replace(" ", "_")
        return re.sub(r'[\\/*?:"<>|]', "", text)
    
    date_str = datetime.now().strftime("%Y-%m-%d_%H-%M")
    filename = f"{clean_filename(CITY)}_{clean_filename(SEARCH_QUERY)}_{date_str}.xlsx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Компании"
    
    headers = [
        "ФО", "РЕГИОН", "ГОРОД", "НАИМ", "КРАТКНАИМ",
        "ФИО", "ИО", "ИНФО", "КОМУ",
        "АДРЕС", "ТЕЛ", "ПОЧТА1", "ССЫЛКИ", "Ссылка на карточку"
    ]
    ws.append(headers)
    
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    
    for company in companies:
        row_data = []
        for header in headers:
            value = company.get(header, "Н/Д")
            if value is None or value == "":
                value = "Н/Д"
            row_data.append(value)
        ws.append(row_data)
    
    column_widths = {
        "A": 8, "B": 18, "C": 15, "D": 45, "E": 30,
        "F": 28, "G": 22, "H": 20, "I": 40,
        "J": 35, "K": 15, "L": 25, "M": 30, "N": 40
    }
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width
    
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = thin_border
    
    wb.save(filepath)
    print(f"\n✅ Данные сохранены: {os.path.abspath(filepath)}")
    return filepath


# =====================================================================
#                         ГЛАВНАЯ ФУНКЦИЯ
# =====================================================================

def main():
    print("=" * 70)
    print("🚀 БЫСТРЫЙ ПАРСЕР КОМПАНИЙ С ЯНДЕКС.КАРТ")
    print("=" * 70)
    print(f"\n📌 ТЕКУЩИЕ НАСТРОЙКИ:")
    print(f"   Тип предприятий: {SEARCH_QUERY}")
    print(f"   Регион: {CITY}")
    print(f"   Нужно компаний: {MAX_COMPANIES}")
    print(f"   ФО: {FEDERAL_DISTRICT}")
    
    start_time = time.time()
    driver = create_driver()
    companies = []
    existing_names = set()
    
    try:
        if not perform_search(driver):
            print("❌ Не удалось выполнить поиск")
            driver.quit()
            return
        
        links = collect_card_links_fast(driver)
        
        if not links:
            print("❌ Не найдено компаний")
            driver.quit()
            return
        
        print(f"\n📊 Найдено {len(links)} ссылок. Начинаем сбор (пропускаем дубликаты)...\n")
        
        collected_count = 0
        for idx, link in enumerate(links, start=1):
            if collected_count >= MAX_COMPANIES:
                break
                
            try:
                data = extract_company_data_fast(driver, link, existing_names)
                if data is not None:
                    companies.append(data)
                    collected_count += 1
                    print(f"[{collected_count}/{MAX_COMPANIES}] Уникальных компаний")
                else:
                    print(f"[{idx}/{len(links)}] Пропущено (дубликат или ошибка)")
            except Exception as e:
                print(f"[{idx}/{len(links)}] Ошибка: {e}")
            
            # Пауза только между успешными парсингами
            if len(companies) % 5 == 0:
                random_pause(0.5, 1.0)
        
    finally:
        driver.quit()
    
    elapsed = time.time() - start_time
    minutes = int(elapsed // 60)
    seconds = int(elapsed % 60)
    
    print(f"\n{'='*70}")
    print(f"📊 ИТОГИ (за {minutes} мин {seconds} сек):")
    print(f"   Собрано уникальных компаний: {len(companies)}")
    print(f"   Требовалось: {MAX_COMPANIES}")
    
    if companies:
        filepath = save_to_excel(companies)
        print(f"   Файл: {os.path.basename(filepath)}")
    else:
        print("\n❌ Не удалось собрать данные")
    print(f"{'='*70}")


if __name__ == "__main__":
    main()