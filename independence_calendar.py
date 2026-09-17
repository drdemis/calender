"""
Independence & Holiday Calendar
--------------------------------
A desktop app (Tkinter, no extra installs needed) that tracks, per country:
  - President name
  - Minister of Defence name
  - Ambassador to Lebanon name
  - A list of important dates (Independence Day + any other holidays)

Features
  - Main panel: table of all countries with their NEXT upcoming holiday
    and how many days remain, plus a details panel showing the
    president / minister of defence / ambassador for the selected country.
  - Notifications: when you open the app (and every hour while it's open)
    it checks all holidays and pops up an alert for any holiday that is
    within the "notify me X days before" threshold (default 7 days,
    adjustable in the main tab).
  - "Add / Edit Country" tab: create a new country, edit an existing one,
    delete a country, and add/remove its holidays.
  - Everything is saved to a local JSON file (countries.json) next to this
    script, so your data persists between runs.

How to run in VS Code
  1. Make sure Python 3 is installed (Tkinter ships with standard Python
     on Windows/macOS; on Linux you may need `sudo apt install python3-tk`).
  2. Open this file in VS Code and press Run (or `python independence_calendar.py`
     in the terminal).
"""

import json
import os
import re
import sys
import threading
import tkinter as tk
from pathlib import Path
from urllib.request import Request, urlopen, urlretrieve
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import datetime, date

try:
    import pystray
except ImportError:  # pragma: no cover
    pystray = None

try:
    from PIL import Image, ImageDraw
except ImportError:  # pragma: no cover
    Image = None
    ImageDraw = None

try:
    from openpyxl import Workbook, load_workbook
except ImportError:  # pragma: no cover
    Workbook = None
    load_workbook = None

APP_DIR = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
RESOURCE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
DATA_FILE = os.path.join(APP_DIR, "countries.json")
BUNDLED_DATA_FILE = os.path.join(RESOURCE_DIR, "countries.json")

DEFAULT_NOTIFY_DAYS = 7
DEVELOPER_PASSWORD = "331902"
TRANSLATIONS = {
    "en": {
        "app_title": "Independence & Holiday Calendar",
        "main_panel": "Main Panel",
        "edit_panel": "Add / Edit Country",
        "notify_me": "Notify me when a holiday is within:",
        "days": "days",
        "check_now": "Check Now",
        "next_notification": "Next: {name} - {date}",
        "no_next_notification": "Next notification: none",
        "previous": "Previous",
        "next": "Next",
        "done": "Done",
        "refresh": "Refresh",
        "import_excel": "Import Excel",
        "export_excel": "Export Excel",
        "search_country": "Search country:",
        "country": "Country",
        "next_holiday": "Next Holiday",
        "date": "Date",
        "days_left": "Days Left",
        "country_details": "Country Details",
        "flag": "Flag",
        "country_label": "Country:",
        "president_label": "President:",
        "minister_label": "Minister of Defence:",
        "ambassador_label": "Ambassador to Lebanon:",
        "embassy_location_label": "Embassy Location:",
        "embassy_phone_label": "Embassy Phone:",
        "embassy_location_tab": "Location",
        "embassy_phone_tab": "Phone",
        "holidays": "Holidays:",
        "language": "Language:",
        "english": "English",
        "arabic": "العربية",
        "select_country_to_edit": "Select country to edit:",
        "new_country": "New Country",
        "delete_country": "Delete Country",
        "country_name": "Country Name:",
        "president": "President:",
        "minister_of_defence": "Minister of Defence:",
        "ambassador_to_lebanon": "Ambassador to Lebanon:",
        "holiday_name": "Holiday name:",
        "date_hint": "Date (MM-DD):",
        "add_holiday": "Add Holiday",
        "remove_selected": "Remove Selected",
        "save_country": "Save Country",
        "upcoming_holidays": "Upcoming Holidays",
        "no_holidays_window": "No holidays within the notification window.",
        "today": "today",
        "day": "day",
        "days_short": "days",
        "missing_holiday": "Please enter both a holiday name and a date (MM-DD).",
        "invalid_date": "Please use the format MM-DD, e.g. 11-22 for Nov 22.",
        "missing_name": "Please enter a country name.",
        "missing_selection": "Select a country from the dropdown first.",
        "saved": "Saved",
        "import_complete": "Import complete",
        "export_complete": "Export complete",
        "import_failed": "Import failed",
        "export_failed": "Export failed",
        "confirm_delete": "Confirm delete",
        "delete_question": "Delete '{name}' and all its data?",
        "saved_country": "'{name}' has been saved.",
        "check_now_alt": "Check Now",
        "in_days": "in {n} day",
    },
    "ar": {
        "app_title": "تقويم الاستقلال والعطلات",
        "main_panel": "لوحة الرئيسية",
        "edit_panel": "إضافة / تعديل دولة",
        "notify_me": "أخبرني عندما تكون عطلة خلال:",
        "days": "أيام",
        "check_now": "تحقق الآن",
        "next_notification": "التنبيه التالي: {name} - {date}",
        "no_next_notification": "التنبيه التالي: لا يوجد",
        "previous": "السابق",
        "next": "التالي",
        "done": "تم",
        "refresh": "تحديث",
        "import_excel": "استيراد Excel",
        "export_excel": "تصدير Excel",
        "search_country": "ابحث عن دولة:",
        "country": "الدولة",
        "next_holiday": "العطلة القادمة",
        "date": "التاريخ",
        "days_left": "الأيام المتبقية",
        "country_details": "تفاصيل الدولة",
        "flag": "العلم",
        "country_label": "الدولة:",
        "president_label": "الرئيس:",
        "minister_label": "وزير الدفاع:",
        "ambassador_label": "السفير إلى لبنان:",
        "embassy_location_label": "موقع السفارة:",
        "embassy_phone_label": "هاتف السفارة:",
        "embassy_location_tab": "الموقع",
        "embassy_phone_tab": "الهاتف",
        "holidays": "العطلات:",
        "language": "اللغة:",
        "english": "English",
        "arabic": "العربية",
        "select_country_to_edit": "اختر دولة للتعديل:",
        "new_country": "دولة جديدة",
        "delete_country": "حذف الدولة",
        "country_name": "اسم الدولة:",
        "president": "الرئيس:",
        "minister_of_defence": "وزير الدفاع:",
        "ambassador_to_lebanon": "السفير إلى لبنان:",
        "holiday_name": "اسم العطلة:",
        "date_hint": "التاريخ (MM-DD):",
        "add_holiday": "إضافة عطلة",
        "remove_selected": "حذف المحدد",
        "save_country": "حفظ الدولة",
        "upcoming_holidays": "العطلات القادمة",
        "no_holidays_window": "لا توجد عطلات داخل نافذة التنبيه.",
        "today": "اليوم",
        "day": "يوم",
        "days_short": "أيام",
        "missing_holiday": "يرجى إدخال اسم عطلة وتاريخ (MM-DD).",
        "invalid_date": "يرجى استخدام الصيغة MM-DD، مثل 11-22 لـ 22 نوفمبر.",
        "missing_name": "يرجى إدخال اسم الدولة.",
        "missing_selection": "اختر دولة من القائمة أولاً.",
        "saved": "تم الحفظ",
        "import_complete": "اكتمل الاستيراد",
        "export_complete": "اكتمل التصدير",
        "import_failed": "فشل الاستيراد",
        "export_failed": "فشل التصدير",
        "confirm_delete": "تأكيد الحذف",
        "delete_question": "هل تريد حذف '{name}' وجميع بياناته؟",
        "saved_country": "تم حفظ '{name}'.",
        "check_now_alt": "تحقق الآن",
        "in_days": "خلال {n} يوم",
    },
}
FLAG_EMOJIS = {
    "Afghanistan": "🇦🇫",
    "Albania": "🇦🇱",
    "Algeria": "🇩🇿",
    "Angola": "🇦🇴",
    "Argentina": "🇦🇷",
    "Armenia": "🇦🇲",
    "Australia": "🇦🇺",
    "Austria": "🇦🇹",
    "Azerbaijan": "🇦🇿",
    "Bahrain": "🇧🇭",
    "Belarus": "🇧🇾",
    "Belgium": "🇧🇪",
    "Bolivia": "🇧🇴",
    "Bosnia and Herzegovina": "🇧🇦",
    "Botswana": "🇧🇼",
    "Brazil": "🇧🇷",
    "Britain": "🇬🇧",
    "Bulgaria": "🇧🇬",
    "Cambodia": "🇰🇭",
    "Cameroon": "🇨🇲",
    "Canada": "🇨🇦",
    "Chile": "🇨🇱",
    "China": "🇨🇳",
    "Colombia": "🇨🇴",
    "Costa Rica": "🇨🇷",
    "Croatia": "🇭🇷",
    "Cuba": "🇨🇺",
    "Czech Republic": "🇨🇿",
    "Denmark": "🇩🇰",
    "Dominican Republic": "🇩🇴",
    "Ecuador": "🇪🇨",
    "Egypt": "🇪🇬",
    "El Salvador": "🇸🇻",
    "England": "🇬🇧",
    "Estonia": "🇪🇪",
    "Ethiopia": "🇪🇹",
    "Finland": "🇫🇮",
    "France": "🇫🇷",
    "Gaza": "🇵🇸",
    "Georgia": "🇬🇪",
    "Germany": "🇩🇪",
    "Ghana": "🇬🇭",
    "Greece": "🇬🇷",
    "Guatemala": "🇬🇹",
    "Honduras": "🇭🇳",
    "Hungary": "🇭🇺",
    "Iceland": "🇮🇸",
    "India": "🇮🇳",
    "Indonesia": "🇮🇩",
    "Iran": "🇮🇷",
    "Iraq": "🇮🇶",
    "Ireland": "🇮🇪",
    "Italy": "🇮🇹",
    "Japan": "🇯🇵",
    "Jordan": "🇯🇴",
    "Kazakhstan": "🇰🇿",
    "Kenya": "🇰🇪",
    "Kuwait": "🇰🇼",
    "Kyrgyzstan": "🇰🇬",
    "Laos": "🇱🇦",
    "Latvia": "🇱🇻",
    "Lebanon": "🇱🇧",
    "Lithuania": "🇱🇹",
    "Malaysia": "🇲🇾",
    "Mexico": "🇲🇽",
    "Moldova": "🇲🇩",
    "Mongolia": "🇲🇳",
    "Morocco": "🇲🇦",
    "Myanmar": "🇲🇲",
    "Nepal": "🇳🇵",
    "Netherlands": "🇳🇱",
    "New Zealand": "🇳🇿",
    "Nigeria": "🇳🇬",
    "Nicaragua": "🇳🇮",
    "Norway": "🇳🇴",
    "Oman": "🇴🇲",
    "Pakistan": "🇵🇰",
    "Palestine": "🇵🇸",
    "Panama": "🇵🇦",
    "Paraguay": "🇵🇾",
    "Peru": "🇵🇪",
    "Philippines": "🇵🇭",
    "Poland": "🇵🇱",
    "Portugal": "🇵🇹",
    "Qatar": "🇶🇦",
    "Romania": "🇷🇴",
    "Russia": "🇷🇺",
    "Saudi": "🇸🇦",
    "Saudi Arabia": "🇸🇦",
    "Senegal": "🇸🇳",
    "Serbia": "🇷🇸",
    "Singapore": "🇸🇬",
    "Slovakia": "🇸🇰",
    "Slovenia": "🇸🇮",
    "South Africa": "🇿🇦",
    "South Korea": "🇰🇷",
    "Spain": "🇪🇸",
    "Sri Lanka": "🇱🇰",
    "Sweden": "🇸🇪",
    "Switzerland": "🇨🇭",
    "Syria": "🇸🇾",
    "Thailand": "🇹🇭",
    "Tunisia": "🇹🇳",
    "Turkey": "🇹🇷",
    "UAE": "🇦🇪",
    "Uganda": "🇺🇬",
    "Ukraine": "🇺🇦",
    "United Arab Emirates": "🇦🇪",
    "United Arab Emirates (UAE)": "🇦🇪",
    "United Kingdom": "🇬🇧",
    "United States": "🇺🇸",
    "United States of America": "🇺🇸",
    "USA": "🇺🇸",
    "UK": "🇬🇧",
    "Uruguay": "🇺🇾",
    "Uzbekistan": "🇺🇿",
    "Venezuela": "🇻🇪",
    "Vietnam": "🇻🇳",
    "Yemen": "🇾🇪",
    "Zimbabwe": "🇿🇼",
}
FLAG_OPTIONS = [(name, emoji) for name, emoji in sorted(FLAG_EMOJIS.items(), key=lambda item: item[0])]
FLAG_ICON_FONT = ("Arial", 14)
FLAG_IMAGE_DIR = os.path.join(APP_DIR, "flags")
BUNDLED_FLAG_IMAGE_DIR = os.path.join(RESOURCE_DIR, "flags")
FLAG_COUNTRY_CODES = {
    "Afghanistan": "af",
    "Albania": "al",
    "Algeria": "dz",
    "Angola": "ao",
    "Argentina": "ar",
    "Armenia": "am",
    "Australia": "au",
    "Austria": "at",
    "Azerbaijan": "az",
    "Bahrain": "bh",
    "Belarus": "by",
    "Belgium": "be",
    "Bolivia": "bo",
    "Bosnia and Herzegovina": "ba",
    "Botswana": "bw",
    "Brazil": "br",
    "Britain": "gb",
    "Bulgaria": "bg",
    "Cambodia": "kh",
    "Cameroon": "cm",
    "Canada": "ca",
    "Chile": "cl",
    "China": "cn",
    "Colombia": "co",
    "Costa Rica": "cr",
    "Croatia": "hr",
    "Cuba": "cu",
    "Czech Republic": "cz",
    "Denmark": "dk",
    "Dominican Republic": "do",
    "Ecuador": "ec",
    "Egypt": "eg",
    "El Salvador": "sv",
    "England": "gb",
    "Estonia": "ee",
    "Ethiopia": "et",
    "Finland": "fi",
    "France": "fr",
    "Gaza": "ps",
    "Georgia": "ge",
    "Germany": "de",
    "Ghana": "gh",
    "Greece": "gr",
    "Guatemala": "gt",
    "Honduras": "hn",
    "Hungary": "hu",
    "Iceland": "is",
    "India": "in",
    "Indonesia": "id",
    "Iran": "ir",
    "Iraq": "iq",
    "Ireland": "ie",
    "Italy": "it",
    "Japan": "jp",
    "Jordan": "jo",
    "Kazakhstan": "kz",
    "Kenya": "ke",
    "Kuwait": "kw",
    "Kyrgyzstan": "kg",
    "Laos": "la",
    "Latvia": "lv",
    "Lebanon": "lb",
    "Lithuania": "lt",
    "Malaysia": "my",
    "Mexico": "mx",
    "Moldova": "md",
    "Mongolia": "mn",
    "Morocco": "ma",
    "Myanmar": "mm",
    "Nepal": "np",
    "Netherlands": "nl",
    "New Zealand": "nz",
    "Nigeria": "ng",
    "Nicaragua": "ni",
    "Norway": "no",
    "Oman": "om",
    "Pakistan": "pk",
    "Palestine": "ps",
    "Panama": "pa",
    "Paraguay": "py",
    "Peru": "pe",
    "Philippines": "ph",
    "Poland": "pl",
    "Portugal": "pt",
    "Qatar": "qa",
    "Romania": "ro",
    "Russia": "ru",
    "Saudi": "sa",
    "Saudi Arabia": "sa",
    "Senegal": "sn",
    "Serbia": "rs",
    "Singapore": "sg",
    "Slovakia": "sk",
    "Slovenia": "si",
    "South Africa": "za",
    "South Korea": "kr",
    "Spain": "es",
    "Sri Lanka": "lk",
    "Sweden": "se",
    "Switzerland": "ch",
    "Syria": "sy",
    "Thailand": "th",
    "Tunisia": "tn",
    "Turkey": "tr",
    "UAE": "ae",
    "Uganda": "ug",
    "Ukraine": "ua",
    "United Arab Emirates": "ae",
    "United Arab Emirates (UAE)": "ae",
    "United Kingdom": "gb",
    "United States": "us",
    "United States of America": "us",
    "USA": "us",
    "UK": "gb",
    "Uruguay": "uy",
    "Uzbekistan": "uz",
    "Venezuela": "ve",
    "Vietnam": "vn",
    "Yemen": "ye",
    "Zimbabwe": "zw",
}
FLAG_IMAGE_CACHE = {}


def normalize_country_name(country_name):
    if not country_name:
        return ""
    cleaned = str(country_name).strip()
    aliases = {
        "us": "United States",
        "usa": "United States",
        "u.s.a.": "United States",
        "united states of america": "United States",
        "united states": "United States",
        "uk": "United Kingdom",
        "united kingdom": "United Kingdom",
        "great britain": "United Kingdom",
        "canada": "Canada",
        "france": "France",
        "china": "China",
        "japan": "Japan",
        "lebanon": "Lebanon",
        "south korea": "South Korea",
        "republic of korea": "South Korea",
        "uae": "United Arab Emirates",
        "united arab emirates": "United Arab Emirates",
        "saudi arabia": "Saudi Arabia",
        "saudi": "Saudi Arabia",
    }
    lowered = cleaned.lower()
    if cleaned in FLAG_COUNTRY_CODES:
        return cleaned
    return aliases.get(lowered, cleaned)


def is_required_holiday_name(name):
    text = re.sub(r"[^a-z0-9]+", " ", str(name or "").lower()).strip()
    if not text:
        return False
    has_independence = "independence" in text
    has_national = "national" in text
    has_day = "day" in text
    return (has_independence and has_day) or (has_national and has_day)


def filter_allowed_holidays(holidays):
    if not holidays:
        return []
    filtered = []
    seen = set()
    for holiday in holidays:
        if not isinstance(holiday, dict):
            continue
        name = str(holiday.get("name", "")).strip()
        date_value = str(holiday.get("date", "")).strip()
        if not name or not date_value:
            continue
        if not is_required_holiday_name(name):
            continue
        key = (name.lower(), date_value.lower())
        if key in seen:
            continue
        seen.add(key)
        filtered.append({"name": name, "date": date_value})
    return filtered


def get_country_code(country_name):
    normalized = normalize_country_name(country_name)
    return FLAG_COUNTRY_CODES.get(normalized, FLAG_COUNTRY_CODES.get(normalized.title(), ""))


def ensure_flag_assets():
    os.makedirs(FLAG_IMAGE_DIR, exist_ok=True)
    for country_name, code in FLAG_COUNTRY_CODES.items():
        flag_path = os.path.join(FLAG_IMAGE_DIR, f"{code}.png")
        if os.path.exists(flag_path):
            continue
        bundled_flag_path = os.path.join(BUNDLED_FLAG_IMAGE_DIR, f"{code}.png")
        if os.path.exists(bundled_flag_path):
            try:
                with open(bundled_flag_path, "rb") as source, open(flag_path, "wb") as target:
                    target.write(source.read())
                continue
            except OSError:
                pass
        url = f"https://flagcdn.com/w80/{code}.png"
        try:
            urlretrieve(url, flag_path)
        except Exception:
            pass


def preload_all_country_flag_images():
    for country_name in list(FLAG_COUNTRY_CODES.keys()):
        try:
            get_flag_photo(country_name)
        except Exception:
            pass


def get_flag_photo(country_name):
    code = get_country_code(country_name)
    if not code:
        return None
    flag_path = os.path.join(FLAG_IMAGE_DIR, f"{code}.png")
    if not os.path.exists(flag_path):
        ensure_flag_assets()
    if not os.path.exists(flag_path):
        return None
    if flag_path not in FLAG_IMAGE_CACHE:
        try:
            FLAG_IMAGE_CACHE[flag_path] = tk.PhotoImage(file=flag_path)
        except Exception:
            return None
    return FLAG_IMAGE_CACHE[flag_path]


# ----------------------------------------------------------------------
#  Data layer
# ----------------------------------------------------------------------
class CountryStore:
    """Handles loading/saving country data to a local JSON file."""

    def __init__(self, path=DATA_FILE):
        self.path = path
        self.countries = []
        self.load()

    @staticmethod
    def _sanitize_country(country):
        if not isinstance(country, dict):
            return {}
        clean = dict(country)
        clean.setdefault("flag", "")
        clean["holidays"] = filter_allowed_holidays(clean.get("holidays", []))
        return clean

    def load(self):
        if os.path.exists(self.path):
            with open(self.path, "r", encoding="utf-8") as f:
                self.countries = json.load(f)
        elif os.path.exists(BUNDLED_DATA_FILE):
            with open(BUNDLED_DATA_FILE, "r", encoding="utf-8") as f:
                self.countries = json.load(f)
            self.save()
        else:
            # Starter example so the app isn't empty on first run
            self.countries = [
                {
                    "name": "Lebanon",
                    "president": "Joseph Aoun",
                    "minister_of_defence": "Michel Menassa",
                    "ambassador_to_lebanon": "N/A (home country)",
                    "flag": "🇱🇧",
                    "holidays": [
                        {"name": "Independence Day", "date": "11-22"}
                    ],
                },
                {
                    "name": "USA",
                    "president": "Donald Trump",
                    "minister_of_defence": "Pete Hegseth",
                    "ambassador_to_lebanon": "",
                    "flag": "🇺🇸",
                    "holidays": [
                        {"name": "Independence Day", "date": "07-04"}
                    ],
                },
                {
                    "name": "Canada",
                    "president": "Mark Carney",
                    "minister_of_defence": "Bill Blair",
                    "ambassador_to_lebanon": "",
                    "flag": "🇨🇦",
                    "holidays": [
                        {"name": "Independence Day", "date": "07-01"}
                    ],
                },
                {
                    "name": "France",
                    "president": "Emmanuel Macron",
                    "minister_of_defence": "Sébastien Lecornu",
                    "ambassador_to_lebanon": "",
                    "flag": "🇫🇷",
                    "holidays": [
                        {"name": "Independence Day", "date": "07-14"}
                    ],
                },
                {
                    "name": "UK",
                    "president": "King Charles III",
                    "minister_of_defence": "John Healey",
                    "ambassador_to_lebanon": "",
                    "flag": "🇬🇧",
                    "holidays": [
                        {"name": "Independence Day", "date": "04-23"}
                    ],
                },
                {
                    "name": "China",
                    "president": "Xi Jinping",
                    "minister_of_defence": "Dong Jun",
                    "ambassador_to_lebanon": "",
                    "flag": "🇨🇳",
                    "holidays": [
                        {"name": "National Day", "date": "10-01"}
                    ],
                },
                {
                    "name": "Japan",
                    "president": "Fumio Kishida",
                    "minister_of_defence": "Gen Nakatani",
                    "ambassador_to_lebanon": "",
                    "flag": "🇯🇵",
                    "holidays": [
                        {"name": "Independence Day", "date": "02-11"}
                    ],
                },
            ]
        self.countries = [self._sanitize_country(c) for c in self.countries if isinstance(c, dict)]
        self.save()

    def save(self):
        sanitized = [self._sanitize_country(c) for c in self.countries if isinstance(c, dict)]
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(sanitized, f, indent=2, ensure_ascii=False)

    def get(self, name):
        for c in self.countries:
            if c["name"] == name:
                return c
        return None

    def upsert(self, country):
        if not isinstance(country, dict):
            return
        country = self._sanitize_country(country)
        existing = self.get(country["name"])
        if existing:
            existing.update(country)
            existing["holidays"] = filter_allowed_holidays(existing.get("holidays", []))
        else:
            self.countries.append(country)
        self.save()

    def delete(self, name):
        self.countries = [c for c in self.countries if c["name"] != name]
        self.save()

    def rename(self, old_name, new_name):
        c = self.get(old_name)
        if c:
            c["name"] = new_name
            self.save()


# ----------------------------------------------------------------------
#  Date helpers
# ----------------------------------------------------------------------
def next_occurrence(mm_dd, today=None):
    """Given a 'MM-DD' string, return the next date() this occurs on
    (today or in the future), rolling over to next year if needed."""
    today = today or date.today()
    try:
        month, day = map(int, mm_dd.split("-"))
    except (ValueError, AttributeError):
        return None
    for year in (today.year, today.year + 1):
        try:
            d = date(year, month, day)
        except ValueError:
            continue  # e.g. invalid Feb 29 on a non-leap year
        if d >= today:
            return d
    return None


def days_until(d):
    return (d - date.today()).days


def normalize_excel_date(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        value = value.strftime("%m-%d")
    text = str(value).strip()
    if not text:
        return ""
    text = text.replace("/", "-").replace(".", "-")
    match = re.search(r"(\d{1,2})\s*-\s*(\d{1,2})", text)
    if match:
        month, day = match.groups()
        return f"{int(month):02d}-{int(day):02d}"
    try:
        dt = datetime.strptime(text, "%Y-%m-%d")
        return dt.strftime("%m-%d")
    except ValueError:
        pass
    try:
        dt = datetime.strptime(text, "%m/%d/%Y")
        return dt.strftime("%m-%d")
    except ValueError:
        pass
    return text


def import_excel_data(file_path, store):
    if load_workbook is None:
        raise RuntimeError("openpyxl is not installed. Install it with: pip install openpyxl")

    wb = load_workbook(file_path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        return 0

    raw_headers = [str(v).strip() if v is not None else "" for v in rows[0]]
    headers = [re.sub(r"[^a-z0-9]+", "_", h.lower()).strip("_") for h in raw_headers]

    def idx_of(*names):
        for name in names:
            for i, h in enumerate(headers):
                if h == name:
                    return i
        return None

    country_idx = idx_of("country", "name", "country_name")
    president_idx = idx_of("president", "president_name")
    minister_idx = idx_of("minister_of_defence", "minister_of_defense", "minister", "defence_minister", "defense_minister")
    ambassador_idx = idx_of("ambassador_to_lebanon", "ambassador", "ambassador_name")
    embassy_location_idx = idx_of("embassy_location", "embassy_address", "embassy_location_address", "location")
    embassy_phone_idx = idx_of("embassy_phone", "embassy_telephone", "embassy_tel", "phone", "telephone")
    holiday_idx = idx_of("holiday_name", "holiday", "event_name", "name")
    date_idx = idx_of("date", "holiday_date", "date_mm_dd", "independence_day")

    if country_idx is None:
        if len(headers) >= 6:
            country_idx = 0
            president_idx = 1 if len(headers) > 1 else None
            minister_idx = 2 if len(headers) > 2 else None
            ambassador_idx = 3 if len(headers) > 3 else None
            holiday_idx = 4 if len(headers) > 4 else None
            date_idx = 5 if len(headers) > 5 else None
        else:
            raise ValueError("The Excel sheet is missing required columns like Country and Date.")

    imported = {}
    for row in rows[1:]:
        if not row or all(v is None or str(v).strip() == "" for v in row):
            continue

        country_name = str(row[country_idx]).strip() if country_idx is not None and country_idx < len(row) and row[country_idx] is not None else ""
        if not country_name:
            continue

        if country_name not in imported:
            imported[country_name] = {
                "name": country_name,
                "president": "",
                "minister_of_defence": "",
                "ambassador_to_lebanon": "",
                "embassy_location": "",
                "embassy_phone": "",
                "holidays": [],
            }

        country = imported[country_name]
        if president_idx is not None and president_idx < len(row) and row[president_idx] is not None:
            country["president"] = str(row[president_idx]).strip()
        if minister_idx is not None and minister_idx < len(row) and row[minister_idx] is not None:
            country["minister_of_defence"] = str(row[minister_idx]).strip()
        if ambassador_idx is not None and ambassador_idx < len(row) and row[ambassador_idx] is not None:
            country["ambassador_to_lebanon"] = str(row[ambassador_idx]).strip()
        if embassy_location_idx is not None and embassy_location_idx < len(row) and row[embassy_location_idx] is not None:
            country["embassy_location"] = str(row[embassy_location_idx]).strip()
        if embassy_phone_idx is not None and embassy_phone_idx < len(row) and row[embassy_phone_idx] is not None:
            country["embassy_phone"] = str(row[embassy_phone_idx]).strip()

        if holiday_idx is not None and date_idx is not None and holiday_idx < len(row) and date_idx < len(row):
            holiday_name = row[holiday_idx]
            holiday_date = row[date_idx]
            if holiday_name is not None and str(holiday_name).strip() and holiday_date is not None and str(holiday_date).strip():
                normalized_date = normalize_excel_date(holiday_date)
                if normalized_date:
                    country["holidays"].append({
                        "name": str(holiday_name).strip(),
                        "date": normalized_date,
                    })

    for country_name, imported_country in imported.items():
        imported_country["holidays"] = filter_allowed_holidays(imported_country.get("holidays", []))
        existing = store.get(country_name)
        if existing:
            if not existing.get("president") and imported_country.get("president"):
                existing["president"] = imported_country["president"]
            if not existing.get("minister_of_defence") and imported_country.get("minister_of_defence"):
                existing["minister_of_defence"] = imported_country["minister_of_defence"]
            if not existing.get("ambassador_to_lebanon") and imported_country.get("ambassador_to_lebanon"):
                existing["ambassador_to_lebanon"] = imported_country["ambassador_to_lebanon"]
            if not existing.get("embassy_location") and imported_country.get("embassy_location"):
                existing["embassy_location"] = imported_country["embassy_location"]
            if not existing.get("embassy_phone") and imported_country.get("embassy_phone"):
                existing["embassy_phone"] = imported_country["embassy_phone"]
            existing_holidays = {h["name"].lower(): h for h in existing.get("holidays", [])}
            for h in imported_country["holidays"]:
                if h["name"].lower() not in existing_holidays:
                    existing.setdefault("holidays", []).append(h)
        else:
            store.countries.append(imported_country)

    store.save()
    return len(imported)


def export_excel_data(file_path, store):
    if Workbook is None:
        raise RuntimeError("openpyxl is not installed. Install it with: pip install openpyxl")

    wb = Workbook()
    ws = wb.active
    ws.title = "Countries"
    ws.append([
        "Country",
        "President",
        "Minister of Defence",
        "Ambassador to Lebanon",
        "Embassy Location",
        "Embassy Phone",
        "Holiday Name",
        "Date (MM-DD)",
    ])

    for country in store.countries:
        holidays = country.get("holidays", [])
        if not holidays:
            ws.append([
                country.get("name", ""),
                country.get("president", ""),
                country.get("minister_of_defence", ""),
                country.get("ambassador_to_lebanon", ""),
                country.get("embassy_location", ""),
                country.get("embassy_phone", ""),
                "",
                "",
            ])
            continue

        for holiday in holidays:
            ws.append([
                country.get("name", ""),
                country.get("president", ""),
                country.get("minister_of_defence", ""),
                country.get("ambassador_to_lebanon", ""),
                country.get("embassy_location", ""),
                country.get("embassy_phone", ""),
                holiday.get("name", ""),
                holiday.get("date", ""),
            ])

    ws.freeze_panes = "A2"
    for col in ["A", "B", "C", "D", "E", "F", "G", "H"]:
        ws.column_dimensions[col].width = 24
    wb.save(file_path)
    return ws.max_row - 1


def next_holiday_for_country(country):
    """Return (holiday_dict, date_obj) for the soonest upcoming holiday, or (None, None)."""
    best = None
    best_date = None
    for h in country.get("holidays", []):
        d = next_occurrence(h.get("date", ""))
        if d and (best_date is None or d < best_date):
            best, best_date = h, d
    return best, best_date


def fetch_json(url):
    request = Request(url, headers={"User-Agent": "IndependenceCalendar/1.0"})
    with urlopen(request, timeout=20) as response:
        return json.loads(response.read().decode("utf-8"))


def country_flag_from_code(code):
    return "".join(chr(127397 + ord(letter)) for letter in code.upper() if "A" <= letter <= "Z")


# ----------------------------------------------------------------------
#  Main Application
# ----------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.language = "en"
        self.store = CountryStore()
        ensure_flag_assets()
        preload_all_country_flag_images()
        self.notify_days = tk.IntVar(value=DEFAULT_NOTIFY_DAYS)
        self._tray_icon = None
        self._background_mode = True

        self.bind("<Control-Shift-D>", lambda event: self._prompt_developer_access())
        self.protocol("WM_DELETE_WINDOW", self._handle_close_requested)

        self.title(self.t("app_title"))
        self.geometry("900x560")

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        self.main_tab = MainTab(notebook, self)
        self.edit_tab = EditTab(notebook, self)

        notebook.add(self.main_tab, text=self.t("main_panel"))
        notebook.add(self.edit_tab, text=self.t("edit_panel"))

        self.refresh_all()

        self._create_tray_icon()
        self.after(250, self._start_background_mode)

        # Check for upcoming holidays shortly after startup, then hourly.
        self.after(500, self.check_notifications)
        self.after(60 * 60 * 1000, self._notification_loop)

    def _build_tray_icon_image(self):
        if Image is None or ImageDraw is None:
            return None

        img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        draw.rounded_rectangle((8, 8, 56, 56), radius=10, fill=(29, 78, 216, 255))
        draw.rounded_rectangle((16, 16, 48, 48), radius=8, fill=(255, 255, 255, 255))
        draw.line((20, 30, 44, 30), fill=(29, 78, 216, 255), width=4)
        draw.line((32, 18, 32, 46), fill=(29, 78, 216, 255), width=4)
        draw.line((20, 44, 44, 20), fill=(255, 94, 94, 255), width=4)
        return img

    def _create_tray_icon(self):
        if pystray is None:
            return

        if self._tray_icon is not None:
            return

        icon_image = self._build_tray_icon_image()

        def show_hide(_):
            self._toggle_main_window()

        def check_now(_):
            self.check_notifications(silent_if_none=False)

        menu = (
            pystray.MenuItem("Show / Hide", show_hide),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Check Now", check_now),
            pystray.MenuItem("Exit", self._exit_app),
        )

        if icon_image is None:
            icon_image = "🗓"

        self._tray_icon = pystray.Icon("IndependenceCalendar", icon_image, "Independence & Holiday Calendar", menu=menu)

        def runner():
            self._tray_icon.run()

        threading.Thread(target=runner, daemon=True).start()

    def _toggle_main_window(self):
        if self.winfo_viewable():
            self.withdraw()
        else:
            self.deiconify()
            self.update_idletasks()
            self.focus_force()

    def _start_background_mode(self):
        if self._background_mode and self._tray_icon is not None:
            self.withdraw()

    def _handle_close_requested(self):
        if self._tray_icon is not None:
            self.withdraw()
            if hasattr(self._tray_icon, "notify"):
                try:
                    self._tray_icon.notify("App is running in the background.")
                except Exception:
                    pass
            return
        self.destroy()

    def _exit_app(self, icon=None, item=None):
        if self._tray_icon is not None:
            try:
                self._tray_icon.stop()
            except Exception:
                pass
        self.destroy()

    def _notification_loop(self):
        self.check_notifications()
        self.after(60 * 60 * 1000, self._notification_loop)

    def _prompt_developer_access(self):
        password = simpledialog.askstring("Developer Access", "Enter developer password:", show="*")
        if password != DEVELOPER_PASSWORD:
            messagebox.showerror("Access denied", "Incorrect password.")
            return
        self._open_developer_panel()

    def _open_developer_panel(self):
        if getattr(self, "_dev_window", None) is not None and self._dev_window.winfo_exists():
            self._dev_window.focus_set()
            return

        win = tk.Toplevel(self)
        win.title("Developer Access")
        win.geometry("440x300")
        win.transient(self)
        win.grab_set()
        self._dev_window = win

        ttk.Label(win, text="Developer Controls", font=("", 11, "bold")).pack(pady=(18, 12))

        actions = [
            ("Open countries.json", self._open_data_file),
            ("Open flags folder", self._open_flags_folder),
            ("Show app status", self._show_app_status),
            ("Backup data", self._backup_data_file),
            ("Reset app data", self._reset_app_data),
        ]

        for label, action in actions:
            ttk.Button(win, text=label, command=action, width=22).pack(pady=5)

        ttk.Button(win, text="Close", command=win.destroy).pack(pady=12)

    def _open_data_file(self):
        if os.path.exists(DATA_FILE):
            try:
                os.startfile(DATA_FILE)
            except Exception:
                messagebox.showinfo("Developer file", DATA_FILE)
        else:
            messagebox.showinfo("Developer file", "countries.json has not been created yet.")

    def _open_flags_folder(self):
        folder = FLAG_IMAGE_DIR
        os.makedirs(folder, exist_ok=True)
        try:
            os.startfile(folder)
        except Exception:
            messagebox.showinfo("Flags folder", folder)

    def _show_app_status(self):
        countries = len(self.store.countries)
        msg = (
            f"Country count: {countries}\n"
            f"Notification threshold: {self.notify_days.get()} days\n"
            f"Language: {'Arabic' if self.language == 'ar' else 'English'}\n"
            f"Data file: {DATA_FILE}"
        )
        messagebox.showinfo("App Status", msg)

    def _backup_data_file(self):
        if not os.path.exists(DATA_FILE):
            messagebox.showinfo("Backup", "No data file exists yet.")
            return
        backup_path = os.path.splitext(DATA_FILE)[0] + ".backup.json"
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as src, open(backup_path, "w", encoding="utf-8") as dst:
                dst.write(src.read())
            messagebox.showinfo("Backup created", f"Saved to:\n{backup_path}")
        except Exception as exc:
            messagebox.showerror("Backup failed", str(exc))

    def _reset_app_data(self):
        if not messagebox.askyesno("Reset app data", "Reset all saved country data and start fresh?"):
            return
        if os.path.exists(DATA_FILE):
            os.remove(DATA_FILE)
        self.store = CountryStore()
        self.refresh_all()
        if getattr(self, "_dev_window", None) is not None and self._dev_window.winfo_exists():
            self._dev_window.destroy()
        messagebox.showinfo("Developer reset", "App data has been reset.")

    def t(self, key, **kwargs):
        lang = TRANSLATIONS.get(self.language, TRANSLATIONS["en"])
        text = lang.get(key, TRANSLATIONS["en"].get(key, key))
        return text.format(**kwargs) if kwargs else text

    def set_language(self, code):
        if code not in TRANSLATIONS:
            code = "en"
        self.language = code
        self.title(self.t("app_title"))
        if hasattr(self, "main_tab"):
            self.main_tab.apply_language()
        if hasattr(self, "edit_tab"):
            self.edit_tab.apply_language()

    def get_country_flag(self, country_name):
        if not country_name:
            return "🏳️"
        if isinstance(country_name, dict):
            country_name = country_name.get("name", "")
        country_name = str(country_name).strip()
        return FLAG_EMOJIS.get(country_name, "🏳️")

    def get_country_flag_image(self, country_name):
        return get_flag_photo(country_name)

    def _resize_flag_for_notification(self, flag_photo, max_size=24):
        if flag_photo is None:
            return None
        width = flag_photo.width()
        height = flag_photo.height()
        if width <= max_size and height <= max_size:
            return flag_photo
        scale = max(1, max(width // max_size, height // max_size))
        try:
            return flag_photo.subsample(scale)
        except Exception:
            return flag_photo

    def refresh_all(self):
        self.main_tab.refresh()
        self.edit_tab.refresh_country_list()

    def import_excel_data_from_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel file to import",
            filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xls"), ("All files", "*.*")],
        )
        if not file_path:
            return
        try:
            imported_count = import_excel_data(file_path, self.store)
            self.refresh_all()
            messagebox.showinfo(self.t("import_complete"), f"Imported {imported_count} country entries from the Excel file.")
        except Exception as exc:
            messagebox.showerror(self.t("import_failed"), f"Could not import the Excel file.\n{exc}")

    def export_excel_data_to_file(self):
        file_path = filedialog.asksaveasfilename(
            title="Save as Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not file_path:
            return
        try:
            exported_rows = export_excel_data(file_path, self.store)
            messagebox.showinfo(self.t("export_complete"), f"Exported {exported_rows} rows to the Excel file.")
        except Exception as exc:
            messagebox.showerror(self.t("export_failed"), f"Could not export to Excel.\n{exc}")

    def _show_event_popup(self, event, upcoming=None, event_index=0):
        if getattr(self, "_event_popup", None) is not None and self._event_popup.winfo_exists():
            self._event_popup.destroy()

        upcoming = upcoming or self.get_upcoming_notifications()
        name, holiday, event_date, days_left = event
        popup = tk.Toplevel(self)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.configure(bg="#2b3747")
        self._event_popup = popup

        width, height = 440, 178
        screen_width = popup.winfo_screenwidth()
        screen_height = popup.winfo_screenheight()
        popup.geometry(f"{width}x{height}+{screen_width - width - 24}+{screen_height - height - 56}")

        content = tk.Frame(popup, bg="#2b3747", padx=16, pady=12)
        content.pack(fill="both", expand=True)
        content.config(highlightbackground="#4b5563", highlightthickness=1)

        title = tk.Label(
            content,
            text=self.t("upcoming_holidays"),
            bg="#2b3747",
            fg="#93c5fd",
            font=("Segoe UI", 10, "bold"),
            anchor="w",
        )
        title.pack(fill="x")

        close_button = tk.Button(
            popup,
            text="X",
            command=popup.destroy,
            bg="#2b3747",
            fg="#d1d5db",
            activebackground="#374151",
            activeforeground="white",
            bd=0,
            highlightthickness=0,
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
        )
        close_button.place(relx=1.0, x=-10, y=8, anchor="ne")

        event_label_text = f"{name} — {holiday}"
        fallback_flag = self.get_country_flag(name)
        flag_photo = self.get_country_flag_image(name)
        header_row = tk.Frame(content, bg="#2b3747")
        header_row.pack(fill="x", pady=(4, 0))

        flag_label = tk.Label(
            header_row,
            text=fallback_flag,
            bg="#2b3747",
            fg="white",
            font=("Segoe UI Emoji", 20, "bold"),
            anchor="center",
            justify="center",
            width=4,
            height=2,
            relief="solid",
            bd=1,
            highlightbackground="#4b5563",
        )
        flag_label.pack(side="left", padx=(0, 8), pady=2)

        if flag_photo is not None:
            flag_photo = self._resize_flag_for_notification(flag_photo, max_size=32)
            flag_label.config(image=flag_photo, compound="center", text="", width=4, height=2)
            flag_label.image = flag_photo

        event_label = tk.Label(
            header_row,
            text=event_label_text,
            bg="#2b3747",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            anchor="w",
            justify="left",
        )
        event_label.pack(side="left", fill="x", expand=True)

        when = self.t("today") if days_left == 0 else self.t("in_days", n=days_left)
        details = tk.Label(
            content,
            text=f"{event_date.strftime('%b %d, %Y')}  •  {when}",
            bg="#2b3747",
            fg="#d1d5db",
            font=("Segoe UI", 9),
            anchor="w",
        )
        details.pack(fill="x")

        actions = tk.Frame(content, bg="#2b3747")
        actions.pack(fill="x", pady=(10, 0))

        previous_button = tk.Button(
            actions,
            text=self.t("previous"),
            command=lambda: self._show_event_popup(upcoming[event_index - 1], upcoming, event_index - 1),
            state="normal" if event_index > 0 else "disabled",
            bg="#374151",
            fg="white",
            activebackground="#4b5563",
            activeforeground="white",
            bd=0,
            padx=12,
            pady=4,
            cursor="hand2",
        )
        previous_button.pack(side="left")

        next_button = tk.Button(
            actions,
            text=self.t("next"),
            command=lambda: self._show_event_popup(upcoming[event_index + 1], upcoming, event_index + 1),
            state="normal" if event_index + 1 < len(upcoming) else "disabled",
            bg="#2563eb",
            fg="white",
            activebackground="#1d4ed8",
            activeforeground="white",
            bd=0,
            padx=12,
            pady=4,
            cursor="hand2",
        )
        next_button.pack(side="right")

        done_button = tk.Button(
            actions,
            text=self.t("done"),
            command=popup.destroy,
            bg="#374151",
            fg="white",
            activebackground="#4b5563",
            activeforeground="white",
            bd=0,
            padx=12,
            pady=4,
            cursor="hand2",
        )
        done_button.pack(side="right", padx=(0, 8))

        for widget in (popup, content, title, event_label, details):
            widget.bind("<Button-1>", lambda _event: self._toggle_main_window())

    def check_notifications(self, silent_if_none=True):
        upcoming = self.get_upcoming_notifications()

        if upcoming:
            self._show_event_popup(upcoming[0])
        elif not silent_if_none:
            messagebox.showinfo(self.t("upcoming_holidays"), self.t("no_holidays_window"))

    def get_upcoming_notifications(self):
        threshold = self.notify_days.get()
        upcoming = []
        for c in self.store.countries:
            h, d = next_holiday_for_country(c)
            if h and 0 <= days_until(d) <= threshold:
                upcoming.append((c["name"], h["name"], d, days_until(d)))

        upcoming.sort(key=lambda x: x[3])
        return upcoming


# ----------------------------------------------------------------------
#  Main Panel Tab
# ----------------------------------------------------------------------
class MainTab(ttk.Frame):
    def __init__(self, parent, app: App):
        super().__init__(parent)
        self.app = app
        self.flag_images = {}
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)

        self.lang_var = tk.StringVar(value="English" if self.app.language == "en" else "العربية")
        ttk.Label(top, text=self.app.t("language")).pack(side="left")
        self.lang_combo = ttk.Combobox(top, textvariable=self.lang_var, state="readonly", width=12)
        self.lang_combo["values"] = ["English", "العربية"]
        self.lang_combo.pack(side="left", padx=(5, 10))
        self.lang_combo.bind("<<ComboboxSelected>>", lambda e: self.app.set_language("ar" if self.lang_var.get() == "العربية" else "en"))

        self.notify_label = ttk.Label(top, text=self.app.t("notify_me"))
        self.notify_label.pack(side="left")
        spin = ttk.Spinbox(top, from_=1, to=90, width=5, textvariable=self.app.notify_days)
        spin.pack(side="left", padx=5)
        self.days_label = ttk.Label(top, text=self.app.t("days"))
        self.days_label.pack(side="left")

        self.check_button = ttk.Button(top, text=self.app.t("check_now"), command=lambda: self.app.check_notifications(silent_if_none=False))
        self.check_button.pack(side="left", padx=15)
        self.next_notification_button = ttk.Button(top, command=self.show_next_notification)
        self.next_notification_button.pack(side="left", padx=(0, 10))
        self.refresh_button = ttk.Button(top, text=self.app.t("refresh"), command=self.refresh)
        self.refresh_button.pack(side="left")
        self.import_button = ttk.Button(top, text=self.app.t("import_excel"), command=self.app.import_excel_data_from_file)
        self.import_button.pack(side="left", padx=10)
        self.export_button = ttk.Button(top, text=self.app.t("export_excel"), command=self.app.export_excel_data_to_file)
        self.export_button.pack(side="left", padx=10)

        search_row = ttk.Frame(self)
        search_row.pack(fill="x", padx=10, pady=(0, 8))
        self.search_label = ttk.Label(search_row, text=self.app.t("search_country"))
        self.search_label.pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(search_row, textvariable=self.search_var, width=30).pack(side="left", padx=5)

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=10, pady=5)

        # Left: list of countries with next holiday
        left = ttk.Frame(body)
        left.pack(side="left", fill="both", expand=True)

        columns = ("country", "holiday", "date", "days_left")
        self.tree = ttk.Treeview(left, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("country", text=self.app.t("country"))
        self.tree.heading("holiday", text=self.app.t("next_holiday"))
        self.tree.heading("date", text=self.app.t("date"))
        self.tree.heading("days_left", text=self.app.t("days_left"))
        self.tree.column("country", width=150)
        self.tree.column("holiday", width=180)
        self.tree.column("date", width=100, anchor="center")
        self.tree.column("days_left", width=90, anchor="center")
        self.tree.pack(fill="both", expand=True, side="left")

        scroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        scroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)

        self.tree.tag_configure("soon", background="#ffe9a8")
        self.tree.tag_configure("today", background="#ffb3b3")

        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        # Right: details panel
        right = ttk.LabelFrame(body, text=self.app.t("country_details"), padding=12)
        right.pack(side="left", fill="y", padx=(10, 0))

        self.flag_frame = ttk.LabelFrame(right, text=self.app.t("flag"), padding=12)
        self.flag_frame.pack(fill="x", pady=(0, 10))
        self.flag_label = tk.Label(self.flag_frame, text="🏳️", font=("Segoe UI Emoji", 54, "bold"), compound="center", anchor="center")
        self.flag_label.pack(fill="x", expand=True, pady=8)

        self.detail_name = self._detail_row(right, self.app.t("country_label"))
        self.detail_president = self._detail_row(right, self.app.t("president_label"))
        self.detail_minister = self._detail_row(right, self.app.t("minister_label"))
        self.detail_ambassador = self._detail_row(right, self.app.t("ambassador_label"))

        self.embassy_tabs = ttk.Notebook(right)
        self.embassy_tabs.pack(fill="x", pady=(8, 3))
        location_tab = ttk.Frame(self.embassy_tabs, padding=6)
        phone_tab = ttk.Frame(self.embassy_tabs, padding=6)
        self.embassy_tabs.add(location_tab, text=self.app.t("embassy_location_tab"))
        self.embassy_tabs.add(phone_tab, text=self.app.t("embassy_phone_tab"))
        self.detail_embassy_location = self._detail_value(location_tab)
        self.detail_embassy_phone = self._detail_value(phone_tab)

        ttk.Label(right, text=self.app.t("holidays"), font=("", 10, "bold")).pack(anchor="w", pady=(10, 2))
        self.holiday_list = tk.Listbox(right, width=38, height=10)
        self.holiday_list.pack()

    def apply_language(self):
        self.lang_var.set("English" if self.app.language == "en" else "العربية")
        self.notify_label.config(text=self.app.t("notify_me"))
        self.days_label.config(text=self.app.t("days"))
        self.check_button.config(text=self.app.t("check_now"))
        self.update_next_notification_button()
        self.refresh_button.config(text=self.app.t("refresh"))
        self.import_button.config(text=self.app.t("import_excel"))
        self.export_button.config(text=self.app.t("export_excel"))
        self.search_label.config(text=self.app.t("search_country"))
        self.tree.heading("country", text=self.app.t("country"))
        self.tree.heading("holiday", text=self.app.t("next_holiday"))
        self.tree.heading("date", text=self.app.t("date"))
        self.tree.heading("days_left", text=self.app.t("days_left"))
        self.flag_frame.config(text=self.app.t("flag"))
        self.detail_name.config(text=self.app.t("country_label"))
        self.detail_president.config(text=self.app.t("president_label"))
        self.detail_minister.config(text=self.app.t("minister_label"))
        self.detail_ambassador.config(text=self.app.t("ambassador_label"))
        self.embassy_tabs.tab(0, text=self.app.t("embassy_location_tab"))
        self.embassy_tabs.tab(1, text=self.app.t("embassy_phone_tab"))
        self.holiday_listbox = getattr(self, "holiday_list", None)
        if self.holiday_listbox is not None:
            pass
        self.root = getattr(self, "master", None)
        self.master = getattr(self, "master", None)
        self._lang_frame = getattr(self, "_lang_frame", None)
        self._lang_frame = self

    def _detail_row(self, parent, label_text):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=3)
        ttk.Label(row, text=label_text, width=20, anchor="w").pack(side="left")
        value = ttk.Label(row, text="-", anchor="w", wraplength=200, font=("", 9, "bold"))
        value.pack(side="left")
        return value

    def _detail_value(self, parent):
        value = ttk.Label(parent, text="-", anchor="w", wraplength=220, font=("", 9, "bold"))
        value.pack(fill="x")
        return value

    def update_next_notification_button(self):
        upcoming = self.app.get_upcoming_notifications()
        if not upcoming:
            self.next_notification_button.config(text=self.app.t("no_next_notification"), state="disabled")
            return

        country_name, holiday_name, event_date, _days_left = upcoming[0]
        self.next_notification_button.config(
            text=self.app.t(
                "next_notification",
                name=f"{holiday_name} ({country_name})",
                date=event_date.strftime("%b %d"),
            ),
            state="normal",
        )

    def show_next_notification(self):
        upcoming = self.app.get_upcoming_notifications()
        if upcoming:
            self.app._show_event_popup(upcoming[0])

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        self.flag_images.clear()
        self.update_next_notification_button()
        threshold = self.app.notify_days.get()
        search_text = (self.search_var.get() or "").strip().lower()

        for c in self.app.store.countries:
            if search_text and search_text not in c["name"].lower():
                continue

            flag_photo = self.app.get_country_flag_image(c["name"])
            if flag_photo is not None:
                self.flag_images[c["name"]] = flag_photo

            h, d = next_holiday_for_country(c)
            if h:
                left = days_until(d)
                tag = ()
                if left == 0:
                    tag = ("today",)
                elif left <= threshold:
                    tag = ("soon",)
                country_label = c["name"]
                if flag_photo is None:
                    country_label = f"{c.get('flag') or self.app.get_country_flag(c['name'])} {c['name']}"
                insert_kw = {
                    "iid": c["name"],
                    "values": (country_label, h["name"], d.strftime("%b %d, %Y"), left),
                    "tags": tag,
                }
                if flag_photo is not None:
                    insert_kw["image"] = flag_photo
                self.tree.insert("", "end", **insert_kw)
            else:
                country_label = c["name"]
                if flag_photo is None:
                    country_label = f"{c.get('flag') or self.app.get_country_flag(c['name'])} {c['name']}"
                insert_kw = {
                    "iid": c["name"],
                    "values": (country_label, "—", "—", "—"),
                }
                if flag_photo is not None:
                    insert_kw["image"] = flag_photo
                self.tree.insert("", "end", **insert_kw)

        if not self.tree.get_children():
            self.clear_details()
            return

        first_country = self.tree.get_children()[0]
        self.tree.selection_set(first_country)
        self.on_select(None)

    def clear_details(self):
        for lbl in (
            self.detail_name,
            self.detail_president,
            self.detail_minister,
            self.detail_ambassador,
            self.detail_embassy_location,
            self.detail_embassy_phone,
        ):
            lbl.config(text="-")
        self.flag_label.config(text="🏳️", image="", font=("Segoe UI Emoji", 54, "bold"), compound="center", anchor="center")
        self.holiday_list.delete(0, "end")

    def on_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        c = self.app.store.get(sel[0])
        if not c:
            return
        self.detail_name.config(text=c["name"])
        self.detail_president.config(text=c.get("president") or "-")
        self.detail_minister.config(text=c.get("minister_of_defence") or "-")
        self.detail_ambassador.config(text=c.get("ambassador_to_lebanon") or "-")
        self.detail_embassy_location.config(text=c.get("embassy_location") or "-")
        self.detail_embassy_phone.config(text=c.get("embassy_phone") or "-")
        flag_image = self.app.get_country_flag_image(c["name"])
        if flag_image:
            self.flag_label.config(image=flag_image, compound="center", text="", font=("Segoe UI Emoji", 54, "bold"), anchor="center")
            self.flag_label.image = flag_image
        else:
            self.flag_label.config(text=c.get("flag") or self.app.get_country_flag(c["name"]), image="", font=("Segoe UI Emoji", 54, "bold"), compound="center", anchor="center")

        self.holiday_list.delete(0, "end")
        for h in c.get("holidays", []):
            d = next_occurrence(h.get("date", ""))
            date_str = d.strftime("%b %d, %Y") if d else h.get("date", "?")
            self.holiday_list.insert("end", f"{h['name']}  —  {date_str}")


# ----------------------------------------------------------------------
#  Add / Edit Tab
# ----------------------------------------------------------------------
class EditTab(ttk.Frame):
    def __init__(self, parent, app: App):
        super().__init__(parent)
        self.app = app
        self.current_holidays = []  # working copy of holidays for the country being edited
        self.editing_original_name = None  # tracks name if editing existing country
        self._build()

    def _build(self):
        # --- pick existing country to edit, or start new ---
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)

        self.select_country_label = ttk.Label(top, text=self.app.t("select_country_to_edit"))
        self.select_country_label.pack(side="left")
        self.country_picker = ttk.Combobox(top, state="readonly", width=25)
        self.country_picker.pack(side="left", padx=5)
        self.country_picker.bind("<<ComboboxSelected>>", self.load_selected_country)

        self.new_country_button = ttk.Button(top, text=self.app.t("new_country"), command=self.new_country)
        self.new_country_button.pack(side="left", padx=10)
        self.delete_button = ttk.Button(top, text=self.app.t("delete_country"), command=self.delete_country)
        self.delete_button.pack(side="left")
        self.import_button = ttk.Button(top, text=self.app.t("import_excel"), command=self.app.import_excel_data_from_file)
        self.import_button.pack(side="left", padx=10)
        self.export_button = ttk.Button(top, text=self.app.t("export_excel"), command=self.app.export_excel_data_to_file)
        self.export_button.pack(side="left", padx=10)

        # --- form fields ---
        form = ttk.LabelFrame(self, text="Country Info", padding=12)
        form.pack(fill="x", padx=10, pady=8)

        self.name_var = tk.StringVar()
        self.president_var = tk.StringVar()
        self.minister_var = tk.StringVar()
        self.ambassador_var = tk.StringVar()
        self.embassy_location_var = tk.StringVar()
        self.embassy_phone_var = tk.StringVar()
        self.flag_var = tk.StringVar()

        self._form_row(form, self.app.t("country_name"), self.name_var)
        self._form_row(form, self.app.t("president"), self.president_var)
        self._form_row(form, self.app.t("minister_of_defence"), self.minister_var)
        self._form_row(form, self.app.t("ambassador_to_lebanon"), self.ambassador_var)

        embassy_tabs = ttk.Notebook(form)
        embassy_tabs.pack(fill="x", pady=(8, 3))
        location_tab = ttk.Frame(embassy_tabs, padding=8)
        phone_tab = ttk.Frame(embassy_tabs, padding=8)
        embassy_tabs.add(location_tab, text=self.app.t("embassy_location_tab"))
        embassy_tabs.add(phone_tab, text=self.app.t("embassy_phone_tab"))
        self.embassy_tabs = embassy_tabs
        self._form_row(location_tab, self.app.t("embassy_location_label"), self.embassy_location_var)
        self._form_row(phone_tab, self.app.t("embassy_phone_label"), self.embassy_phone_var)

        flag_row = ttk.Frame(form)
        flag_row.pack(fill="x", pady=3)
        self.flag_label = ttk.Label(flag_row, text=self.app.t("flag"), width=22, anchor="w")
        self.flag_label.pack(side="left")
        flag_values = [f"{emoji} {name}" for name, emoji in FLAG_OPTIONS]
        self.flag_combo = ttk.Combobox(flag_row, textvariable=self.flag_var, values=flag_values, state="readonly", width=28)
        self.flag_combo.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.flag_preview = tk.Label(flag_row, text="🏳️", font=FLAG_ICON_FONT)
        self.flag_preview.pack(side="left", padx=(0, 0))
        self.flag_var.set("🏳️ Custom")
        self.name_var.trace_add("write", lambda *_: self._set_flag_selection(self.name_var.get()))

        # --- holidays editor ---
        hol_frame = ttk.LabelFrame(self, text="Holidays (Independence Day, etc.)", padding=12)
        hol_frame.pack(fill="both", expand=True, padx=10, pady=8)

        add_row = ttk.Frame(hol_frame)
        add_row.pack(fill="x", pady=4)

        self.holiday_name_label = ttk.Label(add_row, text=self.app.t("holiday_name"))
        self.holiday_name_label.pack(side="left")
        self.holiday_name_var = tk.StringVar()
        ttk.Entry(add_row, textvariable=self.holiday_name_var, width=25).pack(side="left", padx=5)

        self.date_hint_label = ttk.Label(add_row, text=self.app.t("date_hint"))
        self.date_hint_label.pack(side="left", padx=(10, 0))
        self.holiday_date_var = tk.StringVar()
        ttk.Entry(add_row, textvariable=self.holiday_date_var, width=8).pack(side="left", padx=5)

        self.add_holiday_button = ttk.Button(add_row, text=self.app.t("add_holiday"), command=self.add_holiday)
        self.add_holiday_button.pack(side="left", padx=10)
        self.remove_button = ttk.Button(add_row, text=self.app.t("remove_selected"), command=self.remove_holiday)
        self.remove_button.pack(side="left")

        self.holiday_listbox = tk.Listbox(hol_frame, height=8)
        self.holiday_listbox.pack(fill="both", expand=True, pady=(6, 0))

        # --- save ---
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=10, pady=10)
        self.save_button = ttk.Button(bottom, text=self.app.t("save_country"), command=self.save_country)
        self.save_button.pack(side="right")

        self.developer_button = tk.Button(
            self,
            text="🛡 DEV",
            command=self.app._prompt_developer_access,
            width=8,
            height=1,
            bd=1,
            highlightthickness=1,
            bg="#f4f4f4",
            activebackground="#e8e8e8",
            fg="#333333",
            relief="raised",
            font=("Arial", 8, "bold"),
            padx=4,
            pady=2,
        )
        self.developer_button.place(relx=0.5, rely=0.9, anchor="center")

        self.new_country()  # start blank

    def apply_language(self):
        self.select_country_label.config(text=self.app.t("select_country_to_edit"))
        self.new_country_button.config(text=self.app.t("new_country"))
        self.delete_button.config(text=self.app.t("delete_country"))
        self.import_button.config(text=self.app.t("import_excel"))
        self.export_button.config(text=self.app.t("export_excel"))
        self.flag_label.config(text=self.app.t("flag"))
        self.embassy_tabs.tab(0, text=self.app.t("embassy_location_tab"))
        self.embassy_tabs.tab(1, text=self.app.t("embassy_phone_tab"))
        self.holiday_name_label.config(text=self.app.t("holiday_name"))
        self.date_hint_label.config(text=self.app.t("date_hint"))
        self.add_holiday_button.config(text=self.app.t("add_holiday"))
        self.remove_button.config(text=self.app.t("remove_selected"))
        self.save_button.config(text=self.app.t("save_country"))
        self.country_picker.set(self.country_picker.get())

    def _form_row(self, parent, label_text, var):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=3)
        ttk.Label(row, text=label_text, width=22, anchor="w").pack(side="left")
        entry = ttk.Entry(row, textvariable=var, width=40)
        entry.pack(side="left")
        self._enable_paste(entry)
        return entry

    def _enable_paste(self, entry):
        menu = tk.Menu(entry, tearoff=False)
        menu.add_command(label="Paste", command=lambda: self._paste_into(entry))

        def show_menu(event):
            entry.focus_set()
            menu.tk_popup(event.x_root, event.y_root)

        entry.bind("<Button-3>", show_menu)

    def _paste_into(self, entry):
        try:
            value = self.clipboard_get()
        except tk.TclError:
            return
        try:
            entry.delete("sel.first", "sel.last")
        except tk.TclError:
            pass
        entry.insert("insert", value)

    def _set_flag_selection(self, country_name):
        country_name = (country_name or "").strip()
        emoji = FLAG_EMOJIS.get(country_name, "🏳️")
        value = f"{emoji} {country_name or 'Custom'}" if country_name else "🏳️ Custom"
        self.flag_var.set(value)
        flag_image = self.app.get_country_flag_image(country_name)
        if flag_image:
            self.flag_preview.config(image=flag_image, compound="center", text="")
            self.flag_preview.image = flag_image
        else:
            self.flag_preview.config(text=emoji if country_name else "🏳️", image="")

    def refresh_country_list(self):
        names = [c["name"] for c in self.app.store.countries]
        self.country_picker["values"] = names

    def new_country(self):
        self.editing_original_name = None
        self.name_var.set("")
        self.president_var.set("")
        self.minister_var.set("")
        self.ambassador_var.set("")
        self.embassy_location_var.set("")
        self.embassy_phone_var.set("")
        self.flag_var.set("🏳️ Custom")
        self.flag_preview.config(text="🏳️", image="")
        self.current_holidays = []
        self.refresh_holiday_listbox()
        self.country_picker.set("")

    def load_selected_country(self, event=None):
        name = self.country_picker.get()
        c = self.app.store.get(name)
        if not c:
            return
        self.editing_original_name = name
        self.name_var.set(c["name"])
        self.president_var.set(c.get("president", ""))
        self.minister_var.set(c.get("minister_of_defence", ""))
        self.ambassador_var.set(c.get("ambassador_to_lebanon", ""))
        self.embassy_location_var.set(c.get("embassy_location", ""))
        self.embassy_phone_var.set(c.get("embassy_phone", ""))
        flag_emoji = c.get("flag") or self.app.get_country_flag(c["name"])
        self.flag_var.set(f"{flag_emoji} {c['name']}")
        flag_image = self.app.get_country_flag_image(c["name"])
        if flag_image:
            self.flag_preview.config(image=flag_image, compound="center", text="")
            self.flag_preview.image = flag_image
        else:
            self.flag_preview.config(text=flag_emoji, image="")
        self.current_holidays = list(c.get("holidays", []))
        self.refresh_holiday_listbox()

    def refresh_holiday_listbox(self):
        self.holiday_listbox.delete(0, "end")
        for h in self.current_holidays:
            self.holiday_listbox.insert("end", f"{h['name']}  —  {h['date']}")

    def add_holiday(self):
        name = self.holiday_name_var.get().strip()
        date_str = self.holiday_date_var.get().strip()
        if not name or not date_str:
            messagebox.showwarning("Missing info", "Please enter both a holiday name and a date (MM-DD).")
            return
        try:
            month, day = map(int, date_str.split("-"))
            date(2001, month, day)  # 2001 is just a validation year (non-leap)
        except (ValueError, AttributeError):
            messagebox.showerror("Invalid date", "Please use the format MM-DD, e.g. 11-22 for Nov 22.")
            return

        self.current_holidays.append({"name": name, "date": f"{month:02d}-{day:02d}"})
        self.current_holidays = filter_allowed_holidays(self.current_holidays)
        self.holiday_name_var.set("")
        self.holiday_date_var.set("")
        self.refresh_holiday_listbox()

    def remove_holiday(self):
        sel = self.holiday_listbox.curselection()
        if not sel:
            return
        del self.current_holidays[sel[0]]
        self.refresh_holiday_listbox()

    def delete_country(self):
        name = self.country_picker.get()
        if not name:
            messagebox.showinfo("No selection", "Select a country from the dropdown first.")
            return
        if messagebox.askyesno("Confirm delete", f"Delete '{name}' and all its data?"):
            self.app.store.delete(name)
            self.app.refresh_all()
            self.new_country()

    def save_country(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("Missing name", "Please enter a country name.")
            return

        selected_flag = self.flag_var.get().strip()
        if selected_flag and selected_flag != "🏳️ Custom":
            flag_emoji = selected_flag.split(" ", 1)[0]
        else:
            flag_emoji = self.app.get_country_flag(name)
        flag_image = self.app.get_country_flag_image(name)
        if flag_image:
            self.flag_preview.config(image=flag_image, compound="center", text="")
            self.flag_preview.image = flag_image
        else:
            self.flag_preview.config(text=flag_emoji, image="")

        # If renaming an existing country, remove the old entry first
        if self.editing_original_name and self.editing_original_name != name:
            self.app.store.delete(self.editing_original_name)

        self.current_holidays = filter_allowed_holidays(self.current_holidays)
        country = {
            "name": name,
            "flag": flag_emoji,
            "president": self.president_var.get().strip(),
            "minister_of_defence": self.minister_var.get().strip(),
            "ambassador_to_lebanon": self.ambassador_var.get().strip(),
            "embassy_location": self.embassy_location_var.get().strip(),
            "embassy_phone": self.embassy_phone_var.get().strip(),
            "holidays": self.current_holidays,
        }
        self.app.store.upsert(country)
        self.app.refresh_all()
        messagebox.showinfo("Saved", f"'{name}' has been saved.")
        self.editing_original_name = name
        self.country_picker.set(name)


if __name__ == "__main__":
    app = App()
    app._start_background_mode()
    app.mainloop()