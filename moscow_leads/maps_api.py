"""
Сбор компаний Москвы с телефонами через официальные API Google Maps и 2ГИС.

Нужны ключи (задаются переменными окружения):
    GOOGLE_MAPS_API_KEY  — Google Cloud, включить «Places API (New)»
    DGIS_API_KEY         — https://dev.2gis.com (Places API / Catalog API)
Можно задать один или оба: результаты объединяются, дубли по номеру убираются.

Запуск:
    pip install requests openpyxl
    export GOOGLE_MAPS_API_KEY=...
    python maps_api.py --limit 3000 --out moscow_leads.xlsx
"""
import argparse
import os
import sys
import time

import requests

from parse_moscow import normalize_phone, select, write_excel

# Поисковый запрос -> (категория, подкатегория)
QUERIES = {
    "ресторан": ("Рестораны и кафе", "Ресторан"),
    "кафе": ("Рестораны и кафе", "Кафе"),
    "кофейня": ("Рестораны и кафе", "Кофейня"),
    "бар": ("Рестораны и кафе", "Бар"),
    "пекарня": ("Рестораны и кафе", "Пекарня"),
    "салон красоты": ("Красота и уход", "Салон красоты"),
    "барбершоп": ("Красота и уход", "Барбершоп"),
    "ногтевая студия": ("Красота и уход", "Маникюр"),
    "косметология": ("Красота и уход", "Косметология"),
    "стоматология": ("Медицина", "Стоматология"),
    "медицинский центр": ("Медицина", "Клиника"),
    "ветеринарная клиника": ("Медицина", "Ветклиника"),
    "фитнес клуб": ("Спорт и фитнес", "Фитнес-клуб"),
    "йога студия": ("Спорт и фитнес", "Йога"),
    "танцевальная студия": ("Спорт и фитнес", "Танцы"),
    "автосервис": ("Автобизнес", "Автосервис"),
    "автомойка": ("Автобизнес", "Автомойка"),
    "шиномонтаж": ("Автобизнес", "Шиномонтаж"),
    "гостиница": ("Гостиницы", "Отель"),
    "языковая школа": ("Образование", "Языковая школа"),
    "детский развивающий центр": ("Образование", "Детский центр"),
    "автошкола": ("Образование", "Автошкола"),
    "химчистка": ("Бытовые услуги", "Химчистка"),
    "ремонт телефонов": ("Бытовые услуги", "Ремонт техники"),
    "ремонт квартир": ("Бытовые услуги", "Ремонт квартир"),
    "клининговая компания": ("Бытовые услуги", "Клининг"),
    "агентство недвижимости": ("Компании и офисы", "Агентство недвижимости"),
    "юридические услуги": ("Компании и офисы", "Юридические услуги"),
    "бухгалтерские услуги": ("Компании и офисы", "Бухгалтерия"),
    "туристическое агентство": ("Компании и офисы", "Турагентство"),
    "грузоперевозки": ("Компании и офисы", "Логистика"),
    "производство мебели": ("Компании и офисы", "Производство"),
    "цветочный магазин": ("Магазины", "Цветы"),
    "магазин одежды": ("Магазины", "Одежда"),
    "зоомагазин": ("Магазины", "Зоотовары"),
    "квест": ("Развлечения", "Квест"),
    "кальянная": ("Развлечения", "Кальянная"),
}

# Районы — чтобы обойти лимит выдачи (Google отдаёт максимум 60 мест на запрос).
DISTRICTS = [
    "Арбат", "Тверской", "Пресненский", "Хамовники", "Замоскворечье", "Басманный",
    "Таганский", "Мещанский", "Красносельский", "Якиманка", "Сокол", "Аэропорт",
    "Хорошёвский", "Савёловский", "Беговой", "Останкинский", "Алексеевский",
    "Марьина Роща", "Сокольники", "Преображенское", "Измайлово", "Перово",
    "Новогиреево", "Лефортово", "Нижегородский", "Люблино", "Марьино", "Печатники",
    "Текстильщики", "Кузьминки", "Даниловский", "Нагатино", "Чертаново", "Бирюлёво",
    "Царицыно", "Орехово-Борисово", "Зюзино", "Черёмушки", "Академический",
    "Гагаринский", "Ломоносовский", "Тёплый Стан", "Ясенево", "Коньково",
    "Раменки", "Дорогомилово", "Кунцево", "Фили-Давыдково", "Солнцево",
    "Тропарёво-Никулино", "Хорошёво-Мнёвники", "Строгино", "Митино", "Тушино",
    "Щукино", "Ховрино", "Беломорская", "Отрадное", "Бибирево", "Медведково",
    "Свиблово", "Бабушкинский", "Ростокино", "Гольяново", "Новокосино", "Выхино",
    "Жулебино", "Зеленоград", "Коммунарка",
]

# Прямоугольник Москвы в пределах МКАД + Зеленоград/Новая Москва частично.
MOSCOW_RECT = {"low": {"latitude": 55.49, "longitude": 37.17},
               "high": {"latitude": 56.03, "longitude": 37.97}}


def empty_row(name, cat, phones, source):
    return {
        "Название": name, "Категория": cat[0], "Подкатегория": cat[1],
        "Телефон": phones[0], "Доп. телефоны": ", ".join(phones[1:]),
        "Сайт": "", "Есть сайт": "Нет", "Email": "", "Адрес": "", "Часы работы": "",
        "VK / Telegram": "", "Ссылка на карте": "", "Источник": source,
        "Статус звонка": "Не звонили", "Комментарий": "",
    }


def google_search(key, text):
    """До 60 мест на запрос (3 страницы по 20)."""
    url = "https://places.googleapis.com/v1/places:searchText"
    headers = {
        "X-Goog-Api-Key": key,
        "X-Goog-FieldMask": ",".join([
            "places.id", "places.displayName", "places.formattedAddress",
            "places.nationalPhoneNumber", "places.internationalPhoneNumber",
            "places.websiteUri", "places.googleMapsUri", "places.businessStatus",
            "nextPageToken",
        ]),
    }
    body = {"textQuery": text, "languageCode": "ru", "regionCode": "RU", "pageSize": 20,
            "locationRestriction": {"rectangle": MOSCOW_RECT}}
    for _ in range(3):
        resp = requests.post(url, json=body, headers=headers, timeout=30)
        if resp.status_code != 200:
            sys.exit(f"Google API ответил {resp.status_code}: {resp.text[:300]}")
        data = resp.json()
        yield from data.get("places", [])
        token = data.get("nextPageToken")
        if not token:
            break
        body["pageToken"] = token
        time.sleep(1)


def google_row(place, cat):
    if place.get("businessStatus", "OPERATIONAL") != "OPERATIONAL":
        return None
    phones = normalize_phone(place.get("internationalPhoneNumber") or place.get("nationalPhoneNumber") or "")
    if not phones:
        return None
    row = empty_row(place.get("displayName", {}).get("text", ""), cat, phones, "Google Maps")
    site = place.get("websiteUri", "")
    row.update({"Сайт": site, "Есть сайт": "Да" if site else "Нет",
                "Адрес": place.get("formattedAddress", "").replace(", Россия", ""),
                "Ссылка на карте": place.get("googleMapsUri", "")})
    return row


def dgis_search(key, text, pages=5):
    url = "https://catalog.api.2gis.com/3.0/items"
    for page in range(1, pages + 1):
        params = {"q": text, "region_id": 32, "type": "branch", "page": page, "page_size": 10,
                  "fields": "items.contact_groups,items.address,items.full_address_name,items.schedule",
                  "key": key}
        resp = requests.get(url, params=params, timeout=30)
        data = resp.json()
        if data.get("meta", {}).get("code") != 200:
            if page == 1:
                print(f"  2ГИС: {data.get('meta', {}).get('error', {}).get('message')}", file=sys.stderr)
            return
        items = data.get("result", {}).get("items", [])
        yield from items
        if len(items) < 10:
            return
        time.sleep(0.3)


def dgis_row(item, cat):
    contacts = [c for g in item.get("contact_groups", []) for c in g.get("contacts", [])]
    raw_phones = ";".join(c.get("value", "") for c in contacts if c.get("type") == "phone")
    phones = normalize_phone(raw_phones)
    if not phones:
        return None
    row = empty_row(item.get("name", ""), cat, phones, "2ГИС")
    by_type = {c.get("type"): c.get("value", "") for c in contacts}
    site = by_type.get("website", "")
    row.update({
        "Сайт": site, "Есть сайт": "Да" if site else "Нет", "Email": by_type.get("email", ""),
        "Адрес": item.get("address_name") or item.get("full_address_name", ""),
        "VK / Telegram": by_type.get("vkontakte") or by_type.get("telegram", ""),
        "Ссылка на карте": f"https://2gis.ru/moscow/firm/{item.get('id', '').split('_')[0]}",
    })
    return row


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--limit", type=int, default=3000)
    ap.add_argument("--out", default="moscow_leads.xlsx")
    ap.add_argument("--districts", type=int, default=len(DISTRICTS),
                    help="сколько районов обходить (меньше = быстрее и дешевле)")
    ap.add_argument("--any-site", dest="no_site_first", action="store_false")
    args = ap.parse_args()

    gkey, dkey = os.environ.get("GOOGLE_MAPS_API_KEY"), os.environ.get("DGIS_API_KEY")
    if not gkey and not dkey:
        sys.exit("Задайте GOOGLE_MAPS_API_KEY и/или DGIS_API_KEY.")

    rows, seen = [], set()
    per_cat_target = args.limit // len({c[0] for c in QUERIES.values()}) * 2

    def add(row):
        if row and row["Название"] and row["Телефон"] not in seen:
            seen.update([row["Телефон"]] + [p for p in row["Доп. телефоны"].split(", ") if p])
            rows.append(row)

    for query, cat in QUERIES.items():
        for district in DISTRICTS[:args.districts]:
            if sum(r["Категория"] == cat[0] for r in rows) >= per_cat_target:
                break
            text = f"{query} {district} Москва"
            before = len(rows)
            if gkey:
                for p in google_search(gkey, text):
                    add(google_row(p, cat))
            if dkey:
                for it in dgis_search(dkey, text):
                    add(dgis_row(it, cat))
            print(f"{text}: +{len(rows) - before} (всего {len(rows)})", file=sys.stderr)

    rows = select(rows, args.limit, args.no_site_first)
    write_excel(rows, args.out)
    print(f"Сохранено {len(rows)} записей в {args.out}", file=sys.stderr)


if __name__ == "__main__":
    main()
