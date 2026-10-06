#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Парсер реестра ФСА (pub.fsa.gov.ru): сертификаты и декларации о соответствии.

Выбирает документы, срок действия которых истекает в ближайшие N дней,
фильтрует по наименованию продукции и выгружает в Excel / CSV.

Запуск:
    python fsa_parser.py                 # по настройкам из config.ini
    python fsa_parser.py --days 90       # разово изменить порог
    python fsa_parser.py --debug         # сохранить «сырые» ответы сайта для диагностики
"""
import argparse
import configparser
import csv
import json
import os
import random
import re
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import requests

# ----------------------------------------------------------------------------
# Консоль: UTF-8 и цвета (в т.ч. в Windows 10/11)
# ----------------------------------------------------------------------------
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:
        pass
if os.name == "nt":
    os.system("")  # включает ANSI-последовательности в cmd

C = {"b": "\033[94m", "g": "\033[92m", "y": "\033[93m", "r": "\033[91m",
     "d": "\033[90m", "x": "\033[0m"}


def paint(text, color):
    return f"{C[color]}{text}{C['x']}"


# ----------------------------------------------------------------------------
# Константы
# ----------------------------------------------------------------------------
BASE = "https://pub.fsa.gov.ru"
ANON_LOGIN = {"username": "anonymous", "password": "hrgesf7HDR67Bd"}  # публичный гостевой вход сайта

DOC_KINDS = {
    "certificates": {
        "title": "Сертификат",
        "list_url": f"{BASE}/api/v1/rss/common/certificates/get",
        "detail_url": f"{BASE}/api/v1/rss/common/certificates/{{id}}",
        "sort_column": "endDate",
    },
    "declarations": {
        "title": "Декларация",
        "list_url": f"{BASE}/api/v1/rds/common/declarations/get",
        "detail_url": f"{BASE}/api/v1/rds/common/declarations/{{id}}",
        "sort_column": "declEndDate",  # у сущности SearchDeclaration нет поля endDate
    },
}
STATUS_ACTIVE = 6  # статус «Действует»

# Справочник ТР из фильтра сайта: название -> внутренний код (по нему сайт сам отбирает документы)
TECH_REGS = {
    "Решение Совета ЕЭК от 18.04.2018 № 44": 4511534,
    "ТР ТС 001/2011": 33,
    "ТР ТС 002/2011": 18,
    "ТР ТС 003/2011": 37,
    "ТР ТС 004/2011": 11,
    "ТР ТС 005/2011": 38,
    "ТР ТС 006/2011": 1,
    "ТР ТС 007/2011": 39,
    "ТР ТС 008/2011": 17,
    "ТР ТС 009/2011": 28,
    "ТР ТС 010/2011": 14,
    "ТР ТС 011/2011": 6,
    "ТР ТС 012/2011": 12,
    "ТР ТС 013/2011": 31,
    "ТР ТС 014/2011": 29,
    "ТР ТС 015/2011": 32,
    "ТР ТС 016/2011": 16,
    "ТР ТС 017/2011": 8,
    "ТР ТС 018/2011": 30,
    "ТР ТС 019/2011": 36,
    "ТР ТС 020/2011": 24,
    "ТР ТС 021/2011": 27,
    "ТР ТС 022/2011": 97161,
    "ТР ТС 023/2011": 22,
    "ТР ТС 024/2011": 20,
    "ТР ТС 025/2012": 4521250,
    "ТР ТС 026/2012": 40,
    "ТР ТС 027/2012": 97185,
    "ТР ТС 028/2012": 26,
    "ТР ТС 029/2012": 23,
    "ТР ТС 030/2012": 34,
    "ТР ТС 031/2012": 10,
    "ТР ТС 032/2013": 5,
    "ТР ТС 033/2013": 21,
    "ТР ТС 034/2013": 2,
    "ТР ТС 035/2014": 35,
    "ТР ЕАЭС 036/2016": 97149,
    "ТР ЕАЭС 037/2016": 4501729,
    "ТР ЕАЭС 038/2016": 97167,
    "ТР ЕАЭС 039/2016": 97156,
    "ТР ЕАЭС 040/2016": 15,
    "ТР ЕАЭС 041/2017": 97209,
    "ТР ЕАЭС 042/2017": 97210,
    "ТР ЕАЭС 043/2017": 97211,
    "ТР ЕАЭС 044/2017": 97208,
    "ТР ЕАЭС 045/2017": 99333,
    "ТР ЕАЭС 046/2018": 4507056,
    "ТР ЕАЭС 047/2018": 101330,
    "123-ФЗ": 56,
    "ТР РФ 007/2008": 49,
    "ТР РФ 015/2009": 97187,
    "ТР РФ 016/2010": 97170,
    "ТР РФ 022/2010": 97163,
    "ТР РФ 023/2010": 60,
    "ТР РФ 024/2010": 97169,
    "ГОСТ Р 54008-2022": 4518156,
    "ТР ЕАЭС 052/2021": 4522924,
    "ТР ЕАЭС 050/2021": 4517710,
    "ТР ЕАЭС 051/2021": 4516872,
    "ТР ЕАЭС 048/2019": 4508584,
    "126-ФЗ": 98706,
}

# Полные названия самых частых регламентов (для столбца в выгрузке)
TECH_REG_TITLES = {
    "ТР ТС 005/2011": "О безопасности упаковки",
    "ТР ТС 021/2011": "О безопасности пищевой продукции",
    "ТР ТС 022/2011": "Пищевая продукция в части ее маркировки",
    "ТР ТС 023/2011": "Технический регламент на соковую продукцию из фруктов и овощей",
    "ТР ТС 024/2011": "Технический регламент на масложировую продукцию",
    "ТР ТС 027/2012": "О безопасности отдельных видов специализированной пищевой продукции",
    "ТР ТС 029/2012": "Требования безопасности пищевых добавок, ароматизаторов и технологических вспомогательных средств",
    "ТР ТС 033/2013": "О безопасности молока и молочной продукции",
    "ТР ТС 034/2013": "О безопасности мяса и мясной продукции",
    "ТР ЕАЭС 040/2016": "О безопасности рыбы и рыбной продукции",
}

COLUMN_TITLES = {
    "company": "Наименование юрлица",
    "head": "ФИО руководителя",
    "last_name": "Фамилия",
    "first_name": "Имя",
    "patronymic": "Отчество",
    "number": "Номер документа",
    "doc_type": "Тип документа",
    "end_date": "Дата окончания",
    "days_left": "Дней до конца",
    "email": "Электронный адрес",
    "product": "Наименование продукции",
    "tech_reg": "Наименование ТР ТС",
}
COLUMN_WIDTHS = {"company": 34, "head": 26, "last_name": 18, "first_name": 16, "patronymic": 20, "number": 34, "doc_type": 14,
                 "end_date": 16, "days_left": 14, "email": 30, "product": 50, "tech_reg": 40}

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


# ----------------------------------------------------------------------------
# Настройки
# ----------------------------------------------------------------------------
def tech_reg_label(name):
    t = TECH_REG_TITLES.get(name)
    return f"{name} ({t})" if t else name


def resolve_tech_regs(raw_items):
    """'ТР ТС 021/2011', '021/2011', 'тр еаэс 040/2016' -> [(название, код)]."""
    out = []
    for raw in raw_items:
        m = re.search(r"(\d{3})\s*/\s*(\d{4})", raw)
        if not m:
            sys.exit(f"В tech_regs не понятен номер: «{raw}». Пример: ТР ТС 021/2011")
        num = f"{m[1]}/{m[2]}"
        found = [k for k in TECH_REGS if k.endswith(" " + num)]
        if not found:
            sys.exit(f"ТР «{raw}» нет в справочнике сайта. Список доступных: python fsa_parser.py --list-tr")
        pick = next((k for k in found if k.split()[1].lower() in raw.lower().replace(" ", "")
                     or k.split()[1].lower() in raw.lower()), found[0]) if len(found) > 1 else found[0]
        if (pick, TECH_REGS[pick]) not in out:
            out.append((pick, TECH_REGS[pick]))
    return out


def load_config(path):
    cp = configparser.ConfigParser(inline_comment_prefixes=None)
    if not Path(path).exists():
        sys.exit(f"Не найден файл настроек: {path}")
    cp.read(path, encoding="utf-8-sig")

    def lst(section, key, default=""):
        raw = cp.get(section, key, fallback=default)
        return [x.strip() for x in raw.split(",") if x.strip()]

    cfg = {
        "days": cp.getint("filter", "days_threshold", fallback=60),
        "keywords": [k.lower() for k in lst("filter", "product_keywords")],
        "tech_regs": resolve_tech_regs(lst("filter", "tech_regs")),
        "kinds": [k for k in lst("filter", "document_types", "certificates,declarations")
                  if k in DOC_KINDS],
        "format": cp.get("output", "format", fallback="xlsx").strip().lower(),
        "out_dir": cp.get("output", "output_dir", fallback="выгрузки").strip(),
        "prefix": cp.get("output", "file_prefix", fallback="fsa").strip(),
        "columns": [c for c in lst("output", "columns") if c in COLUMN_TITLES]
                   or list(COLUMN_TITLES),
        "name_case": cp.get("output", "name_case", fallback="title").strip().lower(),
        "urgent": cp.getint("output", "urgent_days", fallback=30),
        "preview": cp.getint("output", "preview_rows", fallback=7),
        "page_size": min(cp.getint("network", "page_size", fallback=100), 100),
        "delay": cp.getfloat("network", "delay", fallback=0.3),
        "retries": cp.getint("network", "retries", fallback=4),
        "batch_size": cp.getint("network", "batch_size", fallback=100),
        "batch_pause": cp.getfloat("network", "batch_pause", fallback=30),
        "max_requests": cp.getint("network", "max_requests", fallback=0),
        "cache_dir": cp.get("cache", "cache_dir", fallback="cache").strip(),
        "ttl_days": cp.getint("cache", "details_ttl_days", fallback=30),
        "dedupe": cp.getboolean("cache", "dedupe_applicants", fallback=True),
        "proxy": cp.get("network", "proxy", fallback="").strip(),
        "chunk_days": max(cp.getint("network", "chunk_days", fallback=3), 1),
    }
    if not cfg["kinds"]:
        sys.exit("В config.ini не указаны document_types (certificates / declarations).")
    return cfg


# ----------------------------------------------------------------------------
# Клиент API ФСА
# ----------------------------------------------------------------------------
class ApiBadRequest(RuntimeError):
    """Сайт ответил 400 - повторять тот же запрос бессмысленно."""


class NetworkGiveUp(RuntimeError):
    """Сеть/сайт не отвечает (вероятно, блокировка IP). Дальше стучаться нельзя."""


class FilterIgnored(NetworkGiveUp):
    """Сайт не применил фильтр по ТР ТС - выгрузка получилась бы неверной."""


class RunLimit(NetworkGiveUp):
    """Исчерпан лимит запросов за один запуск (max_requests в config.ini)."""


class FsaClient:
    def __init__(self, cfg):
        self.delay = cfg["delay"]
        self.retries = max(cfg["retries"], 1)
        self.batch_size = cfg["batch_size"]
        self.batch_pause = cfg["batch_pause"]
        self.max_requests = cfg["max_requests"]
        self.made = 0
        self.s = requests.Session()
        proxy = (cfg["proxy"] or "").strip()
        if proxy.lower() == "none":              # принудительно без прокси
            self.s.trust_env = False
        elif proxy:                              # свой прокси из config.ini
            self.s.proxies = {"http": proxy, "https": proxy}
        self.s.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
            "Accept": "application/json, text/plain, */*",
            "Content-Type": "application/json",
            "Origin": BASE,
            "Referer": f"{BASE}/rss/certificate",
        })
        self.token = None

    def login(self):
        r = self.s.post(f"{BASE}/login", json=ANON_LOGIN, timeout=30)
        r.raise_for_status()
        token = r.headers.get("Authorization") or r.headers.get("authorization")
        if not token:
            raise NetworkGiveUp("Сайт не вернул токен авторизации (блокировка или изменился вход)")
        self.token = token
        self.s.headers["Authorization"] = token

    def _pace(self):
        """Неравномерные паузы + длинный перерыв после каждой пачки запросов."""
        self.made += 1
        time.sleep(self.delay * random.uniform(0.6, 1.6))
        if self.batch_size and self.made % self.batch_size == 0 and self.batch_pause > 0:
            sys.stdout.write(f"\r   {paint(f'Пауза {self.batch_pause:.0f} с, чтобы не нагружать сайт...', 'd')}"
                             + " " * 20)
            sys.stdout.flush()
            time.sleep(self.batch_pause)

    def _request(self, method, url, **kw):
        if self.max_requests and self.made >= self.max_requests:
            raise RunLimit(f"достигнут лимит {self.max_requests} запросов за запуск (max_requests)")
        last = None
        for attempt in range(1, self.retries + 1):
            wait = min(5 * 3 ** (attempt - 1), 120)       # 5, 15, 45, 120 с
            try:
                if not self.token:
                    self.login()
                r = self.s.request(method, url, timeout=60, **kw)
                sc = r.status_code
                if sc == 400:
                    raise ApiBadRequest(r.text[:300] or "HTTP 400")
                if sc in (401, 403):
                    self.token = None                      # токен истёк - войдём заново
                    last = RuntimeError(f"HTTP {sc}")
                elif sc == 429 or sc >= 500:
                    last = RuntimeError(f"HTTP {sc}")
                    ra = r.headers.get("Retry-After", "")
                    if ra.isdigit():
                        wait = min(int(ra), 300)
                else:
                    r.raise_for_status()
                    data = r.json()
                    self._pace()
                    return data
            except requests.exceptions.ProxyError as e:
                raise NetworkGiveUp(
                    "Не удаётся подключиться к прокси (он не запущен или указан неверно).\n"
                    "  Запустите прокси/VPN либо задайте параметр proxy в config.ini:\n"
                    "    proxy =                        - брать из системы (HTTPS_PROXY)\n"
                    "    proxy = none                   - работать без прокси\n"
                    "    proxy = http://127.0.0.1:10809 - указать прокси явно\n"
                    f"  Детали: {str(e)[:200]}")
            except (requests.RequestException, ValueError) as e:
                last = e
            if attempt < self.retries:
                time.sleep(wait)
        raise NetworkGiveUp(f"сайт не отвечает ({str(last)[:150]}). Возможно, IP временно заблокирован")

    def list_page(self, kind, page, size, d_from, d_to, tr_id=None):
        body = {
            "size": size,
            "page": page,
            "filter": {
                "status": [STATUS_ACTIVE],
                "endDate": {"minDate": d_from.isoformat(), "maxDate": d_to.isoformat()},
                "regDate": {"minDate": None, "maxDate": None},
                "idCertType": [], "idObjectType": [], "idProductType": [],
                "idGroupRU": [], "idGroupEEU": [], "idTechReg": [tr_id] if tr_id else [], "idApplicantType": [],
                "idCertScheme": [], "idDeclScheme": [], "idProductOrigin": [],
                "idProductEEU": [], "idProductRU": [],
                "columnsSearch": [],
                "awaitForApprove": None, "awaitOperatorCheck": None, "editApp": None,
                "violationSend": None, "isProtected": None, "isNotAccredited": None,
            },
            "columnsSort": [{"column": DOC_KINDS[kind]["sort_column"], "sort": "ASC"}],
        }
        return self._request("POST", DOC_KINDS[kind]["list_url"], data=json.dumps(body))

    def detail(self, kind, doc_id):
        return self._request("GET", DOC_KINDS[kind]["detail_url"].format(id=doc_id))


# ----------------------------------------------------------------------------
# Разбор ответов (сделан устойчивым к небольшим отличиям в названиях полей)
# ----------------------------------------------------------------------------
def first_key(d, *names, suffix=None):
    """Первое непустое значение по списку имён ключей (или по окончанию ключа)."""
    if not isinstance(d, dict):
        return None
    for n in names:
        v = d.get(n)
        if v not in (None, "", []):
            return v
    if suffix:
        for k, v in d.items():
            if k.lower().endswith(suffix.lower()) and v not in (None, "", []):
                return v
    return None


def parse_date(v):
    if not v:
        return None
    if isinstance(v, (int, float)):                       # unix-время в мс
        return datetime.fromtimestamp(v / 1000).date()
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", str(v))
    if m:
        return date(int(m[1]), int(m[2]), int(m[3]))
    m = re.match(r"(\d{2})\.(\d{2})\.(\d{4})", str(v))
    if m:
        return date(int(m[3]), int(m[2]), int(m[1]))
    return None


def collect_strings(obj):
    if isinstance(obj, str):
        yield obj
    elif isinstance(obj, dict):
        for v in obj.values():
            yield from collect_strings(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from collect_strings(v)


def find_email(applicant, whole):
    """Сначала ищем среди контактов заявителя, затем по всему документу."""
    for src in (applicant, whole):
        if not src:
            continue
        # приоритет - структуры контактов
        contacts = src.get("contacts") if isinstance(src, dict) else None
        for text in collect_strings(contacts or []):
            m = EMAIL_RE.search(text)
            if m:
                return m.group(0).lower()
        for text in collect_strings(src):
            m = EMAIL_RE.search(text)
            if m:
                return m.group(0).lower()
    return ""


def find_head(applicant):
    """ФИО руководителя заявителя."""
    if not isinstance(applicant, dict):
        return ""
    direct = first_key(applicant, "headFio", "headName", "directorFio", "directorName",
                       "managerFio", "manager", "headFullName", "fioHead")
    if isinstance(direct, str):
        return direct.strip()
    if isinstance(direct, dict):
        applicant = direct
    parts = [applicant.get("surname"), applicant.get("firstName"), applicant.get("patronymic")]
    if any(parts):
        return " ".join(p.strip() for p in parts if p).strip()
    for k, v in applicant.items():                         # запасной вариант
        if isinstance(v, str) and any(w in k.lower() for w in ("head", "director", "fio")):
            return v.strip()
    return ""


def extract_applicant(det):
    """Из карточки документа берём только то, что нужно: юрлицо, руководитель, e-mail."""
    applicant = (det or {}).get("applicant") or {}
    return {
        "company": str(first_key(applicant, "fullName", "shortName", "name") or "").strip(),
        "head": find_head(applicant),
        "email": find_email(applicant, det),
    }


def applicant_key(item):
    """Ключ заявителя для общего кэша (ИНН, если есть в списке, иначе нормализованное имя)."""
    v = first_key(item, "applicantInn", "applicantINN", "inn", "idApplicant")
    if v:
        return f"id:{v}"
    n = " ".join(str(first_key(item, "applicantName") or "").lower().split())
    return f"n:{n}" if n else None


NAME_PARTICLES = {"оглы", "огли", "кызы", "кизи", "гызы", "угли", "улы", "кзы"}


def clean_name(name, mode="title"):
    """title - ЗАГЛАВНЫЕ буквы превращаем в обычный регистр; registry - как в реестре."""
    name = " ".join(str(name or "").split())
    if mode == "registry":
        return name
    out = []
    for w in name.split(" "):
        if w.lower() in NAME_PARTICLES:
            out.append(w.lower())
        elif w.isupper():
            out.append(w.title())
        else:
            out.append(w)
    return " ".join(out)


def split_name(name):
    """В реестре порядок: Фамилия Имя Отчество. 'Мамед оглы' остаётся в отчестве целиком."""
    parts = name.split()
    if not parts:
        return "", "", ""
    return parts[0], (parts[1] if len(parts) > 1 else ""), " ".join(parts[2:])


def build_row(kind, item, info, today, name_case="title", tech_regs=()):
    info = info or {}
    head = clean_name(info.get("head", ""), name_case)
    last_name, first_name, patronymic = split_name(head)
    end = parse_date(first_key(item, "certEndDate", "declEndDate", "endDate", suffix="EndDate"))
    company = info.get("company") or first_key(item, "applicantName") or ""
    product = first_key(item, "productFullName", "productName") or ""
    return {
        "company": str(company).strip(),
        "head": head,
        "last_name": last_name,
        "first_name": first_name,
        "patronymic": patronymic,
        "number": str(first_key(item, "number") or "").strip(),
        "doc_type": DOC_KINDS[kind]["title"],
        "end_date": end,
        "days_left": (end - today).days if end else None,
        "email": info.get("email", ""),
        "product": str(product).strip().replace("\n", " "),
        "tech_reg": "; ".join(tech_reg_label(n) for n in sorted(tech_regs)),
        "_has_info": bool(info),
    }


# ----------------------------------------------------------------------------
# Кэш на диске: найденное не пропадает при обрыве, повторный запуск продолжает с места остановки
# ----------------------------------------------------------------------------
class Cache:
    def __init__(self, folder, ttl_days):
        Path(folder).mkdir(parents=True, exist_ok=True)
        self.path = Path(folder) / "fsa_cache.json"
        self.data = {"windows": {}, "details": {}, "apps": {}}
        if self.path.exists():
            try:
                self.data.update(json.loads(self.path.read_text(encoding="utf-8")))
            except Exception:
                pass                                   # битый кэш просто игнорируем
        now, ttl, today = time.time(), ttl_days * 86400, date.today().isoformat()
        self.data["windows"] = {k: v for k, v in self.data["windows"].items() if k.endswith("|" + today)}
        for sect in ("details", "apps"):
            self.data[sect] = {k: v for k, v in self.data[sect].items() if now - v.get("t", 0) < ttl}
        self._dirty = 0
        self._last_save = time.time()

    def save(self):
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self.data, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, self.path)                     # запись атомарная - файл не испортится при обрыве
        self._dirty = 0
        self._last_save = time.time()

    def put_window(self, key, items):
        self.data["windows"][key] = items
        if time.time() - self._last_save > 20:       # на диск - не чаще раза в 20 сек (и при выходе)
            self.save()

    def get_info(self, doc_key, app_key):
        v = self.data["details"].get(doc_key) or (self.data["apps"].get(app_key) if app_key else None)
        return {k: v[k] for k in ("company", "head", "email")} if v else None

    def put_info(self, doc_key, app_key, info):
        rec = dict(info, t=time.time())
        self.data["details"][doc_key] = rec
        if app_key:
            self.data["apps"][app_key] = rec
        self._dirty += 1
        if self._dirty >= 20:
            self.save()


# ----------------------------------------------------------------------------
# Сбор данных
# ----------------------------------------------------------------------------
def progress(done, total, label, unit=None):
    if unit:   # done - найдено документов, total - пройдено дней
        sys.stdout.write(f"\r   {paint(label, 'd')} найдено {done}, пройдено {total} {unit}      ")
        sys.stdout.flush()
        return
    total = max(total, 1)
    pct = int(done * 100 / total)
    sys.stdout.write(f"\r   {paint(label, 'd')} {done:>6}/{total:<6} ({pct:>3}%)   ")
    sys.stdout.flush()


MAX_WINDOW = 1500   # если в окне больше записей - делим окно пополам


def slim(it):
    """В кэш пишем только нужные поля - файл кэша остаётся небольшим."""
    keep = ("id", "number", "productFullName", "productName", "applicantName", "applicantInn",
            "applicantINN", "inn", "idApplicant")
    return {k: v for k, v in it.items() if k in keep or k.lower().endswith("enddate")}


def absorb(cfg, items, acc, tr_name=None):
    for it in items:
        if tr_name:                                  # один документ может подпадать под несколько ТР
            acc["trs"].setdefault(it.get("id"), set()).add(tr_name)
        if it.get("id") in acc["seen"]:
            continue
        acc["seen"].add(it.get("id"))
        name = str(first_key(it, "productFullName", "productName") or "").lower()
        if not cfg["keywords"] or any(k in name for k in cfg["keywords"]):
            acc["matched"].append(it)


def fetch_window(client, kind, cfg, cache, a, b, acc, state, tr=None):
    """Читает все страницы за окно дат [a, b]. При лимитах сайта делит окно пополам."""
    tr_name, tr_id = tr if tr else (None, None)
    key = f"{kind}|{tr_id}|{a}|{b}|{date.today()}"
    cached = cache.data["windows"].get(key)
    if cached is not None:                       # окно уже загружено ранее - сайт не трогаем
        absorb(cfg, cached, acc, tr_name)
        return

    def split():
        span = (b - a).days
        if span >= 2:
            mid = a + timedelta(days=span // 2)
            parts = [(a, mid), (mid, b)]         # перекрытие на 1 день - дубли снимает seen
        else:
            parts = [(a, a), (b, b)]
        for x, y in parts:
            fetch_window(client, kind, cfg, cache, x, y, acc, state, tr)

    page, got = 0, []
    while True:
        try:
            resp = client.list_page(kind, page, cfg["page_size"], a, b, tr_id)
        except ApiBadRequest as e:
            if a < b:
                return split()
            state["warnings"].append(f"{kind} {a:%d.%m.%Y}, стр. {page}: {e}")
            return
        total = resp.get("total") or 0
        if page == 0 and total > MAX_WINDOW and a < b:
            return split()
        items = resp.get("items") or []
        got += [slim(i) for i in items]
        absorb(cfg, items, acc, tr_name)
        label = f"{DOC_KINDS[kind]['title']} {tr_name + ' ' if tr_name else ''}{a:%d.%m}:"
        progress(len(acc["seen"]), state["done_days"], label, unit="дн.")
        if not items or (page + 1) * cfg["page_size"] >= total:
            cache.put_window(key, got)
            return
        page += 1


def check_tech_reg_filter(client, kind, cfg, d_from, d_to):
    """Страховка: если сайт игнорирует фильтр по ТР, остановиться, а не выгрузить весь реестр под видом отобранного."""
    if not cfg["tech_regs"]:
        return
    base = client.list_page(kind, 0, 10, d_from, d_to).get("total") or 0
    one = client.list_page(kind, 0, 10, d_from, d_to, cfg["tech_regs"][0][1]).get("total") or 0
    if base > 200 and one >= base:
        raise FilterIgnored(
            f"сайт не применил фильтр по ТР для «{DOC_KINDS[kind]['title']}» "
            f"({one} из {base} документов). Пришлите разработчику вывод 'python fsa_parser.py --debug'")


def fetch_candidates(client, kind, cfg, cache, d_from, d_to, acc, state, tr=None):
    """Читает реестр окнами по нескольку дней. Результат накапливается в acc, даже если упадём на полпути."""
    cur = d_from
    while cur <= d_to:
        end = min(cur + timedelta(days=cfg["chunk_days"]), d_to)
        fetch_window(client, kind, cfg, cache, cur, end, acc, state, tr)
        state["done_days"] = (end - d_from).days
        cur = end + timedelta(days=1)
    sys.stdout.write("\r" + " " * 78 + "\r")


def fmt_int(n):
    return f"{n:,}".replace(",", " ")


def run(cfg, debug=False):
    today = date.today()
    d_from, d_to = today, today + timedelta(days=cfg["days"])
    client = FsaClient(cfg)
    cache = Cache(cfg["cache_dir"], cfg["ttl_days"])
    state = {"warnings": [], "done_days": 0, "incomplete": False, "reason": "", "missing": 0}

    def stop(reason):
        if not state["incomplete"]:
            state["reason"] = reason
        state["incomplete"] = True
        sys.stdout.write("\r" + " " * 78 + "\r")

    print(paint("-" * 64, "d"))
    print(paint("Реестр ФСА: сертификаты и декларации о соответствии", "d"))
    print()

    pool, infos, tr_by_doc = [], {}, {}
    try:
        # 1-2. Загрузка реестра и фильтр по продукции
        print(f" {paint('1.', 'b')} Загружаю реестр (истекают {d_from:%d.%m.%Y} - {d_to:%d.%m.%Y})")
        total_loaded = 0
        for kind in cfg["kinds"]:
            acc = {"seen": set(), "matched": [], "trs": {}}
            try:
                check_tech_reg_filter(client, kind, cfg, d_from, d_to)
                for tr in (cfg["tech_regs"] or [None]):          # без ТР - читаем весь реестр
                    fetch_candidates(client, kind, cfg, cache, d_from, d_to, acc, state, tr)
            finally:
                total_loaded += len(acc["seen"])
                pool += [(kind, it) for it in acc["matched"]]
                for it in acc["matched"]:
                    tr_by_doc[f"{kind}:{it['id']}"] = acc["trs"].get(it["id"], set())
                if debug and acc["matched"]:
                    Path(f"debug_list_{kind}.json").write_text(
                        json.dumps(acc["matched"][0], ensure_ascii=False, indent=2), encoding="utf-8")
        for w in state["warnings"]:
            print(paint(f"    ВНИМАНИЕ, данные могли быть неполными: {w}", "y"))
        print(f"    ... {paint(fmt_int(total_loaded) + ' документов', 'd')}")
        trs = ", ".join(n for n, _ in cfg["tech_regs"]) or "любые"
        kw = ", ".join(cfg["keywords"]) or "не используются"
        print(f" {paint('2.', 'b')} Отбор: ТР - {trs}; слова в названии продукции - {kw}")
        print(f"    ... {paint(fmt_int(len(pool)) + ' документов', 'd')}")
    except NetworkGiveUp as e:
        stop(str(e))
    except KeyboardInterrupt:
        stop("прервано пользователем (Ctrl+C)")
    except Exception as e:
        stop(f"непредвиденная ошибка: {e}")

    # 3. Карточки заявителей (руководитель, e-mail): сначала кэш, потом сайт
    print(f" {paint('3.', 'b')} Получаю данные заявителей (руководитель, e-mail)")
    online = not state["incomplete"]               # если на этапе 1 сломались - сайт больше не трогаем
    hits = fetched = 0
    try:
        for i, (kind, it) in enumerate(pool, 1):
            dk, ak = f"{kind}:{it['id']}", (applicant_key(it) if cfg["dedupe"] else None)
            info = cache.get_info(dk, ak)
            if info:
                hits += 1
            elif online:
                try:
                    det = client.detail(kind, it["id"])
                except NetworkGiveUp:
                    raise
                except Exception as e:             # одна неудача не должна ронять всю выгрузку
                    print(paint(f"\n    не удалось получить {it.get('number')}: {e}", "r"))
                    continue
                info = extract_applicant(det)
                fetched += 1
                if debug and fetched == 1:
                    Path(f"debug_detail_{kind}.json").write_text(
                        json.dumps(det, ensure_ascii=False, indent=2), encoding="utf-8")
                if any(info.values()):             # пустой результат не кэшируем
                    cache.put_info(dk, ak, info)
            if info:
                infos[dk] = info
            progress(i, len(pool), "Заявители:")
    except NetworkGiveUp as e:
        stop(str(e))
    except KeyboardInterrupt:
        stop("прервано пользователем (Ctrl+C)")
    except Exception as e:
        stop(f"непредвиденная ошибка: {e}")
    finally:
        cache.save()
    sys.stdout.write("\r" + " " * 78 + "\r")
    print(f"    ... из кэша {hits}, запрошено с сайта {fetched}")

    rows = []
    for kind, it in pool:
        row = build_row(kind, it, infos.get(f"{kind}:{it['id']}"), today, cfg["name_case"],
                        tr_by_doc.get(f"{kind}:{it['id']}", ()))
        if row["end_date"] and 0 <= row["days_left"] <= cfg["days"]:
            rows.append(row)
    rows.sort(key=lambda r: (r["days_left"], r["company"]))
    state["missing"] = sum(1 for r in rows if not r["_has_info"])
    print(f"    ... {paint(fmt_int(len(rows)) + ' документов', 'y')}")
    return rows, state


# ----------------------------------------------------------------------------
# Вывод в файлы
# ----------------------------------------------------------------------------
def fmt_value(col, v):
    if col == "end_date":
        return v.strftime("%d.%m.%Y") if v else ""
    return "" if v is None else v


def unique_path(folder, stem, ext):
    p = folder / f"{stem}.{ext}"
    n = 2
    while p.exists():                              # старые выгрузки не перезаписываем
        p = folder / f"{stem}_{n}.{ext}"
        n += 1
    return p


def save_csv(rows, cols, path):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:   # BOM - чтобы Excel не ломал кириллицу
        w = csv.writer(f, delimiter=";")
        w.writerow([COLUMN_TITLES[c] for c in cols])
        for r in rows:
            w.writerow([fmt_value(c, r[c]) for c in cols])


def save_xlsx(rows, cols, path, cfg):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Документы"
    font = Font(name="Arial", size=10)
    head_font = Font(name="Arial", size=10, bold=True)
    thin = Side(style="thin", color="D0D5DD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    red_fill = PatternFill("solid", fgColor="FDE8E8")

    ws.append([COLUMN_TITLES[c] for c in cols])
    for j in range(1, len(cols) + 1):
        c = ws.cell(row=1, column=j)
        c.font, c.border = head_font, border
        c.fill = PatternFill("solid", fgColor="EDEFF3")
        c.alignment = Alignment(vertical="center", wrap_text=True)

    for r in rows:
        ws.append([r[c] if c in ("end_date", "days_left") else fmt_value(c, r[c]) for c in cols])
        i = ws.max_row
        urgent = r["days_left"] is not None and r["days_left"] <= cfg["urgent"]
        for j, col in enumerate(cols, 1):
            cell = ws.cell(row=i, column=j)
            cell.font, cell.border = font, border
            cell.alignment = Alignment(vertical="top", wrap_text=(col in ("company", "product", "tech_reg")))
            if col == "end_date":
                cell.number_format = "DD.MM.YYYY"
            if col in ("end_date", "days_left"):
                cell.alignment = Alignment(horizontal="center", vertical="top")
            if urgent:
                cell.fill = red_fill
                if col == "days_left":
                    cell.font = Font(name="Arial", size=10, bold=True, color="C00000")

    for j, col in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(j)].width = COLUMN_WIDTHS[col]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 24
    wb.save(path)


def save(rows, cfg, partial=False):
    cols = cfg["columns"]
    folder = Path(cfg["out_dir"])
    folder.mkdir(parents=True, exist_ok=True)
    stem = f"{cfg['prefix']}_{cfg['days']}_days_{date.today():%Y-%m-%d}" + ("_частично" if partial else "")
    paths = []
    if cfg["format"] in ("xlsx", "both"):
        p = unique_path(folder, stem, "xlsx")
        save_xlsx(rows, cols, p, cfg)
        paths.append(p)
    if cfg["format"] in ("csv", "both"):
        p = unique_path(folder, stem, "csv")
        save_csv(rows, cols, p)
        paths.append(p)
    return paths


def preview(rows, cfg):
    n = min(cfg["preview"], len(rows))
    if not n:
        return
    print(f"\n Первые {n} строк для контроля:")
    cols = [c for c in cfg["columns"] if c in ("company", "number", "end_date", "days_left", "email")]
    widths = {"company": 28, "number": 32, "end_date": 11, "days_left": 6, "email": 28}
    short = {"company": "Юрлицо", "number": "Номер", "end_date": "Окончание",
             "days_left": "Дней", "email": "E-mail"}
    print(paint("  " + " ".join(short[c].ljust(widths[c]) for c in cols), "d"))
    for r in rows[:n]:
        line = " ".join(str(fmt_value(c, r[c]))[:widths[c]].ljust(widths[c]) for c in cols)
        print("  " + paint(line, "r") if r["days_left"] <= cfg["urgent"] else "  " + line)


# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Парсер реестра ФСА")
    ap.add_argument("--config", default=str(Path(__file__).with_name("config.ini")))
    ap.add_argument("--days", type=int, help="порог в днях (перекрывает config.ini)")
    ap.add_argument("--debug", action="store_true", help="сохранить сырые ответы сайта")
    ap.add_argument("--list-tr", action="store_true", help="показать ТР, доступные для фильтра")
    args = ap.parse_args()

    if args.list_tr:
        print("ТР, которые можно указать в config.ini (параметр tech_regs):\n")
        for name in TECH_REGS:
            print("  " + tech_reg_label(name))
        return

    cfg = load_config(args.config)
    if args.days:
        cfg["days"] = args.days

    rows, state = run(cfg, debug=args.debug)
    partial = state["incomplete"] or state["missing"] > 0

    if not rows:
        print(paint("\nДокументов для выгрузки нет - файл не создан.", "y"))
    else:
        print(f" {paint('4.', 'b')} Сохраняю файл")
        paths = save(rows, cfg, partial)
        print(f"    ... {paint('готово', 'g')}")
        preview(rows, cfg)
        print()
        label = "Файл ЧАСТИЧНЫЙ" if partial else "Файл готов"
        print(paint(f"Найдено {fmt_int(len(rows))} документов. {label}.", "y" if partial else "g"))
        for p in paths:
            print(paint(str(p.resolve()), "d"))

    if partial:
        print()
        if state["incomplete"]:
            print(paint(f"Выгрузка остановлена: {state['reason']}", "r"))
        if state["missing"]:
            print(paint(f"У {fmt_int(state['missing'])} документов нет руководителя/e-mail (не успели загрузить).", "y"))
        print("Всё найденное сохранено в кэше. Подождите (при блокировке - несколько часов),")
        print("запустите парсер снова - он продолжит с места остановки, а не начнёт заново.")
        sys.exit(2)


if __name__ == "__main__":
    main()