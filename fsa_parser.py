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
    },
    "declarations": {
        "title": "Декларация",
        "list_url": f"{BASE}/api/v1/rds/common/declarations/get",
        "detail_url": f"{BASE}/api/v1/rds/common/declarations/{{id}}",
    },
}
STATUS_ACTIVE = 6  # статус «Действует»

COLUMN_TITLES = {
    "company": "Наименование юрлица",
    "head": "ФИО руководителя",
    "number": "Номер документа",
    "doc_type": "Тип документа",
    "end_date": "Дата окончания",
    "days_left": "Дней до конца",
    "email": "Электронный адрес",
    "product": "Наименование продукции",
}
COLUMN_WIDTHS = {"company": 34, "head": 26, "number": 34, "doc_type": 14,
                 "end_date": 16, "days_left": 14, "email": 30, "product": 50}

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


# ----------------------------------------------------------------------------
# Настройки
# ----------------------------------------------------------------------------
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
        "kinds": [k for k in lst("filter", "document_types", "certificates,declarations")
                  if k in DOC_KINDS],
        "format": cp.get("output", "format", fallback="xlsx").strip().lower(),
        "out_dir": cp.get("output", "output_dir", fallback="выгрузки").strip(),
        "prefix": cp.get("output", "file_prefix", fallback="fsa").strip(),
        "columns": [c for c in lst("output", "columns") if c in COLUMN_TITLES]
                   or list(COLUMN_TITLES),
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

    def list_page(self, kind, page, size, d_from, d_to):
        body = {
            "size": size,
            "page": page,
            "filter": {
                "status": [STATUS_ACTIVE],
                "endDate": {"minDate": d_from.isoformat(), "maxDate": d_to.isoformat()},
                "regDate": {"minDate": None, "maxDate": None},
                "idCertType": [], "idObjectType": [], "idProductType": [],
                "idGroupRU": [], "idGroupEEU": [], "idTechReg": [], "idApplicantType": [],
                "idCertScheme": [], "idDeclScheme": [], "idProductOrigin": [],
                "idProductEEU": [], "idProductRU": [],
                "columnsSearch": [],
                "awaitForApprove": None, "awaitOperatorCheck": None, "editApp": None,
                "violationSend": None, "isProtected": None, "isNotAccredited": None,
            },
            "columnsSort": [{"column": "endDate", "sort": "ASC"}],
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


def build_row(kind, item, info, today):
    info = info or {}
    end = parse_date(first_key(item, "certEndDate", "declEndDate", "endDate", suffix="EndDate"))
    company = info.get("company") or first_key(item, "applicantName") or ""
    product = first_key(item, "productFullName", "productName") or ""
    return {
        "company": str(company).strip(),
        "head": info.get("head", ""),
        "number": str(first_key(item, "number") or "").strip(),
        "doc_type": DOC_KINDS[kind]["title"],
        "end_date": end,
        "days_left": (end - today).days if end else None,
        "email": info.get("email", ""),
        "product": str(product).strip().replace("\n", " "),
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

    def save(self):
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self.data, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, self.path)                     # запись атомарная - файл не испортится при обрыве
        self._dirty = 0

    def put_window(self, key, items):
        self.data["windows"][key] = items
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


def absorb(cfg, items, acc):
    for it in items:
        if it.get("id") in acc["seen"]:
            continue
        acc["seen"].add(it.get("id"))
        name = str(first_key(it, "productFullName", "productName") or "").lower()
        if not cfg["keywords"] or any(k in name for k in cfg["keywords"]):
            acc["matched"].append(it)


def fetch_window(client, kind, cfg, cache, a, b, acc, state):
    """Читает все страницы за окно дат [a, b]. При лимитах сайта делит окно пополам."""
    key = f"{kind}|{a}|{b}|{date.today()}"
    cached = cache.data["windows"].get(key)
    if cached is not None:                       # окно уже загружено ранее - сайт не трогаем
        absorb(cfg, cached, acc)
        return

    def split():
        span = (b - a).days
        if span >= 2:
            mid = a + timedelta(days=span // 2)
            parts = [(a, mid), (mid, b)]         # перекрытие на 1 день - дубли снимает seen
        else:
            parts = [(a, a), (b, b)]
        for x, y in parts:
            fetch_window(client, kind, cfg, cache, x, y, acc, state)

    page, got = 0, []
    while True:
        try:
            resp = client.list_page(kind, page, cfg["page_size"], a, b)
        except ApiBadRequest as e:
            if a < b:
                return split()
            state["warnings"].append(f"{kind} {a:%d.%m.%Y}, стр. {page}: {e}")
            return
        total = resp.get("total") or 0
        if page == 0 and total > MAX_WINDOW and a < b:
            return split()
        items = resp.get("items") or []
        got += items
        absorb(cfg, items, acc)
        progress(len(acc["seen"]), state["done_days"], f"{DOC_KINDS[kind]['title']} {a:%d.%m}:", unit="дн.")
        if not items or (page + 1) * cfg["page_size"] >= total:
            cache.put_window(key, got)
            return
        page += 1


def fetch_candidates(client, kind, cfg, cache, d_from, d_to, acc, state):
    """Читает реестр окнами по нескольку дней. Результат накапливается в acc, даже если упадём на полпути."""
    cur = d_from
    while cur <= d_to:
        end = min(cur + timedelta(days=cfg["chunk_days"]), d_to)
        fetch_window(client, kind, cfg, cache, cur, end, acc, state)
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

    pool, infos = [], {}
    try:
        # 1-2. Загрузка реестра и фильтр по продукции
        print(f" {paint('1.', 'b')} Загружаю реестр (истекают {d_from:%d.%m.%Y} - {d_to:%d.%m.%Y})")
        total_loaded = 0
        for kind in cfg["kinds"]:
            acc = {"seen": set(), "matched": []}
            try:
                fetch_candidates(client, kind, cfg, cache, d_from, d_to, acc, state)
            finally:
                total_loaded += len(acc["seen"])
                pool += [(kind, it) for it in acc["matched"]]
                if debug and acc["matched"]:
                    Path(f"debug_list_{kind}.json").write_text(
                        json.dumps(acc["matched"][0], ensure_ascii=False, indent=2), encoding="utf-8")
        for w in state["warnings"]:
            print(paint(f"    ВНИМАНИЕ, данные могли быть неполными: {w}", "y"))
        print(f"    ... {paint(fmt_int(total_loaded) + ' документов', 'd')}")
        kw = ", ".join(cfg["keywords"]) or "без фильтра"
        print(f" {paint('2.', 'b')} Отбираю нужную продукцию ({kw})")
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
        row = build_row(kind, it, infos.get(f"{kind}:{it['id']}"), today)
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
            cell.alignment = Alignment(vertical="top", wrap_text=(col in ("company", "product")))
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
    args = ap.parse_args()

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
