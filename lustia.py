import os
import sys
import time
import json
import requests
from datetime import datetime
from pathlib import Path
import argparse
import configparser
import glob

try:
    from tqdm import tqdm
    from colorama import init, Fore, Style
    init(autoreset=True)
except ImportError:
    print("Установи зависимости: pip install requests tqdm colorama")
    sys.exit(1)

# ========================== КОНФИГУРАЦИЯ ==========================

BASE_URL = "http://localhost:8000/api"
API_KEY = ""  # Не нужен для локального сервера

# Кеш для мгновенного поиска
cache = {}
CACHE_TTL = 300  # 5 минут

# Конфиг для загрузчика
CONFIG_FILE = Path.home() / ".dbuploader.ini"
SUPPORTED = {".csv", ".json", ".xlsx", ".xls", ".txt"}
CHUNK_SIZE = 1024 * 1024  # 1 МБ
MAX_RETRIES = 3
RETRY_DELAY = 3  # сек

# ========================== ЦВЕТА ==========================

class Colors:
    RED = '\033[91m'
    DARK_RED = '\033[31m'
    BLOOD_RED = '\033[38;5;88m'
    WHITE = '\033[97m'
    GRAY = '\033[90m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    CYAN = '\033[96m'
    MAGENTA = '\033[95m'
    BOLD = '\033[1m'
    END = '\033[0m'

def c(text, color=Colors.WHITE):
    return f"{color}{text}{Colors.END}"

def clear():
    os.system('cls' if os.name == 'nt' else 'clear')

# Цветные сообщения для загрузчика
def ok(msg): print(Fore.GREEN + "✅ " + msg)
def err(msg): print(Fore.RED + "❌ " + msg)
def info(msg): print(Fore.CYAN + "ℹ " + msg)
def warn(msg): print(Fore.YELLOW + "⚠ " + msg)

def hr():
    print(Fore.RED + Style.DIM + "─" * 70)

def fmt_size(n: int) -> str:
    """Форматирование размера в байтах в человекочитаемый вид"""
    if n == 0:
        return "0 Б"
    for unit in ["Б", "КБ", "МБ", "ГБ", "ТБ"]:
        if abs(n) < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} ПБ"

# ========================== БАННЕР ==========================

BANNER = f"""
{c('╭───────────────────╮', Colors.DARK_RED)}	{c('╭───────────────────────────────────────────────────────────────────────────────╮', Colors.DARK_RED)} {c('╭───────────────────╮', Colors.DARK_RED)}
{c('│', Colors.DARK_RED)}		      {c('│', Colors.DARK_RED)}	{c('│', Colors.DARK_RED)}    	{c('▄▄▄', Colors.RED)}      {c('▄▄▄', Colors.DARK_RED)}  {c('▄▄▄', Colors.RED)}  {c('▄▄▄▄▄▄▄', Colors.DARK_RED)} {c('▄▄▄▄▄▄▄▄▄', Colors.RED)} {c('▄▄▄▄▄', Colors.DARK_RED)} {c('▄▄▄▄▄▄▄▄▄', Colors.RED)} {c('▄▄▄▄▄', Colors.DARK_RED)}   {c('▄▄▄▄', Colors.RED)}  	{c('│', Colors.DARK_RED)} {c('│  by @frameworkq', Colors.GRAY)}   {c('│', Colors.DARK_RED)}
{c('│', Colors.DARK_RED)}		      {c('│', Colors.DARK_RED)}	{c('│', Colors.DARK_RED)}        {c('███', Colors.RED)}      {c('███', Colors.DARK_RED)}  {c('███', Colors.RED)} {c('█████▀▀▀', Colors.DARK_RED)} {c('▀▀▀███▀▀▀', Colors.RED)}  {c('███', Colors.DARK_RED)}  {c('▀▀▀███▀▀▀', Colors.RED)}  {c('███', Colors.DARK_RED)}  {c('▄██▀▀██▄', Colors.RED)}	{c('│', Colors.DARK_RED)} {c('│  and @t1mott', Colors.GRAY)}      {c('│', Colors.DARK_RED)}
{c('│', Colors.DARK_RED)}		      {c('│', Colors.DARK_RED)}	{c('│', Colors.DARK_RED)}    	{c('███', Colors.RED)}      {c('███', Colors.DARK_RED)}  {c('███', Colors.RED)}  {c('▀████▄', Colors.DARK_RED)}     {c('███', Colors.RED)}     {c('███', Colors.DARK_RED)}     {c('███', Colors.RED)}     {c('███', Colors.DARK_RED)}  {c('███', Colors.RED)}  {c('███', Colors.DARK_RED)} 	{c('│', Colors.DARK_RED)} {c('│', Colors.DARK_RED)}                   {c('│', Colors.DARK_RED)}
{c('│', Colors.DARK_RED)}		      {c('│', Colors.DARK_RED)}	{c('│', Colors.DARK_RED)}    	 {c('███', Colors.RED)}      {c('███▄▄███', Colors.DARK_RED)}    {c('▀████', Colors.RED)}    {c('███', Colors.DARK_RED)}     {c('███', Colors.RED)}     {c('███', Colors.DARK_RED)}  {c('███▀▀███', Colors.RED)} 	{c('│', Colors.DARK_RED)} {c('│  DataBase - 41TB', Colors.GRAY)}  {c('│', Colors.DARK_RED)}
{c('│', Colors.DARK_RED)}		      {c('│', Colors.DARK_RED)}	{c('│', Colors.DARK_RED)}       {c('████████', Colors.RED)} {c('▀██████▀', Colors.DARK_RED)} {c('███████▀', Colors.RED)}    {c('███', Colors.DARK_RED)}    {c('▄███▄', Colors.RED)}    {c('███', Colors.DARK_RED)}    {c('▄███▄', Colors.RED)} {c('███', Colors.DARK_RED)}  {c('███', Colors.RED)}	{c('│', Colors.DARK_RED)} {c('│', Colors.DARK_RED)}                   {c('│', Colors.DARK_RED)}
{c('│', Colors.DARK_RED)}		      {c('│', Colors.DARK_RED)}	{c('│', Colors.DARK_RED)}    	{c('│', Colors.DARK_RED)} {c('│', Colors.DARK_RED)}	{c('│', Colors.DARK_RED)}    	{c('│', Colors.DARK_RED)} {c('│  version - 1.0', Colors.GRAY)}   {c('│', Colors.DARK_RED)}
{c('╰───────────────────╯', Colors.DARK_RED)}	{c('╰───────────────────────────────────────────────────────────────────────────────╯', Colors.DARK_RED)} {c('╰───────────────────╯', Colors.DARK_RED)}	

{c('╭───────────────────────────────────────────────────────╮', Colors.DARK_RED)}
{c('│', Colors.DARK_RED)}  {c('[1]', Colors.RED)} - Поиск по ФИО      {c('│', Colors.DARK_RED)}  {c('[4]', Colors.RED)} - Поиск по ИНН        {c('│', Colors.DARK_RED)}
{c('│', Colors.DARK_RED)}  {c('[2]', Colors.RED)} - Поиск по номеру   {c('│', Colors.DARK_RED)}  {c('[5]', Colors.RED)} - Поиск по паспорту	{c('│', Colors.DARK_RED)}        
{c('│', Colors.DARK_RED)}  {c('[3]', Colors.RED)} - Поиск по ном.авто {c('│', Colors.DARK_RED)}  {c('[6]', Colors.RED)} - Поиск по адресу	{c('│', Colors.DARK_RED)}
{c('╰───────────────────────────────────────────────────────╯', Colors.DARK_RED)}
{c('╭───────────────────────────────────────────────────────╮', Colors.DARK_RED)}
{c('│', Colors.DARK_RED)}  {c('[7]', Colors.RED)} - Поиск по почте    {c('│', Colors.DARK_RED)}  {c('[10]', Colors.RED)} - Скоро...           {c('│', Colors.DARK_RED)}
{c('│', Colors.DARK_RED)}  {c('[8]', Colors.RED)} - Поиск по нику     {c('│', Colors.DARK_RED)}  {c('[11]', Colors.RED)} - Скоро...           {c('│', Colors.DARK_RED)}        
{c('│', Colors.DARK_RED)}  {c('[9]', Colors.RED)} - Multisearch       {c('│', Colors.DARK_RED)}  {c('[12]', Colors.RED)} - Я боюсь нато	{c('│', Colors.DARK_RED)}
{c('│', Colors.DARK_RED)}  {c('[13]', Colors.RED)} - DB Uploader      {c('│', Colors.DARK_RED)}  {c('[14]', Colors.RED)} - Список баз         {c('│', Colors.DARK_RED)}
{c('╰───────────────────────────────────────────────────────╯', Colors.DARK_RED)}
"""

# ========================== DB UPLOADER API КЛИЕНТ ==========================

def load_config() -> dict:
    cfg = configparser.ConfigParser()
    if CONFIG_FILE.exists():
        cfg.read(CONFIG_FILE)
    return {
        "base_url": cfg.get("api", "base_url", fallback=BASE_URL),
        "api_key": cfg.get("api", "api_key", fallback=API_KEY),
    }

def save_config(base_url: str, api_key: str):
    cfg = configparser.ConfigParser()
    cfg["api"] = {"base_url": base_url.rstrip("/"), "api_key": api_key}
    with open(CONFIG_FILE, "w") as f:
        cfg.write(f)
    ok(f"Конфиг сохранён в {CONFIG_FILE}")

class DBClient:
    def __init__(self, base_url: str, api_key: str):
        self.base_url = base_url.rstrip("/")
        self.headers = {"x-api-key": api_key}
        self.session = requests.Session()
        self.session.headers.update(self.headers)

    def _get(self, path: str, params=None, timeout=30) -> dict:
        url = f"{self.base_url}{path}"
        r = self.session.get(url, params=params, timeout=timeout)
        r.raise_for_status()
        return r.json()

    def _delete(self, path: str, timeout=15) -> dict:
        url = f"{self.base_url}{path}"
        r = self.session.delete(url, timeout=timeout)
        r.raise_for_status()
        return r.json()

    def _post(self, path: str, data=None, files=None, timeout=30) -> dict:
        url = f"{self.base_url}{path}"
        r = self.session.post(url, data=data, files=files, timeout=timeout)
        r.raise_for_status()
        return r.json()

    def list_databases(self) -> list:
        data = self._get("/databases")
        return data.get("databases", [])

    def get_database_info(self, name: str) -> dict:
        """Получение информации о конкретной базе включая размер"""
        try:
            data = self._get(f"/databases/{requests.utils.quote(name, safe='')}")
            return data
        except requests.HTTPError as e:
            if e.response.status_code == 404:
                return {}
            raise

    def search_db(self, query: str, db: str = "", limit: int = 100) -> dict:
        params = {"q": query, "limit": limit}
        if db:
            params["db"] = db
        return self._get("/search", params=params, timeout=60)

    def delete_database(self, name: str) -> bool:
        try:
            self._delete(f"/databases/{requests.utils.quote(name, safe='')}")
            return True
        except requests.HTTPError as e:
            if e.response.status_code == 404:
                return False
            raise

    def upload(self, filepath: Path, name: str = "") -> dict:
        """Загрузка файла любого размера со стримингом и прогресс-баром."""
        file_size = filepath.stat().st_size
        upload_name = name or filepath.stem[:100]
        url = f"{self.base_url}/upload"

        print()
        info(f"Файл:  {filepath.name}")
        info(f"Размер: {fmt_size(file_size)}")
        info(f"База:  {upload_name}")
        hr()

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                with open(filepath, "rb") as f:
                    with tqdm(
                        total=file_size,
                        unit="B",
                        unit_scale=True,
                        unit_divisor=1024,
                        desc=f"  Загрузка (попытка {attempt})",
                        colour="red",
                        bar_format="{desc}: {percentage:3.0f}% |{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]",
                    ) as bar:
                        class ProgressFile:
                            def __init__(self, fobj, pbar):
                                self._f = fobj
                                self._pbar = pbar
                            def read(self, size=-1):
                                chunk = self._f.read(size if size > 0 else CHUNK_SIZE)
                                self._pbar.update(len(chunk))
                                return chunk
                            def __len__(self):
                                return file_size

                        pf = ProgressFile(f, bar)
                        resp = self._post(
                            "/upload",
                            data={"name": upload_name},
                            files={"file": (filepath.name, pf)}
                        )

                return resp

            except requests.exceptions.ConnectionError:
                err(f"Нет соединения с сервером (попытка {attempt}/{MAX_RETRIES})")
            except requests.exceptions.Timeout:
                err(f"Таймаут (попытка {attempt}/{MAX_RETRIES})")
            except requests.exceptions.HTTPError as e:
                try:
                    detail = e.response.json().get("error", e.response.text)
                except Exception:
                    detail = str(e)
                err(f"HTTP {e.response.status_code}: {detail}")
                break
            except (KeyboardInterrupt, SystemExit):
                print()
                warn("Загрузка прервана")
                sys.exit(0)
            except Exception as e:
                err(f"Ошибка: {e} (попытка {attempt}/{MAX_RETRIES})")

            if attempt < MAX_RETRIES:
                print(Fore.YELLOW + f"  Повтор через {RETRY_DELAY} сек...")
                time.sleep(RETRY_DELAY)

        return {}

# ========================== УЛЬТРАБЫСТРЫЙ ПОИСК ==========================

def search(query: str, limit: int = 1000) -> dict:
    """Ультрабыстрый поиск с кешированием"""
    if not query.strip():
        return {"error": "Пустой запрос"}

    # Проверка кеша
    cache_key = f"{query.lower().strip()}_{limit}"
    if cache_key in cache:
        cache_time, cache_data = cache[cache_key]
        if time.time() - cache_time < CACHE_TTL:
            return cache_data

    headers = {"x-api-key": API_KEY}
    params = {"q": query, "limit": limit}

    start_time = time.time()

    try:
        resp = requests.get(
            f"{BASE_URL}/search",
            params=params,
            headers=headers,
            timeout=30
        )

        elapsed_ms = round((time.time() - start_time) * 1000, 2)

        if resp.status_code == 200:
            data = resp.json()
            data["tookMs"] = elapsed_ms
            cache[cache_key] = (time.time(), data)
            return data
        elif resp.status_code == 401:
            return {"error": "❌ Неверный ключ API", "tookMs": elapsed_ms}
        elif resp.status_code == 404:
            return {"error": "❌ Эндпоинт не найден", "tookMs": elapsed_ms}
        else:
            return {"error": f"HTTP {resp.status_code}", "tookMs": elapsed_ms}

    except requests.exceptions.ConnectionError:
        return {"error": "❌ Сервер недоступен", "tookMs": 0}
    except requests.exceptions.Timeout:
        return {"error": "⏱️ Таймаут", "tookMs": 0}
    except Exception as e:
        return {"error": str(e), "tookMs": 0}

# ========================== ПОИСК ВСЕХ ЗАПИСЕЙ ==========================

def search_all(query: str, limit: int = 1000) -> dict:
    """Поиск всех записей с автоматической пагинацией"""
    if not query.strip():
        return {"error": "Пустой запрос"}
    
    all_results = []
    page = 0
    page_size = 1000
    total_count = None
    
    headers = {"x-api-key": API_KEY}
    start_time = time.time()
    
    print(c("\n⏳ Начинаю поиск всех записей...", Colors.YELLOW))
    
    try:
        while True:
            params = {
                "q": query, 
                "limit": page_size,
                "offset": page * page_size
            }
            
            resp = requests.get(
                f"{BASE_URL}/search",
                params=params,
                headers=headers,
                timeout=60
            )
            
            if resp.status_code != 200:
                if resp.status_code == 401:
                    return {"error": "❌ Неверный ключ API"}
                elif resp.status_code == 404:
                    return {"error": "❌ Эндпоинт не найден"}
                else:
                    return {"error": f"HTTP {resp.status_code}"}
            
            data = resp.json()
            results = data.get('results', [])
            
            if not results:
                break
            
            all_results.extend(results)
            
            if total_count is None:
                total_count = data.get('count', 0)
            
            print(c(f"  Загружено {len(all_results)} записей...", Colors.GRAY))
            
            if total_count and len(all_results) >= total_count:
                break
            
            if len(results) < page_size:
                break
            
            page += 1
            
            if len(all_results) > 1000000:
                break
        
        elapsed_ms = round((time.time() - start_time) * 1000, 2)
        
        print(c(f"✅ Найдено {len(all_results)} записей", Colors.GREEN))
        
        return {
            "results": all_results,
            "count": len(all_results),
            "totalCount": total_count,
            "tookMs": elapsed_ms
        }
        
    except requests.exceptions.ConnectionError:
        return {"error": "❌ Сервер недоступен", "tookMs": 0}
    except requests.exceptions.Timeout:
        return {"error": "⏱️ Таймаут", "tookMs": 0}
    except Exception as e:
        return {"error": str(e), "tookMs": 0}

def search_with_limit(query: str) -> dict:
    """Поиск с выбором количества записей"""
    if not query.strip():
        return {"error": "Пустой запрос"}
    
    print(c("\n[?] Сколько записей загрузить?", Colors.CYAN))
    print(c("    [1] - 100 записей (быстро)", Colors.GRAY))
    print(c("    [2] - 1000 записей (стандарт)", Colors.GRAY))
    print(c("    [3] - 10000 записей (медленно)", Colors.GRAY))
    print(c("    [4] - ВСЕ записи (может быть очень медленно)", Colors.RED))
    
    choice = input(c("\n[+] Выберите опцию (1-4): ", Colors.RED)).strip()
    
    limit_map = {
        '1': 100,
        '2': 1000,
        '3': 10000,
        '4': None
    }
    
    limit = limit_map.get(choice, 1000)
    
    if limit is None:
        return search_all(query)
    else:
        return search(query, limit)

# ========================== JSON ВЫВОД ==========================

def print_json(data, query):
    """Красивый JSON вывод"""
    clear()
    
    print("\n" + c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
    print(c("  🔴 РЕЗУЛЬТАТ ПОИСКА", Colors.RED + Colors.BOLD))
    print(c("  Запрос: " + query, Colors.GRAY))
    
    if "tookMs" in data:
        speed_color = Colors.GREEN if data["tookMs"] < 10 else Colors.YELLOW if data["tookMs"] < 50 else Colors.RED
        print(c(f"  Время: {data['tookMs']} мс", speed_color))
    
    print(c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
    
    if "error" in data:
        print(c(f"\n{data['error']}", Colors.RED))
        print(c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
        return
    
    results = data.get('results', [])
    total_count = data.get('totalCount', len(results))
    
    print(c(f"\n📊 Найдено: {total_count} записей", Colors.GREEN))
    print(c(f"📋 Загружено: {len(results)} записей", Colors.CYAN))
    
    if not results:
        print(c("\n❌ Ничего не найдено", Colors.RED))
        print(c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
        return
    
    # Если результатов много, спрашиваем, показывать ли все
    show_all_results = results
    if len(results) > 100:
        show_all = input(c("\n[?] Показать все {len(results)} записей? (да/нет): ", Colors.YELLOW)).strip().lower()
        if show_all not in ['да', 'yes', 'y', 'д']:
            print(c("\n📋 Показаны первые 100 записей", Colors.CYAN))
            show_all_results = results[:100]
    
    # JSON вывод
    print(c("\n📋 JSON:", Colors.CYAN))
    
    if len(results) > 500:
        print(c("[!] Слишком много записей для отображения. Сохраните результат в файл.", Colors.YELLOW))
    else:
        display_data = data.copy()
        display_data["results"] = show_all_results
        print(json.dumps(display_data, indent=2, ensure_ascii=False))
    
    # Таблица с результатами
    print(c("\n" + "═" * 70, Colors.GRAY))
    print(c(f"  📊 ЗАПИСИ", Colors.GREEN))
    print(c("═" * 70, Colors.GRAY))
    
    for idx, item in enumerate(show_all_results, 1):
        print(c(f"\n[{idx}] ", Colors.DARK_RED + Colors.BOLD))
        if isinstance(item, dict):
            for key, val in item.items():
                if key.lower() in ['file', 'filename', 'database', 'db', 'databaseid', 'rank', 'id', 'datasetname', 'datasetid']:
                    continue
                if isinstance(val, dict):
                    print(f"  {c(key + ':', Colors.RED)}")
                    for k, v in val.items():
                        print(f"      {k}: {v}")
                else:
                    if val:
                        print(f"  {c(key + ':', Colors.RED)} {val}")
        else:
            print(f"  {item}")
        print(c("  " + "─" * 50, Colors.DARK_RED))
    
    print(c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
    
    save = input(c("\n[+] Сохранить результат? (да/нет): ", Colors.GREEN)).strip().lower()
    if save in ['да', 'yes', 'y']:
        filename = f"search_{query}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(c(f"\n✅ Сохранено в: {filename}", Colors.GREEN))

# ========================== ФУНКЦИИ МЕНЮ ==========================

def show_menu():
    clear()
    print(BANNER)

def func_10():
    clear()
    print("\n" + c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
    print(c("  🔴 ФУНКЦИЯ 10 — Скоро...", Colors.RED + Colors.BOLD))
    print(c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
    print(c("\n[!] Эта функция будет доступна в следующей версии", Colors.YELLOW))
    input(c("\n[+] Нажмите Enter...", Colors.GRAY))

def func_11():
    clear()
    print("\n" + c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
    print(c("  🔴 ФУНКЦИЯ 11 — Скоро...", Colors.RED + Colors.BOLD))
    print(c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
    print(c("\n[!] Эта функция будет доступна в следующей версии", Colors.YELLOW))
    input(c("\n[+] Нажмите Enter...", Colors.GRAY))

def func_12():
    clear()
    print("\n" + c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
    print(c("  🔴 Я БОЮСЬ НАТО", Colors.RED + Colors.BOLD))
    print(c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
    print(c("\n[!] Нато — это Организация Североатлантического договора", Colors.YELLOW))
    print(c("[!] Шутка! Просто пасхалка", Colors.GRAY))
    print(c("[!] Автор: @frameworkq и @t1mott", Colors.GRAY))
    input(c("\n[+] Нажмите Enter...", Colors.GRAY))

def multisearch():
    clear()
    print("\n" + c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
    print(c("  🔴 MULTISEARCH", Colors.RED + Colors.BOLD))
    print(c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
    
    queries_input = input(c("\n[+] Введите запросы через запятую: ", Colors.RED)).strip()
    if not queries_input:
        return
    
    queries = [q.strip() for q in queries_input.split(',') if q.strip()]
    all_results = []
    
    for q in queries:
        print(c(f"\n[*] Поиск: {q}", Colors.CYAN))
        result = search(q)
        if result.get('results'):
            all_results.extend(result['results'])
    
    if all_results:
        print_json({"ok": True, "results": all_results, "count": len(all_results), "tookMs": 1}, "multisearch")
    else:
        print(c("\n❌ Ничего не найдено", Colors.RED))
    
    input(c("\n[+] Нажмите Enter...", Colors.GRAY))

# ========================== DB UPLOADER ФУНКЦИИ ==========================

def get_files_bulk(path_input: str) -> list:
    """Получение всех файлов из папки или по шаблону"""
    files = []
    
    # Если это папка
    if os.path.isdir(path_input):
        dir_path = Path(path_input)
        for f in dir_path.iterdir():
            if f.is_file() and f.suffix.lower() in SUPPORTED:
                files.append(f)
        return files
    
    # Если это шаблон с *
    if '*' in path_input or '?' in path_input:
        for pattern_path in glob.glob(path_input):
            p = Path(pattern_path)
            if p.is_file() and p.suffix.lower() in SUPPORTED:
                files.append(p)
        return files
    
    # Если это конкретный файл
    p = Path(path_input)
    if p.exists() and p.is_file() and p.suffix.lower() in SUPPORTED:
        return [p]
    
    return files

def db_uploader():
    cfg = load_config()
    
    if not cfg["base_url"] or not cfg["api_key"]:
        cfg["base_url"] = BASE_URL
        cfg["api_key"] = API_KEY
    
    client = DBClient(cfg["base_url"], cfg["api_key"])
    
    clear()
    print("\n" + c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
    print(c("  🔴 DB UPLOADER — Пакетная загрузка баз", Colors.RED + Colors.BOLD))
    print(c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
    
    print(c("\n📌 Введите путь:", Colors.CYAN))
    print(c("  • Папка: /home/user/data/          - загрузит ВСЕ файлы", Colors.GRAY))
    print(c("  • Шаблон: *.csv                    - все CSV файлы", Colors.GRAY))
    print(c("  • Файл: data.csv                   - конкретный файл", Colors.GRAY))
    print(c(f"  • Поддерживаемые форматы: {', '.join(sorted(SUPPORTED))}", Colors.GRAY))
    
    path_input = input(c("\n[+] Путь: ", Colors.RED)).strip()
    
    if not path_input:
        print(c("\n❌ Путь не указан", Colors.YELLOW))
        input(c("\n[+] Нажмите Enter...", Colors.GRAY))
        return
    
    # Получаем файлы
    files = get_files_bulk(path_input)
    
    if not files:
        print(c(f"\n❌ Файлы не найдены: {path_input}", Colors.YELLOW))
        input(c("\n[+] Нажмите Enter...", Colors.GRAY))
        return
    
    # Показываем найденные файлы
    print(c(f"\n📁 Найдено файлов: {len(files)}", Colors.GREEN))
    for f in files[:10]:
        size = fmt_size(f.stat().st_size)
        print(c(f"  • {f.name} ({size})", Colors.WHITE))
    if len(files) > 10:
        print(c(f"  ... и ещё {len(files) - 10} файлов", Colors.GRAY))
    
    total_size = sum(f.stat().st_size for f in files)
    print(c(f"\n💾 Общий размер: {fmt_size(total_size)}", Colors.CYAN))
    
    confirm = input(c("\n[+] Начать ПАКЕТНУЮ загрузку ВСЕХ файлов? (да/нет): ", Colors.RED)).strip().lower()
    if confirm not in ['да', 'yes', 'y', 'д']:
        print(c("\n❌ Загрузка отменена", Colors.YELLOW))
        input(c("\n[+] Нажмите Enter...", Colors.GRAY))
        return
    
    # Пакетная загрузка
    print()
    hr()
    print(c("  ⏳ ЗАГРУЗКА {} ФАЙЛОВ...".format(len(files)), Colors.RED + Colors.BOLD))
    hr()
    
    success_count = 0
    failed = []
    total_rows = 0
    total_uploaded_size = 0
    
    for i, filepath in enumerate(files, 1):
        print()
        print(c(f"[{i}/{len(files)}] {filepath.name}", Colors.CYAN + Colors.BOLD))
        
        # Название базы = имя файла (без расширения)
        db_name = filepath.stem[:100]
        
        result = client.upload(filepath, db_name)
        
        if result.get("ok"):
            db = result.get("database", {})
            rows = db.get('rowCount', 0)
            total_rows += rows
            total_uploaded_size += filepath.stat().st_size
            print(c(f"  ✅ {db_name} ({rows:,} строк)", Colors.GREEN))
            success_count += 1
        else:
            error = result.get('error', 'Ошибка')
            print(c(f"  ❌ {error}", Colors.RED))
            failed.append(filepath.name)
    
    # Итог
    print()
    hr()
    print(c("  📊 ИТОГ ПАКЕТНОЙ ЗАГРУЗКИ", Colors.RED + Colors.BOLD))
    hr()
    print(c(f"  ✅ Успешно загружено: {success_count}/{len(files)}", Colors.GREEN))
    print(c(f"  📊 Всего строк: {total_rows:,}", Colors.CYAN))
    print(c(f"  💾 Загружено данных: {fmt_size(total_uploaded_size)}", Colors.CYAN))
    
    if failed:
        print(c(f"\n  ❌ Не удалось загрузить:", Colors.RED))
        for f in failed:
            print(c(f"    • {f}", Colors.RED))
    
    input(c("\n[+] Нажмите Enter...", Colors.GRAY))

def db_list():
    cfg = load_config()
    if not cfg["base_url"] or not cfg["api_key"]:
        cfg["base_url"] = BASE_URL
        cfg["api_key"] = API_KEY
    
    client = DBClient(cfg["base_url"], cfg["api_key"])
    
    clear()
    print("\n" + c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
    print(c("  📋 СПИСОК БАЗ ДАННЫХ", Colors.RED + Colors.BOLD))
    print(c("═" * 70, Colors.BLOOD_RED + Colors.BOLD))
    
    try:
        dbs = client.list_databases()
    except Exception as e:
        print(c(f"\n❌ Ошибка: {e}", Colors.RED))
        input(c("\n[+] Нажмите Enter...", Colors.GRAY))
        return
    
    if not dbs:
        print(c("\n❌ Нет загруженных баз", Colors.YELLOW))
    else:
        total_size = 0
        
        for i, db in enumerate(dbs, 1):
            db_name = db.get("name", "—")
            created = db.get("createdAt", "")[:10]
            cols = db.get("columns", [])
            row_count = db.get("rowCount", 0)
            
            # Получаем размер базы
            db_size = 0
            try:
                db_info = client.get_database_info(db_name)
                if db_info:
                    db_size = db_info.get("size", 0)
                    if not db_size:
                        db_size = db_info.get("sizeBytes", 0)
                    if not db_size:
                        db_size = db_info.get("fileSize", 0)
            except Exception:
                db_size = row_count * 500 + len(cols) * 100
            
            total_size += db_size
            
            col_preview = ", ".join(cols[:5]) + (f" +{len(cols)-5}" if len(cols) > 5 else "")
            
            print(c(f"\n[{i}] ", Colors.DARK_RED + Colors.BOLD) + c(db_name, Colors.WHITE + Colors.BOLD))
            print(c(f"    📊 {row_count:,} строк  |  📅 {created}", Colors.GRAY))
            print(c(f"    🗂  {col_preview}", Colors.GRAY))
            
            size_str = fmt_size(db_size)
            if db_size > 100 * 1024 * 1024:
                size_color = Colors.RED
            elif db_size > 10 * 1024 * 1024:
                size_color = Colors.YELLOW
            else:
                size_color = Colors.GREEN
            
            print(c(f"    💾 {size_str}", size_color))
        
        print(c(f"\n📊 Всего: {len(dbs)} баз", Colors.GREEN))
        print(c(f"💾 Общий размер: {fmt_size(total_size)}", Colors.CYAN))
    
    input(c("\n[+] Нажмите Enter...", Colors.GRAY))

# ========================== ОСНОВНАЯ ФУНКЦИЯ ==========================

def main():
    labels = {
        '1': 'ФИО', '2': 'номер телефона', '3': 'номер автомобиля',
        '4': 'ИНН', '5': 'паспорт', '6': 'адрес', '7': 'почту', '8': 'никнейм'
    }
    
    while True:
        show_menu()
        choice = input(c("\n[+] Выберите действие: ", Colors.RED)).strip()
        
        if choice == '0':
            print(c("\n[!] Выход", Colors.RED))
            sys.exit(0)
        
        if choice in labels:
            query = input(c(f"\n[+] Введите {labels[choice]}: ", Colors.RED)).strip()
            if query:
                result = search_with_limit(query)
                print_json(result, query)
            else:
                print(c("\n[!] Пустой запрос", Colors.YELLOW))
            input(c("\n[+] Нажмите Enter...", Colors.GRAY))
        
        elif choice == '9':
            multisearch()
        
        elif choice == '10':
            func_10()
        elif choice == '11':
            func_11()
        elif choice == '12':
            func_12()
        elif choice == '13':
            db_uploader()
        elif choice == '14':
            db_list()
        
        else:
            print(c("[!] Неверный выбор", Colors.RED))
            input(c("\n[+] Нажмите Enter...", Colors.GRAY))

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(c("\n[!] Выход", Colors.RED))
        sys.exit(0)