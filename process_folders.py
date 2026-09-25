# -*- coding: utf-8 -*-
"""
Обробка папок з ексель-файлами в C:\\Users\\Dembe\\Desktop\\content\\ter

Для кожної директорії, назва якої складається ТІЛЬКИ з цифр:
  1. знаходимо ексель файл;
  2. на аркуші "Шаблон" шукаємо рядки, де в колонці B стоїть число (reason_id);
  3. відправляємо на сервер UPDATE loaded_products.model SET reason_id = B WHERE art = A;
  4. видаляємо ці рядки через Excel (формули, стилі та картинки не ламаються);
  5. видаляємо колонки, у заголовку яких є "_r" — теж засобами Excel, тож
     решта даних просто зсувається вліво, а посилання у формулах Excel
     переписує сам;
  6. пакуємо папку в <назва папки>.zip поряд з нею.

Спочатку йде запит до бази, і тільки після успішного commit видаляються рядки —
щоб дані не зникли з файлу, якщо база недоступна.
"""

import argparse
import configparser
import os
import re
import sys
import zipfile
from datetime import datetime
from pathlib import Path

import psycopg2
import xlwings as xw

# ---------------------------------------------------------------- налаштування

ROOT_DIR = r"C:\Users\Dembe\Desktop\content\ter"
SHEET_NAME = "Шаблон"          # основний аркуш; якщо немає — береться перший
EXCEL_EXTS = (".xlsx", ".xlsm", ".xls")
HEADER_ROW = 1                 # рядок із назвами колонок
DROP_COL_MARKER = "_r"         # колонки з цим маркером у заголовку видаляються

DB_CONFIG_PATH = Path(__file__).with_name("db.ini")


def load_db_config() -> dict:
    """Параметри бази: db.ini поряд зі скриптом, змінні TER_DB_* мають пріоритет."""
    cfg = configparser.ConfigParser()
    cfg.read(DB_CONFIG_PATH, encoding="utf-8")
    sect = cfg["db"] if cfg.has_section("db") else {}

    def get(key, default=None):
        return os.environ.get("TER_DB_" + key.upper(), sect.get(key, default))

    db = {
        "host": get("host"),
        "port": int(get("port", 5432)),
        "dbname": get("dbname"),
        "user": get("user"),
        "password": get("password"),
        "connect_timeout": 15,
    }
    missing = [k for k, v in db.items() if v in (None, "")]
    if missing:
        raise RuntimeError("не задано параметри бази %s — заповніть %s (приклад: db.ini.example)"
                           % (", ".join(missing), DB_CONFIG_PATH))
    return db

UPDATE_SQL = 'UPDATE "loaded_products"."model" SET "reason_id" = %s WHERE "art" = %s'

LOG_PATH = Path(__file__).with_name("process_folders.log")
_log_file = None


def log(msg=""):
    print(msg)
    global _log_file
    if _log_file is None:
        _log_file = open(LOG_PATH, "a", encoding="utf-8")
        _log_file.write("\n===== %s =====\n" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    _log_file.write(str(msg) + "\n")
    _log_file.flush()


# ------------------------------------------------------------------ допоміжне

def digit_dirs(root: Path) -> list:
    """Директорії, у назві яких тільки цифри."""
    dirs = [p for p in root.iterdir() if p.is_dir() and p.name.isdigit()]
    return sorted(dirs, key=lambda p: int(p.name))


def find_excel_files(folder: Path) -> list:
    files = [
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in EXCEL_EXTS
        and not p.name.startswith("~$")
    ]
    return sorted(files)


def as_reason_id(value):
    """B -> ціле додатне число, або None якщо це не число."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        num = float(value)
    else:
        text = str(value).strip().replace(",", ".")
        if not text:
            return None
        if not re.fullmatch(r"[+-]?\d+(\.\d+)?", text):
            return None
        num = float(text)
    if not num.is_integer() or num <= 0:
        return None
    return int(num)


def as_art(value):
    """A -> код товару рядком."""
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    return text or None


def pick_sheet(book, sheet_name: str):
    for sht in book.sheets:
        if sht.name.strip().lower() == sheet_name.strip().lower():
            return sht
    return book.sheets[0]


def scan_rows(sheet):
    """Повертає (rows, bad) — rows: [(номер рядка, art, reason_id)]."""
    used = sheet.used_range
    last_row = used.last_cell.row
    rows, bad = [], []
    if last_row < 2:
        return rows, bad

    data = sheet.range((2, 1), (last_row, 2)).value
    if last_row == 2:                      # один рядок -> плоский список
        data = [data]

    for offset, pair in enumerate(data):
        row_no = offset + 2
        a_val, b_val = (list(pair) + [None, None])[:2]
        if b_val is None or (isinstance(b_val, str) and not b_val.strip()):
            continue
        reason_id = as_reason_id(b_val)
        art = as_art(a_val)
        if reason_id is None:
            bad.append((row_no, art, b_val, "у B не ціле число"))
            continue
        if not art:
            bad.append((row_no, art, b_val, "порожня комірка A"))
            continue
        rows.append((row_no, art, reason_id))
    return rows, bad


def col_letter(idx: int) -> str:
    """1 -> A, 27 -> AA."""
    letters = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def scan_columns(sheet, marker: str) -> list:
    """Колонки, у заголовку яких є marker -> [(індекс, літера, заголовок)]."""
    if not marker:
        return []
    last_col = sheet.used_range.last_cell.column
    if last_col < 1:
        return []

    header = sheet.range((HEADER_ROW, 1), (HEADER_ROW, last_col)).value
    if last_col == 1:                      # одна колонка -> скаляр
        header = [header]

    cols = []
    for offset, value in enumerate(header):
        if not isinstance(value, str):
            continue
        title = value.strip()
        if marker in title:
            cols.append((offset + 1, col_letter(offset + 1), title))
    return cols


def zip_folder(folder: Path, overwrite: bool) -> Path:
    """Пакує папку в <назва папки>.zip поряд з нею."""
    target = folder.parent / (folder.name + ".zip")
    if target.exists() and not overwrite:
        n = 1
        while True:
            candidate = folder.parent / ("%s_%d.zip" % (folder.name, n))
            if not candidate.exists():
                target = candidate
                break
            n += 1

    tmp = folder.parent / (target.name + ".part")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(folder.rglob("*")):
            if path.name.startswith("~$"):
                continue
            zf.write(path, Path(folder.name) / path.relative_to(folder))
    tmp.replace(target)
    return target


# -------------------------------------------------------------------- головне

def main():
    parser = argparse.ArgumentParser(description="Обробка папок з ексель-файлами")
    parser.add_argument("target", nargs="?",
                        help="папка з контекстного меню: коренева директорія, "
                             "або одна папка з числовою назвою")
    parser.add_argument("--root", default=ROOT_DIR, help="коренева директорія")
    parser.add_argument("--sheet", default=SHEET_NAME, help="назва аркуша з даними")
    parser.add_argument("--dry-run", action="store_true",
                        help="тільки показати, що буде зроблено (без бази, видалення і зіпів)")
    parser.add_argument("--yes", action="store_true", help="не питати підтвердження")
    parser.add_argument("--overwrite-zip", action="store_true",
                        help="перезаписувати існуючий архів (інакше створюється <назва>_1.zip)")
    parser.add_argument("--folders", nargs="*", help="обробити тільки ці папки")
    parser.add_argument("--marker", default=DROP_COL_MARKER,
                        help="маркер у заголовку колонки для видалення (типово '_r')")
    parser.add_argument("--keep-cols", action="store_true",
                        help="не видаляти колонки — обробляти тільки рядки")
    args = parser.parse_args()

    root = Path(args.target or args.root).resolve()
    if not root.is_dir():
        log("[ПОМИЛКА] Директорія не знайдена: %s" % root)
        return 1

    wanted = set(args.folders or [])
    if args.target and root.name.isdigit():
        # клік по самій папці з числовою назвою -> обробляємо тільки її
        wanted = {root.name}
        root = root.parent

    folders = digit_dirs(root)
    if wanted:
        folders = [f for f in folders if f.name in wanted]

    log("Корінь: %s" % root)
    log("Папок з числовою назвою: %d -> %s"
        % (len(folders), ", ".join(f.name for f in folders) or "—"))
    if not folders:
        return 0

    # ---------- 1. читаємо всі файли і збираємо план ----------
    plan = []          # [{folder, file, sheet, rows, bad}]
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for folder in folders:
            excels = find_excel_files(folder)
            if not excels:
                log("[ПРОПУСК] %s: ексель файл не знайдено" % folder.name)
                continue
            if len(excels) > 1:
                log("[УВАГА] %s: знайдено кілька файлів — обробляються всі: %s"
                    % (folder.name, ", ".join(e.name for e in excels)))
            for excel in excels:
                book = app.books.open(str(excel), update_links=False)
                try:
                    sheet = pick_sheet(book, args.sheet)
                    sheet_name = sheet.name
                    rows, bad = scan_rows(sheet)
                    cols = [] if args.keep_cols else scan_columns(sheet, args.marker)
                finally:
                    book.close()
                plan.append({"folder": folder, "file": excel,
                             "sheet": sheet_name, "rows": rows,
                             "bad": bad, "cols": cols})
                log("\n%s / %s [аркуш: %s] — рядків з числом у B: %d"
                    % (folder.name, excel.name, sheet_name, len(rows)))
                log("    колонок з %r у заголовку: %d" % (args.marker, len(cols)))
                for row_no, art, reason_id in rows:
                    log("    рядок %4d: art=%r  reason_id=%d" % (row_no, art, reason_id))
                for row_no, art, b_val, why in bad:
                    log("    [УВАГА] рядок %d: %s (A=%r, B=%r) — не чіпаємо"
                        % (row_no, why, art, b_val))
                for _idx, letter, title in cols:
                    log("    колонка %-3s: %r — на видалення" % (letter, title))
    finally:
        app.quit()

    total_rows = sum(len(p["rows"]) for p in plan)
    log("\nВсього рядків на оновлення та видалення: %d" % total_rows)

    log("Всього колонок на видалення: %d" % sum(len(p["cols"]) for p in plan))

    if args.dry_run:
        log("\n[DRY-RUN] Нічого не змінено: база, файли та архіви не чіпались.")
        return 0

    if not args.yes:
        try:
            answer = input("\nПродовжити? (оновлення бази + видалення рядків і колонок + zip) [y/N]: ")
        except EOFError:
            answer = ""
        if answer.strip().lower() not in ("y", "yes", "т", "так"):
            log("Скасовано користувачем.")
            return 0

    # ---------- 2. база ----------
    log("\nПідключення до бази...")
    try:
        conn = psycopg2.connect(**load_db_config())
    except Exception as exc:
        log("[ПОМИЛКА] Не вдалось підключитись до бази: %s" % exc)
        log("Файли та архіви не чіпались.")
        return 1
    conn.autocommit = False
    try:
        with conn.cursor() as cur:
            cur.execute('SELECT "reason_id" FROM "loaded_products"."reason"')
            valid_reasons = set(r[0] for r in cur.fetchall())
        log("Доступні reason_id у базі: %s" % sorted(valid_reasons))

        for item in plan:
            ok, missing, invalid = [], [], []
            for row_no, art, reason_id in item["rows"]:
                if reason_id not in valid_reasons:
                    invalid.append((row_no, art, reason_id))
                    continue
                with conn.cursor() as cur:
                    cur.execute(UPDATE_SQL, (reason_id, art))
                    if cur.rowcount > 0:
                        ok.append((row_no, art, reason_id, cur.rowcount))
                    else:
                        missing.append((row_no, art, reason_id))
            conn.commit()
            item["ok"] = ok
            item["missing"] = missing
            item["invalid"] = invalid
            log("\n%s / %s: оновлено %d, не знайдено в базі %d, невірний reason_id %d"
                % (item["folder"].name, item["file"].name, len(ok), len(missing), len(invalid)))
            for row_no, art, reason_id in missing:
                log("    [УВАГА] рядок %d: art=%r немає в базі — рядок НЕ видаляється"
                    % (row_no, art))
            for row_no, art, reason_id in invalid:
                log("    [УВАГА] рядок %d: reason_id=%d немає в таблиці reason "
                    "— рядок НЕ видаляється" % (row_no, reason_id))
    except Exception as exc:
        conn.rollback()
        log("[ПОМИЛКА] Робота з базою перервана, зміни відкочені: %s" % exc)
        log("Файли та архіви не чіпались.")
        return 1
    finally:
        conn.close()

    # ---------- 3. видалення рядків і колонок ----------
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    deleted_by_file = {}
    dropped_by_file = {}
    try:
        for item in plan:
            to_delete = item.get("ok", [])
            to_drop = item["cols"]
            if not to_delete and not to_drop:
                log("\n%s / %s: немає що видаляти" % (item["folder"].name, item["file"].name))
                deleted_by_file[item["file"]] = 0
                dropped_by_file[item["file"]] = 0
                continue
            book = app.books.open(str(item["file"]), update_links=False)
            deleted = 0
            dropped = 0
            try:
                sheet = pick_sheet(book, item["sheet"])
                # знизу вгору, щоб не з'їхали номери рядків
                for row_no, art, reason_id, _ in sorted(to_delete, reverse=True):
                    current = as_art(sheet.range("A%d" % row_no).value)
                    if current != art:
                        log("    [УВАГА] рядок %d: у файлі зараз art=%r, очікувалось %r "
                            "— пропускаємо видалення" % (row_no, current, art))
                        continue
                    sheet.range("%d:%d" % (row_no, row_no)).api.Delete()
                    deleted += 1
                # колонки — справа наліво, щоб не з'їхали індекси
                for idx, letter, title in sorted(to_drop, reverse=True):
                    current = sheet.range((HEADER_ROW, idx)).value
                    if isinstance(current, str):
                        current = current.strip()
                    if current != title:
                        log("    [УВАГА] колонка %s: у файлі зараз заголовок %r, "
                            "очікувався %r — пропускаємо видалення"
                            % (letter, current, title))
                        continue
                    sheet.range((HEADER_ROW, idx)).api.EntireColumn.Delete()
                    dropped += 1
                book.save()
            finally:
                book.close()
            deleted_by_file[item["file"]] = deleted
            dropped_by_file[item["file"]] = dropped
            log("\n%s / %s: видалено рядків — %d"
                % (item["folder"].name, item["file"].name, deleted))
            log("%s / %s: видалено колонок — %d"
                % (item["folder"].name, item["file"].name, dropped))
    finally:
        app.quit()

    # ---------- 4. архіви ----------
    log("")
    for folder in dict.fromkeys(p["folder"] for p in plan):
        try:
            archive = zip_folder(folder, args.overwrite_zip)
            size_mb = archive.stat().st_size / 1024 / 1024
            log("Архів: %s (%.1f MB)" % (archive.name, size_mb))
        except Exception as exc:
            log("[ПОМИЛКА] Не вдалось заархівувати %s: %s" % (folder.name, exc))

    # ---------- підсумок ----------
    log("\n===== ПІДСУМОК =====")
    for item in plan:
        problems = (len(item.get("missing", [])) + len(item.get("invalid", []))
                    + len(item["bad"]))
        log("%s / %s: оновлено в базі %d, видалено рядків %d, "
            "видалено колонок %d, проблемних %d"
            % (item["folder"].name, item["file"].name, len(item.get("ok", [])),
               deleted_by_file.get(item["file"], 0),
               dropped_by_file.get(item["file"], 0), problems))
    log("Лог: %s" % LOG_PATH)
    return 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    sys.exit(main())
