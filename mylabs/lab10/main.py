"""
Лабораторная работа №10 — Автоматизация задач
Предметная область: Книжный магазин (обработка файлов заказов)
Модули: pathlib, shutil, sys, logging, csv, json, time
Дисциплина: Скриптовые языки программирования
"""
import sys
import csv
import json
import time
import shutil
import logging
from pathlib import Path
from datetime import datetime


# ====================================================================
# КОНФИГУРАЦИЯ
# ====================================================================

DEFAULT_SOURCE = Path("orders")
DEFAULT_TARGET = Path("processed_orders")
ARCHIVE_LIFETIME_DAYS = 30

EXTENSIONS = {
    "Заказы":   {".csv"},
    "Отчёты":   {".txt", ".log"},
    "Данные":   {".json"},
    "Архивы":   {".zip", ".tar", ".gz", ".rar"},
}
TRASH_EXTENSIONS = {".tmp", ".bak", ".old"}


# ====================================================================
# ЛОГИРОВАНИЕ
# ====================================================================

def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)-8s | %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler("automation.log", encoding="utf-8"),
        ],
    )


def log(message, level="info", preview=False):
    prefix = "[PREVIEW] " if preview else ""
    getattr(logging, level)(prefix + message)


# ====================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ====================================================================

def get_category(ext):
    """Определяет категорию файла по его расширению."""
    for category, extensions in EXTENSIONS.items():
        if ext in extensions:
            return category
    return "Другое"


def generate_unique_name(directory, name, ext):
    """Генерирует уникальное имя файла, добавляя числовой суффикс при конфликте."""
    base = Path(name).stem.lower().replace(" ", "_")
    counter = 1
    new_name = f"{base}{ext}"
    while (directory / new_name).exists():
        new_name = f"{base}_{counter}{ext}"
        counter += 1
    return new_name


def is_file_old(file_path, days):
    """Проверяет, превышает ли возраст файла заданное количество дней."""
    return (time.time() - file_path.stat().st_mtime) > days * 86400


def parse_csv_order(file_path):
    """Читает CSV-файл заказа и возвращает количество строк и итоговую сумму."""
    total = 0.0
    rows = 0
    try:
        with open(file_path, encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                try:
                    price = float(row.get("price", 0) or 0)
                    qty   = int(row.get("quantity", 1) or 1)
                    total += price * qty
                    rows += 1
                except (ValueError, KeyError):
                    pass
    except Exception:
        pass
    return rows, total


def generate_summary_report(processed_orders, target_dir, preview=False):
    """Генерирует сводный текстовый отчёт по всем обработанным CSV-заказам."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = target_dir / f"summary_{timestamp}.txt"

    grand_total = sum(o["total"] for o in processed_orders)
    lines = [
        "=" * 60,
        "СВОДНЫЙ ОТЧЁТ ОБРАБОТКИ ЗАКАЗОВ",
        f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}",
        "=" * 60,
        f"Обработано файлов заказов: {len(processed_orders)}",
        "",
        "ДЕТАЛИ:",
    ]
    for item in processed_orders:
        lines.append(
            f"  {item['name']:35s}  {item['rows']:3d} строк   {item['total']:10.2f} руб."
        )
    lines += [
        "",
        f"ИТОГОВАЯ СУММА ВСЕХ ЗАКАЗОВ: {grand_total:.2f} руб.",
        "=" * 60,
    ]

    if not preview:
        report_path.write_text("\n".join(lines), encoding="utf-8")
        log(f"Сводный отчёт сохранён: {report_path.name}")
    else:
        log(f"Будет создан сводный отчёт: {report_path.name}", preview=True)
        for line in lines:
            log(f"  {line}", preview=True)

    return grand_total


# ====================================================================
# СОЗДАНИЕ ТЕСТОВЫХ ДАННЫХ
# ====================================================================

def create_test_data(source_dir):
    """Создаёт тестовые файлы заказов для демонстрации работы скрипта."""
    source_dir.mkdir(parents=True, exist_ok=True)

    # CSV-заказы
    orders = [
        ("order_001.csv", [(1, "Мастер и Маргарита", 2, 450), (2, "Война и мир", 1, 620)]),
        ("order_002.csv", [(3, "Евгений Онегин", 3, 290), (4, "Анна Каренина", 2, 490)]),
        ("order_003.csv", [(5, "Тихий Дон", 1, 510), (6, "Обломов", 2, 355), (7, "Преступление и наказание", 1, 380)]),
    ]
    for filename, rows in orders:
        with open(source_dir / filename, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["book_id", "title", "quantity", "price"])
            for row in rows:
                writer.writerow(row)

    # Текстовые отчёты
    (source_dir / "report_january.txt").write_text(
        "Отчёт за январь 2026\nПродано книг: 150\nВыручка: 67 500 руб.", encoding="utf-8"
    )
    (source_dir / "report_february.txt").write_text(
        "Отчёт за февраль 2026\nПродано книг: 178\nВыручка: 82 100 руб.", encoding="utf-8"
    )

    # JSON с данными магазина
    data = {
        "store": "Книжный магазин",
        "updated": datetime.now().strftime("%Y-%m-%d"),
        "top_books": ["Мастер и Маргарита", "Война и мир", "Преступление и наказание"],
    }
    (source_dir / "store_data.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # Файлы-мусор (будут удалены)
    (source_dir / "temp.tmp").write_text("temporary data")
    (source_dir / "backup.bak").write_text("old backup")

    # Пустой CSV (будет удалён)
    (source_dir / "empty.csv").write_text("")

    log(f"Тестовые данные созданы в папке: {source_dir}")
    log(f"  Создано {len(orders)} файла заказов, 2 отчёта, 1 JSON, 2 мусорных файла, 1 пустой файл")


# ====================================================================
# ОСНОВНАЯ ЛОГИКА ОБРАБОТКИ
# ====================================================================

def process_files(source_dir, target_dir, mode, preview):
    stats = {"deleted": 0, "moved": 0, "errors": 0}
    type_stats = {}
    processed_orders = []

    if not source_dir.exists():
        log("Исходная папка не найдена. Создаём тестовые данные...", "warning")
        create_test_data(source_dir)

    if not preview:
        target_dir.mkdir(parents=True, exist_ok=True)

    log(f"Режим: {mode.upper()} | Источник: {source_dir} | Цель: {target_dir}")
    log("-" * 60)

    for file_path in sorted(source_dir.rglob("*")):
        if not file_path.is_file():
            continue
        # Не обрабатываем файлы внутри целевой директории
        if target_dir in file_path.parents:
            continue

        ext = file_path.suffix.lower()

        # ── Удаление мусорных файлов ────────────────────────────────
        if ext in TRASH_EXTENSIONS:
            log(f"Удаление временного файла ({ext}): {file_path.name}", preview=preview)
            if not preview:
                file_path.unlink()
                stats["deleted"] += 1
            continue

        # ── Проверка размера ────────────────────────────────────────
        try:
            size = file_path.stat().st_size
        except FileNotFoundError:
            continue

        if size == 0:
            log(f"Удаление пустого файла: {file_path.name}", preview=preview)
            if not preview:
                file_path.unlink()
                stats["deleted"] += 1
            continue

        # ── Определение категории ───────────────────────────────────
        category = get_category(ext)

        # ── Удаление устаревших архивов ─────────────────────────────
        if category == "Архивы" and is_file_old(file_path, ARCHIVE_LIFETIME_DAYS):
            log(f"Удаление устаревшего архива (>{ARCHIVE_LIFETIME_DAYS} дней): {file_path.name}", preview=preview)
            if not preview:
                file_path.unlink()
                stats["deleted"] += 1
            continue

        # ── Подготовка целевой директории ───────────────────────────
        category_path = target_dir / category
        if not preview:
            category_path.mkdir(parents=True, exist_ok=True)

        new_name = generate_unique_name(category_path, file_path.name, ext)
        log(f"Перемещение [{category}]: {file_path.name}  →  {new_name}", preview=preview)

        # ── Анализ содержимого CSV-заказов ──────────────────────────
        if ext == ".csv":
            rows, total = parse_csv_order(file_path)
            processed_orders.append({"name": new_name, "rows": rows, "total": total})
            log(f"  Содержимое CSV: {rows} строк заказа, сумма: {total:.2f} руб.", preview=preview)

        # ── Перемещение файла ───────────────────────────────────────
        if not preview:
            try:
                shutil.move(str(file_path), str(category_path / new_name))
                stats["moved"] += 1
                type_stats[category] = type_stats.get(category, 0) + 1
            except (OSError, shutil.Error) as e:
                log(f"Ошибка при перемещении {file_path.name}: {e}", "error")
                stats["errors"] += 1
        else:
            type_stats[category] = type_stats.get(category, 0) + 1

    # ── Сводный отчёт по заказам ────────────────────────────────────
    if processed_orders:
        orders_dir = target_dir / "Заказы"
        if not preview:
            orders_dir.mkdir(parents=True, exist_ok=True)
        grand_total = generate_summary_report(processed_orders, orders_dir, preview)
        log(f"Итоговая сумма всех заказов: {grand_total:.2f} руб.")

    return stats, type_stats


# ====================================================================
# ВЫВОД ИТОГОВОЙ СТАТИСТИКИ
# ====================================================================

def print_statistics(stats, type_stats, mode, source_dir, target_dir):
    print("\n" + "=" * 60)
    print("ИТОГОВАЯ СТАТИСТИКА ОБРАБОТКИ")
    print("=" * 60)
    print(f"Режим работы  : {mode.upper()}")
    print(f"Источник      : {source_dir}")
    print(f"Цель          : {target_dir}")
    print(f"Удалено файлов: {stats['deleted']}")
    print(f"Перемещено    : {stats['moved']}")
    print(f"Ошибок        : {stats['errors']}")
    if type_stats:
        print("\nПО КАТЕГОРИЯМ:")
        for key in sorted(type_stats):
            print(f"  {key:<20s}: {type_stats[key]} файл(ов)")
    print("=" * 60)


# ====================================================================
# ТОЧКА ВХОДА
# ====================================================================

if __name__ == "__main__":
    setup_logging()

    args = sys.argv
    mode = args[1].lower() if len(args) > 1 else "run"

    if mode not in {"run", "preview"}:
        print("Использование: python main.py [run|preview] [источник] [цель]")
        print("  run     — реальное выполнение операций (по умолчанию)")
        print("  preview — предварительный просмотр без изменений в файловой системе")
        sys.exit(1)

    source_dir = Path(args[2]) if len(args) > 2 else DEFAULT_SOURCE
    target_dir = Path(args[3]) if len(args) > 3 else DEFAULT_TARGET
    preview    = (mode == "preview")

    stats, type_stats = process_files(source_dir, target_dir, mode, preview)
    print_statistics(stats, type_stats, mode, source_dir, target_dir)
