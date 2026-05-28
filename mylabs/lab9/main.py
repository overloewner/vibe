"""
Лабораторная работа №9 — Создание GUI-приложения
Предметная область: Книжный магазин
Инструмент: Tkinter
Дисциплина: Скриптовые языки программирования
"""
import tkinter as tk
from tkinter import messagebox, ttk


# ====================================================================
# ДАННЫЕ КАТАЛОГА КНИГ
# ====================================================================

books = [
    {"title": "Мастер и Маргарита",      "author": "Булгаков М.А.",     "genre": "Роман",          "year": 1967, "price": 450.0},
    {"title": "Преступление и наказание", "author": "Достоевский Ф.М.",  "genre": "Роман",          "year": 1866, "price": 380.0},
    {"title": "Война и мир",              "author": "Толстой Л.Н.",      "genre": "Роман-эпопея",   "year": 1869, "price": 620.0},
    {"title": "Евгений Онегин",           "author": "Пушкин А.С.",       "genre": "Роман в стихах", "year": 1833, "price": 290.0},
    {"title": "Тихий Дон",               "author": "Шолохов М.А.",      "genre": "Роман-эпопея",   "year": 1940, "price": 510.0},
    {"title": "Отцы и дети",             "author": "Тургенев И.С.",     "genre": "Роман",          "year": 1862, "price": 320.0},
    {"title": "Обломов",                  "author": "Гончаров И.А.",     "genre": "Роман",          "year": 1859, "price": 355.0},
    {"title": "Собачье сердце",           "author": "Булгаков М.А.",     "genre": "Повесть",        "year": 1925, "price": 275.0},
    {"title": "Капитанская дочка",        "author": "Пушкин А.С.",       "genre": "Роман",          "year": 1836, "price": 310.0},
    {"title": "Анна Каренина",            "author": "Толстой Л.Н.",      "genre": "Роман",          "year": 1878, "price": 490.0},
]

GENRES = sorted(set(b["genre"] for b in books))


# ====================================================================
# ФУНКЦИИ ОБРАБОТЧИКИ СОБЫТИЙ
# ====================================================================

def search_books():
    """Выполняет поиск книг по выбранному критерию."""
    query = search_entry.get().strip().lower()
    mode = search_mode.get()

    result_list.delete(0, tk.END)

    if not query:
        messagebox.showwarning("Внимание", "Введите поисковый запрос.")
        return

    found = []
    for b in books:
        if mode == "title"  and query in b["title"].lower():
            found.append(b)
        elif mode == "author" and query in b["author"].lower():
            found.append(b)
        elif mode == "genre"  and query in b["genre"].lower():
            found.append(b)

    if not found:
        result_list.insert(tk.END, "— Книги по запросу не найдены —")
    else:
        for b in found:
            result_list.insert(
                tk.END,
                f"{b['title']}  |  {b['author']}  |  {b['genre']}  |  {b['year']} г.  |  {b['price']:.0f} руб."
            )

    status_label.config(text=f"Найдено: {len(found)} книг(и)")


def show_all():
    """Выводит весь каталог книг в список результатов."""
    result_list.delete(0, tk.END)
    for b in books:
        result_list.insert(
            tk.END,
            f"{b['title']}  |  {b['author']}  |  {b['genre']}  |  {b['year']} г.  |  {b['price']:.0f} руб."
        )
    status_label.config(text=f"Всего книг в каталоге: {len(books)}")


def clear_search():
    """Очищает поле поиска и список результатов."""
    search_entry.delete(0, tk.END)
    result_list.delete(0, tk.END)
    status_label.config(text=f"Всего книг в каталоге: {len(books)}")


def add_book():
    """Добавляет новую книгу в каталог после валидации полей."""
    title  = title_entry.get().strip()
    author = author_entry.get().strip()
    genre  = genre_var.get().strip()
    year_s = year_entry.get().strip()
    price_s = price_entry.get().strip()

    # Валидация обязательных строковых полей
    if not title or not author or not genre:
        messagebox.showerror("Ошибка", "Заполните все обязательные поля:\nНазвание, Автор, Жанр.")
        return

    # Валидация года
    try:
        year = int(year_s)
        if year < 1000 or year > 2100:
            raise ValueError
    except ValueError:
        messagebox.showerror("Ошибка", "Год издания должен быть целым числом\nв диапазоне 1000–2100.")
        return

    # Валидация цены
    try:
        price = float(price_s)
        if price <= 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("Ошибка", "Цена должна быть положительным числом.")
        return

    books.append({"title": title, "author": author, "genre": genre, "year": year, "price": price})
    messagebox.showinfo("Успешно", f"Книга «{title}» добавлена в каталог.")

    # Очищаем поля формы
    for widget in (title_entry, author_entry, year_entry, price_entry):
        widget.delete(0, tk.END)

    # Обновляем выпадающий список жанров
    all_genres = sorted(set(b["genre"] for b in books))
    genre_combo["values"] = all_genres

    # Обновляем справочную информацию
    refresh_info()
    status_label.config(text=f"Всего книг в каталоге: {len(books)}")


def refresh_info():
    """Обновляет блок справочной информации о жанрах."""
    info_text.config(state="normal")
    info_text.delete("1.0", tk.END)
    genre_counts = {}
    for b in books:
        genre_counts[b["genre"]] = genre_counts.get(b["genre"], 0) + 1
    for g, cnt in sorted(genre_counts.items()):
        info_text.insert(tk.END, f"• {g}: {cnt} кн.\n")
    info_text.config(state="disabled")


# ====================================================================
# ПОСТРОЕНИЕ ИНТЕРФЕЙСА
# ====================================================================

root = tk.Tk()
root.title("Книжный магазин — Каталог")
root.geometry("1000x580")
root.resizable(False, False)

BG       = "#f0f4f8"
ACCENT   = "#4a90d9"
BTN_FG   = "white"
FONT_MAIN  = ("Arial", 11)
FONT_BOLD  = ("Arial", 12, "bold")
FONT_TITLE = ("Arial", 14, "bold")

root.configure(bg=BG)
root.columnconfigure(0, weight=3)
root.columnconfigure(1, weight=1)
root.rowconfigure(0, weight=1)

# ── ЛЕВАЯ ПАНЕЛЬ (поиск и результаты) ──────────────────────────────
left = tk.Frame(root, bg=BG)
left.grid(row=0, column=0, sticky="nsew", padx=15, pady=15)

tk.Label(left, text="Книжный магазин — Каталог книг",
         font=FONT_TITLE, bg=BG, fg=ACCENT).pack(pady=(0, 8))

# Строка поиска
sf = tk.Frame(left, bg=BG)
sf.pack(fill="x")
tk.Label(sf, text="Поиск:", font=FONT_BOLD, bg=BG).pack(side="left")
search_entry = tk.Entry(sf, font=FONT_MAIN, width=30)
search_entry.pack(side="left", padx=6)
# Нажатие Enter — тоже запускает поиск
search_entry.bind("<Return>", lambda e: search_books())
tk.Button(sf, text="Найти", command=search_books,
          bg=ACCENT, fg=BTN_FG, font=FONT_MAIN, relief="flat", padx=10).pack(side="left", padx=3)
tk.Button(sf, text="✖ Очистить", command=clear_search,
          bg="#e57373", fg=BTN_FG, font=FONT_MAIN, relief="flat", padx=6).pack(side="left", padx=3)

# Радиокнопки режима поиска
mf = tk.Frame(left, bg=BG)
mf.pack(fill="x", pady=5)
search_mode = tk.StringVar(value="title")
tk.Label(mf, text="Искать по:", font=FONT_MAIN, bg=BG).pack(side="left")
for val, lbl in [("title", "Названию"), ("author", "Автору"), ("genre", "Жанру")]:
    tk.Radiobutton(mf, text=lbl, variable=search_mode, value=val,
                   bg=BG, font=FONT_MAIN).pack(side="left", padx=8)

tk.Button(left, text="Показать все книги", command=show_all,
          bg="#66bb6a", fg=BTN_FG, font=FONT_MAIN, relief="flat", padx=8).pack(pady=(0, 6))

# Список результатов со скроллом
lf = tk.Frame(left, bg=BG)
lf.pack(fill="both", expand=True)
sb = tk.Scrollbar(lf)
sb.pack(side="right", fill="y")
result_list = tk.Listbox(lf, font=FONT_MAIN, yscrollcommand=sb.set,
                          height=16, selectbackground=ACCENT, activestyle="none")
result_list.pack(fill="both", expand=True)
sb.config(command=result_list.yview)

status_label = tk.Label(left, text=f"Всего книг в каталоге: {len(books)}",
                         font=("Arial", 10, "italic"), bg=BG, fg="#555")
status_label.pack(pady=(4, 0))

# ── ПРАВАЯ ПАНЕЛЬ (добавление книги + справка) ──────────────────────
right = tk.Frame(root, bg=BG, relief="groove", bd=1)
right.grid(row=0, column=1, sticky="nsew", padx=(0, 15), pady=15)

tk.Label(right, text="Добавить книгу", font=FONT_BOLD, bg=BG, fg=ACCENT).pack(pady=(10, 5))

# Поля формы добавления
for label_txt, var_name in [("Название *", "title_entry"),
                              ("Автор *",   "author_entry"),
                              ("Год *",     "year_entry"),
                              ("Цена (руб.) *", "price_entry")]:
    tk.Label(right, text=label_txt, font=FONT_MAIN, bg=BG, anchor="w").pack(fill="x", padx=12, pady=(3, 0))
    e = tk.Entry(right, font=FONT_MAIN)
    e.pack(fill="x", padx=12, pady=(0, 3))
    globals()[var_name] = e

tk.Label(right, text="Жанр *", font=FONT_MAIN, bg=BG, anchor="w").pack(fill="x", padx=12, pady=(3, 0))
genre_var = tk.StringVar(value=GENRES[0] if GENRES else "")
genre_combo = ttk.Combobox(right, textvariable=genre_var, values=GENRES,
                            font=FONT_MAIN, state="normal")
genre_combo.pack(fill="x", padx=12, pady=(0, 8))

tk.Button(right, text="Добавить в каталог", command=add_book,
          bg=ACCENT, fg=BTN_FG, font=FONT_BOLD, relief="flat", pady=5).pack(fill="x", padx=12, pady=4)

# Справочный блок — жанры
tk.Label(right, text="─" * 28, bg=BG, fg="#ccc").pack(pady=4)
tk.Label(right, text="Жанры в каталоге:", font=FONT_BOLD, bg=BG).pack(padx=12, anchor="w")
info_text = tk.Text(right, font=("Arial", 10), height=8,
                     wrap="word", bg="#e8f0fe", relief="flat", state="normal")
info_text.pack(fill="x", padx=12, pady=5)
refresh_info()


# ====================================================================
# ЗАПУСК ПРИЛОЖЕНИЯ
# ====================================================================
show_all()
root.mainloop()
