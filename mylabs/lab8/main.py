"""
Лабораторная работа №8 — Анализ датасета с Pandas и визуализация
Предметная область: Продажи книг (книжный магазин)
Инструменты: Pandas, Matplotlib, Seaborn, Scikit-learn (Random Forest)
Дисциплина: Скриптовые языки программирования
"""

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")  # Без отображения окон
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score, classification_report
from sklearn.preprocessing import LabelEncoder
import io
import os

# ────────────────────────────────────────────────
# 1. ГЕНЕРАЦИЯ СИНТЕТИЧЕСКОГО ДАТАСЕТА
# ────────────────────────────────────────────────

def generate_dataset(n_rows: int = 220, seed: int = 42) -> pd.DataFrame:
    """Генерирует синтетический датасет о продажах книг."""
    rng = np.random.default_rng(seed)

    genres = ["Роман", "Фантастика", "Детектив", "История", "Бизнес", "Наука"]
    genre_weights = [0.25, 0.20, 0.20, 0.15, 0.10, 0.10]

    genre_col = rng.choice(genres, size=n_rows, p=genre_weights)

    # Генерация с учётом жанра (фантастика и детективы продаются лучше)
    genre_sales_multiplier = {
        "Роман": 1.0, "Фантастика": 1.5, "Детектив": 1.4,
        "История": 0.7, "Бизнес": 0.9, "Наука": 0.6,
    }

    base_price = {
        "Роман": 500, "Фантастика": 600, "Детектив": 550,
        "История": 700, "Бизнес": 800, "Наука": 900,
    }

    year = rng.integers(2000, 2026, size=n_rows)
    pages = rng.integers(100, 800, size=n_rows)
    price = np.array([
        base_price[g] * rng.uniform(0.7, 1.5)
        for g in genre_col
    ]).round(2)
    rating = np.clip(
        np.array([rng.normal(3.8, 0.8) for _ in genre_col]),
        1.0, 5.0
    ).round(1)
    author_rating = np.clip(rng.normal(3.5, 1.0, size=n_rows), 1.0, 5.0).round(1)

    # Продажи зависят от жанра, рейтинга и цены
    sales_count = np.array([
        max(0, int(
            genre_sales_multiplier[g] * 1000
            * (rating[i] / 3.5)
            * (600 / price[i])
            * rng.uniform(0.5, 2.0)
        ))
        for i, g in enumerate(genre_col)
    ])

    # Бестселлер: топ 20% по продажам
    threshold = np.percentile(sales_count, 80)
    is_bestseller = (sales_count >= threshold).astype(int)

    # Намеренно добавляем пропущенные значения (~5%)
    pages_with_nan = pages.astype(float)
    nan_idx = rng.choice(n_rows, size=int(n_rows * 0.05), replace=False)
    pages_with_nan[nan_idx] = np.nan

    author_rating_nan = author_rating.copy()
    nan_idx2 = rng.choice(n_rows, size=int(n_rows * 0.04), replace=False)
    author_rating_nan[nan_idx2] = np.nan

    titles = [f"Книга_{i+1:03d}" for i in range(n_rows)]

    df = pd.DataFrame({
        "book_id": range(1, n_rows + 1),
        "title": titles,
        "genre": genre_col,
        "year": year,
        "price": price,
        "pages": pages_with_nan,
        "rating": rating,
        "sales_count": sales_count,
        "author_rating": author_rating_nan,
        "is_bestseller": is_bestseller,
    })

    return df


# ────────────────────────────────────────────────
# 2. СОХРАНЕНИЕ И ЗАГРУЗКА CSV
# ────────────────────────────────────────────────

def save_and_load(df: pd.DataFrame, csv_path: str) -> pd.DataFrame:
    """Сохраняет датасет в CSV и снова загружает его."""
    df.to_csv(csv_path, index=False, encoding="utf-8")
    print(f"[OK] Датасет сохранён: {csv_path}")
    loaded = pd.read_csv(csv_path)
    print(f"[OK] Датасет загружен: {loaded.shape[0]} строк, {loaded.shape[1]} столбцов")
    return loaded


# ────────────────────────────────────────────────
# 3. БАЗОВЫЙ АНАЛИЗ
# ────────────────────────────────────────────────

def basic_analysis(df: pd.DataFrame) -> None:
    """Выводит базовую информацию о датасете."""
    print("\n" + "=" * 60)
    print("  1. БАЗОВЫЙ АНАЛИЗ ДАТАСЕТА")
    print("=" * 60)

    print("\n>>> df.head(5):")
    print(df.head(5).to_string(index=False))

    print(f"\n>>> df.shape: {df.shape}")
    print(f"    Строк: {df.shape[0]}, Столбцов: {df.shape[1]}")

    print("\n>>> Типы данных (df.dtypes):")
    for col, dtype in df.dtypes.items():
        print(f"    {col:<20} {dtype}")

    print("\n>>> Пропущенные значения (df.isnull().sum()):")
    nulls = df.isnull().sum()
    for col, cnt in nulls.items():
        if cnt > 0:
            pct = cnt / len(df) * 100
            print(f"    {col:<20} {cnt} ({pct:.1f}%)")
    if nulls.sum() == 0:
        print("    Пропущенных значений нет")


# ────────────────────────────────────────────────
# 4. ПРЕДОБРАБОТКА
# ────────────────────────────────────────────────

def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    """Обрабатывает пропуски, проверяет дубликаты."""
    print("\n" + "=" * 60)
    print("  2. ПРЕДОБРАБОТКА")
    print("=" * 60)

    df_clean = df.copy()

    # Заполнение пропусков медианой
    for col in ["pages", "author_rating"]:
        median_val = df_clean[col].median()
        missing_cnt = df_clean[col].isnull().sum()
        df_clean[col] = df_clean[col].fillna(median_val)
        print(f"[OK] {col}: заполнено {missing_cnt} пропусков медианой ({median_val:.1f})")

    # Дубликаты
    dup_count = df_clean.duplicated().sum()
    print(f"\n[OK] Дубликатов строк: {dup_count}")
    if dup_count > 0:
        df_clean = df_clean.drop_duplicates()
        print(f"     Удалено дубликатов: {dup_count}")

    print("\n>>> df.describe():")
    desc = df_clean[["price", "pages", "rating", "sales_count", "author_rating"]].describe()
    print(desc.round(2).to_string())

    return df_clean


# ────────────────────────────────────────────────
# 5. ВИЗУАЛИЗАЦИЯ
# ────────────────────────────────────────────────

def create_visualizations(df: pd.DataFrame, output_dir: str) -> None:
    """Создаёт и сохраняет графики."""
    print("\n" + "=" * 60)
    print("  3. ВИЗУАЛИЗАЦИЯ")
    print("=" * 60)

    os.makedirs(output_dir, exist_ok=True)
    sns.set_theme(style="whitegrid", palette="muted")

    # 1. Корреляционная тепловая карта
    print("\n[~] Корреляционная тепловая карта...")
    fig, ax = plt.subplots(figsize=(9, 7))
    numeric_cols = ["price", "pages", "rating", "sales_count", "author_rating", "year"]
    corr = df[numeric_cols].corr()
    sns.heatmap(
        corr, annot=True, fmt=".2f", cmap="coolwarm",
        center=0, square=True, ax=ax,
        linewidths=0.5, cbar_kws={"shrink": 0.8}
    )
    ax.set_title("Корреляционная матрица числовых признаков", fontsize=14, pad=12)
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=120)
    plt.close(fig)
    print("[OK] Корреляционная тепловая карта создана (в буфер)")

    # 2. Диаграмма рассеяния: цена vs продажи по жанру
    print("[~] Диаграмма рассеяния (price vs sales_count)...")
    fig, ax = plt.subplots(figsize=(10, 6))
    genres = df["genre"].unique()
    palette = sns.color_palette("tab10", len(genres))
    for i, g in enumerate(genres):
        mask = df["genre"] == g
        ax.scatter(df.loc[mask, "price"], df.loc[mask, "sales_count"],
                   label=g, alpha=0.65, s=50, color=palette[i])
    ax.set_xlabel("Цена (руб.)", fontsize=12)
    ax.set_ylabel("Количество продаж", fontsize=12)
    ax.set_title("Зависимость продаж от цены по жанрам", fontsize=14)
    ax.legend(title="Жанр", bbox_to_anchor=(1.01, 1), loc="upper left")
    plt.tight_layout()
    buf2 = io.BytesIO()
    plt.savefig(buf2, format="png", dpi=120)
    plt.close(fig)
    print("[OK] Диаграмма рассеяния создана (в буфер)")

    # 3. Boxplot: продажи по жанру
    print("[~] Ящик с усами (sales_count by genre)...")
    fig, ax = plt.subplots(figsize=(10, 6))
    genre_order = df.groupby("genre")["sales_count"].median().sort_values(ascending=False).index
    sns.boxplot(data=df, x="genre", y="sales_count",
                order=genre_order, palette="Set2", ax=ax)
    ax.set_xlabel("Жанр", fontsize=12)
    ax.set_ylabel("Количество продаж", fontsize=12)
    ax.set_title("Распределение продаж по жанрам (boxplot)", fontsize=14)
    plt.xticks(rotation=30, ha="right")
    plt.tight_layout()
    buf3 = io.BytesIO()
    plt.savefig(buf3, format="png", dpi=120)
    plt.close(fig)
    print("[OK] Ящик с усами создан (в буфер)")

    # 4. KDE: распределение рейтинга
    print("[~] KDE-распределение рейтинга...")
    fig, ax = plt.subplots(figsize=(8, 5))
    sns.kdeplot(data=df, x="rating", hue="is_bestseller",
                fill=True, alpha=0.4, common_norm=False, ax=ax,
                palette={0: "steelblue", 1: "tomato"})
    ax.set_xlabel("Рейтинг книги", fontsize=12)
    ax.set_ylabel("Плотность", fontsize=12)
    ax.set_title("KDE: Распределение рейтинга (бестселлер vs обычная)", fontsize=14)
    legend = ax.get_legend()
    if legend:
        legend.set_title("Бестселлер")
    plt.tight_layout()
    buf4 = io.BytesIO()
    plt.savefig(buf4, format="png", dpi=120)
    plt.close(fig)
    print("[OK] KDE-распределение создано (в буфер)")
    print("\n[OK] Все 4 графика успешно построены")


# ────────────────────────────────────────────────
# 6. МАШИННОЕ ОБУЧЕНИЕ — RANDOM FOREST
# ────────────────────────────────────────────────

def train_random_forest(df: pd.DataFrame) -> None:
    """Обучает Random Forest для предсказания is_bestseller."""
    print("\n" + "=" * 60)
    print("  4. МОДЕЛЬ МАШИННОГО ОБУЧЕНИЯ: Random Forest")
    print("=" * 60)

    df_ml = df.copy()

    # Кодирование категориальных признаков
    le = LabelEncoder()
    df_ml["genre_enc"] = le.fit_transform(df_ml["genre"])

    feature_cols = ["price", "pages", "rating", "sales_count",
                    "author_rating", "year", "genre_enc"]
    target_col = "is_bestseller"

    X = df_ml[feature_cols].values
    y = df_ml[target_col].values

    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.25, random_state=42, stratify=y
    )

    print(f"\nРазбивка: {len(X_train)} обучающих / {len(X_test)} тестовых примеров")
    print(f"Распределение классов (train): {dict(zip(*np.unique(y_train, return_counts=True)))}")

    clf = RandomForestClassifier(
        n_estimators=100,
        max_depth=6,
        random_state=42,
        class_weight="balanced",
    )
    clf.fit(X_train, y_train)
    y_pred = clf.predict(X_test)

    accuracy = accuracy_score(y_test, y_pred)
    print(f"\n>>> Точность (accuracy): {accuracy:.4f} ({accuracy*100:.1f}%)")
    print("\n>>> Отчёт по классификации:")
    print(classification_report(y_test, y_pred, target_names=["Обычная", "Бестселлер"]))

    print(">>> Важность признаков (feature importances):")
    importances = clf.feature_importances_
    feat_imp = sorted(zip(feature_cols, importances), key=lambda x: x[1], reverse=True)
    for feat, imp in feat_imp:
        bar = "█" * int(imp * 40)
        print(f"    {feat:<20} {imp:.4f}  {bar}")


# ────────────────────────────────────────────────
# ТОЧКА ВХОДА
# ────────────────────────────────────────────────

if __name__ == "__main__":
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    CSV_PATH = os.path.join(BASE_DIR, "book_sales.csv")
    PLOTS_DIR = os.path.join(BASE_DIR, "plots_buffer")  # не реальные файлы

    print("=" * 60)
    print("  ЛР8 — Анализ датасета продаж книг (Pandas + sklearn)")
    print("=" * 60)

    # 1. Генерация и сохранение
    df_raw = generate_dataset(n_rows=220)
    df = save_and_load(df_raw, CSV_PATH)

    # 2. Базовый анализ
    basic_analysis(df)

    # 3. Предобработка
    df_clean = preprocess(df)

    # 4. Визуализация (без записи на диск)
    create_visualizations(df_clean, PLOTS_DIR)

    # 5. Машинное обучение
    train_random_forest(df_clean)

    print("\n" + "=" * 60)
    print("  [ЗАВЕРШЕНО] Все операции выполнены успешно.")
    print("=" * 60)
