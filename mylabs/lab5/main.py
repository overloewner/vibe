"""
Лабораторная работа №5 — Работа с базами данных
Предметная область: Система управления библиотекой
Инструменты: SQLAlchemy ORM + SQLite (аналог PostgreSQL)
Дисциплина: Скриптовые языки программирования
"""

from sqlalchemy import (
    create_engine, Column, Integer, String, Text,
    ForeignKey, Date, Table, UniqueConstraint
)
from sqlalchemy.orm import DeclarativeBase, relationship, Session, joinedload, selectinload
from datetime import date
import os

# ────────────────────────────────────────────────
# ПОДКЛЮЧЕНИЕ К БАЗЕ ДАННЫХ
# ────────────────────────────────────────────────

DATABASE_URL = "sqlite:///library.db"
engine = create_engine(DATABASE_URL, echo=False)


# ────────────────────────────────────────────────
# БАЗОВЫЙ КЛАСС МОДЕЛЕЙ
# ────────────────────────────────────────────────

class Base(DeclarativeBase):
    pass


# ────────────────────────────────────────────────
# АССОЦИАТИВНАЯ ТАБЛИЦА (M:N) — читатель ↔ книга
# ────────────────────────────────────────────────

borrowings = Table(
    "borrowings",
    Base.metadata,
    Column("id", Integer, primary_key=True),
    Column("reader_id", Integer, ForeignKey("readers.id"), nullable=False),
    Column("book_id", Integer, ForeignKey("books.id"), nullable=False),
    Column("borrow_date", Date, nullable=False),
    Column("return_date", Date),
)


# ────────────────────────────────────────────────
# МОДЕЛИ
# ────────────────────────────────────────────────

class Author(Base):
    """Автор книги. Связь 1:1 с AuthorProfile, 1:N с Book."""
    __tablename__ = "authors"

    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False)
    birth_year = Column(Integer)

    profile = relationship(
        "AuthorProfile",
        back_populates="author",
        uselist=False,
        cascade="all, delete-orphan",
    )
    books = relationship(
        "Book",
        back_populates="author",
        cascade="all, delete-orphan",
    )

    def __repr__(self):
        return f"<Author id={self.id} name='{self.name}'>"


class AuthorProfile(Base):
    """Расширенный профиль автора (1:1 с Author)."""
    __tablename__ = "author_profiles"

    id = Column(Integer, primary_key=True)
    author_id = Column(Integer, ForeignKey("authors.id"), unique=True, nullable=False)
    nationality = Column(String(60))
    biography = Column(Text)

    author = relationship("Author", back_populates="profile")

    def __repr__(self):
        return f"<AuthorProfile author_id={self.author_id} nationality='{self.nationality}'>"


class Book(Base):
    """Книга. Связь N:1 с Author, N:M с Reader через borrowings."""
    __tablename__ = "books"

    id = Column(Integer, primary_key=True)
    title = Column(String(200), nullable=False)
    year = Column(Integer)
    pages = Column(Integer)
    genre = Column(String(60))
    author_id = Column(Integer, ForeignKey("authors.id"), nullable=False)

    author = relationship("Author", back_populates="books")
    readers = relationship("Reader", secondary=borrowings, back_populates="books")

    def __repr__(self):
        return f"<Book id={self.id} title='{self.title}' year={self.year}>"


class Reader(Base):
    """Читатель. Связь N:M с Book через ассоциативную таблицу borrowings."""
    __tablename__ = "readers"

    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False)
    email = Column(String(100), unique=True)
    phone = Column(String(20))

    books = relationship("Book", secondary=borrowings, back_populates="readers")

    def __repr__(self):
        return f"<Reader id={self.id} name='{self.name}'>"


# ────────────────────────────────────────────────
# СОЗДАНИЕ ТАБЛИЦ
# ────────────────────────────────────────────────

def create_tables():
    Base.metadata.create_all(engine)
    print("[OK] Таблицы созданы: authors, author_profiles, books, readers, borrowings")


# ────────────────────────────────────────────────
# НАПОЛНЕНИЕ ТЕСТОВЫМИ ДАННЫМИ (CREATE)
# ────────────────────────────────────────────────

def populate_data():
    with Session(engine) as session:
        # --- Авторы + профили (1:1) ---
        pushkin = Author(name="Александр Пушкин", birth_year=1799)
        pushkin.profile = AuthorProfile(
            nationality="Российская",
            biography="Великий русский поэт и прозаик, основоположник современного литературного языка."
        )

        tolstoy = Author(name="Лев Толстой", birth_year=1828)
        tolstoy.profile = AuthorProfile(
            nationality="Российская",
            biography="Русский писатель, один из величайших романистов мировой литературы."
        )

        bulgakov = Author(name="Михаил Булгаков", birth_year=1891)
        bulgakov.profile = AuthorProfile(
            nationality="Российская",
            biography="Русский писатель, драматург и театральный режиссёр."
        )

        session.add_all([pushkin, tolstoy, bulgakov])
        session.flush()

        # --- Книги (1:N с авторами) ---
        evgeniy_onegin = Book(title="Евгений Онегин", year=1833, pages=224, genre="Роман в стихах", author_id=pushkin.id)
        captain_daughter = Book(title="Капитанская дочка", year=1836, pages=192, genre="Роман", author_id=pushkin.id)
        war_and_peace = Book(title="Война и мир", year=1869, pages=1274, genre="Роман-эпопея", author_id=tolstoy.id)
        anna_karenina = Book(title="Анна Каренина", year=1878, pages=864, genre="Роман", author_id=tolstoy.id)
        master_margarita = Book(title="Мастер и Маргарита", year=1967, pages=480, genre="Роман", author_id=bulgakov.id)

        session.add_all([evgeniy_onegin, captain_daughter, war_and_peace, anna_karenina, master_margarita])
        session.flush()

        # --- Читатели ---
        reader1 = Reader(name="Иванова Анна", email="ivanova@mail.ru", phone="+7-900-111-22-33")
        reader2 = Reader(name="Петров Сергей", email="petrov@mail.ru", phone="+7-900-222-33-44")
        reader3 = Reader(name="Сидорова Мария", email="sidorova@mail.ru", phone="+7-900-333-44-55")

        session.add_all([reader1, reader2, reader3])
        session.flush()

        # --- Выдача книг (M:N через borrowings) ---
        session.execute(borrowings.insert(), [
            {"reader_id": reader1.id, "book_id": evgeniy_onegin.id,
             "borrow_date": date(2026, 1, 10), "return_date": date(2026, 1, 25)},
            {"reader_id": reader1.id, "book_id": master_margarita.id,
             "borrow_date": date(2026, 2, 1), "return_date": None},
            {"reader_id": reader2.id, "book_id": war_and_peace.id,
             "borrow_date": date(2026, 1, 15), "return_date": date(2026, 2, 10)},
            {"reader_id": reader2.id, "book_id": anna_karenina.id,
             "borrow_date": date(2026, 2, 5), "return_date": None},
            {"reader_id": reader3.id, "book_id": captain_daughter.id,
             "borrow_date": date(2026, 1, 20), "return_date": date(2026, 1, 30)},
            {"reader_id": reader3.id, "book_id": master_margarita.id,
             "borrow_date": date(2026, 2, 3), "return_date": None},
        ])

        session.commit()
        print("[OK] Данные добавлены: 3 автора, 5 книг, 3 читателя, 6 записей о выдаче")


# ────────────────────────────────────────────────
# ЧТЕНИЕ (READ) — запросы на выборку
# ────────────────────────────────────────────────

def read_data():
    with Session(engine) as session:
        print("\n─── 1. Все авторы с профилями (связь 1:1) ─────────────────────")
        authors = session.query(Author).options(joinedload(Author.profile)).all()
        for a in authors:
            nat = a.profile.nationality if a.profile else "—"
            print(f"  {a.name} ({a.birth_year}), национальность: {nat}")

        print("\n─── 2. Книги каждого автора (связь 1:N) ────────────────────────")
        authors = session.query(Author).options(selectinload(Author.books)).all()
        for a in authors:
            titles = ", ".join(b.title for b in a.books)
            print(f"  {a.name}: {titles}")

        print("\n─── 3. Читатели и взятые книги (связь M:N) ─────────────────────")
        readers = session.query(Reader).options(selectinload(Reader.books)).all()
        for r in readers:
            titles = ", ".join(b.title for b in r.books) or "—"
            print(f"  {r.name}: {titles}")

        print("\n─── 4. Фильтрация: книги жанра 'Роман' ─────────────────────────")
        novels = session.query(Book).filter(Book.genre == "Роман").all()
        for b in novels:
            print(f"  «{b.title}» ({b.year}), автор: {b.author.name}")

        print("\n─── 5. Книги, взятые и не возвращённые (return_date IS NULL) ───")
        active = session.execute(
            borrowings.select().where(borrowings.c.return_date == None)
        ).fetchall()
        for row in active:
            book = session.get(Book, row.book_id)
            reader = session.get(Reader, row.reader_id)
            print(f"  «{book.title}» — у читателя {reader.name} с {row.borrow_date}")


# ────────────────────────────────────────────────
# ОБНОВЛЕНИЕ (UPDATE)
# ────────────────────────────────────────────────

def update_data():
    with Session(engine) as session:
        print("\n─── Обновление: добавление страниц книге «Евгений Онегин» ─────")
        book = session.query(Book).filter_by(title="Евгений Онегин").first()
        print(f"  До: pages = {book.pages}")
        book.pages = 256
        session.commit()
        session.refresh(book)
        print(f"  После: pages = {book.pages}")

        print("\n─── Обновление: изменение email читателя Иванова Анна ──────────")
        reader = session.query(Reader).filter_by(name="Иванова Анна").first()
        print(f"  До: email = {reader.email}")
        reader.email = "anna.ivanova@library.ru"
        session.commit()
        session.refresh(reader)
        print(f"  После: email = {reader.email}")


# ────────────────────────────────────────────────
# УДАЛЕНИЕ (DELETE)
# ────────────────────────────────────────────────

def delete_data():
    with Session(engine) as session:
        print("\n─── Удаление: книга «Анна Каренина» ────────────────────────────")
        book = session.query(Book).filter_by(title="Анна Каренина").first()
        if book:
            session.delete(book)
            session.commit()
            print("  Книга «Анна Каренина» удалена")

        remaining = session.query(Book).all()
        print("  Оставшиеся книги:", [b.title for b in remaining])


# ────────────────────────────────────────────────
# УПРАВЛЕНИЕ ТРАНЗАКЦИЕЙ (демонстрация rollback)
# ────────────────────────────────────────────────

def transaction_demo():
    print("\n─── Демонстрация управления транзакцией (rollback) ─────────────")
    with Session(engine) as session:
        try:
            reader = Reader(name="Тестовый Пользователь", email="test@test.com")
            session.add(reader)
            # Намеренно вызываем повторный email для демонстрации rollback
            duplicate = Reader(name="Дубликат", email="test@test.com")
            session.add(duplicate)
            session.commit()
        except Exception as e:
            session.rollback()
            print(f"  Транзакция отменена (rollback): {type(e).__name__}")
            print("  Данные базы данных не изменены")


# ────────────────────────────────────────────────
# ТОЧКА ВХОДА
# ────────────────────────────────────────────────

if __name__ == "__main__":
    # Удаляем старую БД если есть
    if os.path.exists("library.db"):
        os.remove("library.db")

    print("=" * 60)
    print("  ЛР5 — Система управления библиотекой (SQLAlchemy ORM)")
    print("=" * 60)

    create_tables()
    populate_data()
    read_data()
    update_data()
    delete_data()
    transaction_demo()

    print("\n[ЗАВЕРШЕНО] Все операции выполнены успешно.")
