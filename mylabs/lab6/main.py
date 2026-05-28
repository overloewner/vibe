"""
Лабораторная работа №6 — Основы веб-разработки
Предметная область: Книжный магазин (Bookstore)
Инструменты: FastAPI + Pydantic + Uvicorn
Дисциплина: Скриптовые языки программирования
"""

from fastapi import FastAPI, APIRouter, HTTPException, Query
from pydantic import BaseModel, Field
from typing import Optional, List
from datetime import date
import uvicorn

# ============================================================
#  PYDANTIC-МОДЕЛИ
# ============================================================

class CategoryBase(BaseModel):
    name: str = Field(..., min_length=1, max_length=100, description="Название категории")
    description: Optional[str] = Field(None, max_length=500)

class CategoryCreate(CategoryBase):
    pass

class CategoryOut(CategoryBase):
    id: int
    model_config = {"from_attributes": True}


class BookBase(BaseModel):
    title: str = Field(..., min_length=1, max_length=200)
    author: str = Field(..., min_length=1, max_length=100)
    genre: str = Field(..., min_length=1, max_length=60)
    year: int = Field(..., ge=1000, le=2100)
    price: float = Field(..., gt=0, description="Цена в рублях")
    stock_count: int = Field(default=0, ge=0)

class BookCreate(BookBase):
    pass

class BookUpdate(BaseModel):
    title: Optional[str] = Field(None, min_length=1, max_length=200)
    author: Optional[str] = Field(None, min_length=1, max_length=100)
    genre: Optional[str] = Field(None, min_length=1, max_length=60)
    year: Optional[int] = Field(None, ge=1000, le=2100)
    price: Optional[float] = Field(None, gt=0)
    stock_count: Optional[int] = Field(None, ge=0)

class BookOut(BookBase):
    id: int
    model_config = {"from_attributes": True}


class OrderBase(BaseModel):
    book_id: int = Field(..., gt=0)
    customer_name: str = Field(..., min_length=2, max_length=100)
    quantity: int = Field(..., ge=1, le=100)

class OrderCreate(OrderBase):
    pass

class OrderOut(OrderBase):
    id: int
    order_date: date
    model_config = {"from_attributes": True}


# ============================================================
#  IN-MEMORY ХРАНИЛИЩЕ
# ============================================================

books_db: dict[int, dict] = {}
categories_db: dict[int, dict] = {}
orders_db: dict[int, dict] = {}

_book_counter = 0
_cat_counter = 0
_order_counter = 0


def _next_book_id() -> int:
    global _book_counter
    _book_counter += 1
    return _book_counter

def _next_cat_id() -> int:
    global _cat_counter
    _cat_counter += 1
    return _cat_counter

def _next_order_id() -> int:
    global _order_counter
    _order_counter += 1
    return _order_counter


# ============================================================
#  РОУТЕРЫ
# ============================================================

books_router = APIRouter(prefix="/books", tags=["Книги"])
categories_router = APIRouter(prefix="/categories", tags=["Категории"])
orders_router = APIRouter(prefix="/orders", tags=["Заказы"])


# ---------- КНИГИ ----------

@books_router.get("/", response_model=List[BookOut], summary="Список всех книг")
def list_books(
    genre: Optional[str] = Query(None, description="Фильтр по жанру"),
    search: Optional[str] = Query(None, description="Поиск по названию (подстрока)"),
    min_price: Optional[float] = Query(None, gt=0),
    max_price: Optional[float] = Query(None, gt=0),
):
    result = list(books_db.values())
    if genre:
        result = [b for b in result if b["genre"].lower() == genre.lower()]
    if search:
        result = [b for b in result if search.lower() in b["title"].lower()]
    if min_price is not None:
        result = [b for b in result if b["price"] >= min_price]
    if max_price is not None:
        result = [b for b in result if b["price"] <= max_price]
    return result


@books_router.get("/{book_id}", response_model=BookOut, summary="Получить книгу по ID")
def get_book(book_id: int):
    if book_id not in books_db:
        raise HTTPException(status_code=404, detail=f"Книга с id={book_id} не найдена")
    return books_db[book_id]


@books_router.post("/", response_model=BookOut, status_code=201, summary="Добавить книгу")
def create_book(book: BookCreate):
    bid = _next_book_id()
    record = {"id": bid, **book.model_dump()}
    books_db[bid] = record
    return record


@books_router.put("/{book_id}", response_model=BookOut, summary="Обновить книгу")
def update_book(book_id: int, book: BookUpdate):
    if book_id not in books_db:
        raise HTTPException(status_code=404, detail=f"Книга с id={book_id} не найдена")
    existing = books_db[book_id]
    update_data = book.model_dump(exclude_unset=True)
    existing.update(update_data)
    books_db[book_id] = existing
    return existing


@books_router.delete("/{book_id}", status_code=204, summary="Удалить книгу")
def delete_book(book_id: int):
    if book_id not in books_db:
        raise HTTPException(status_code=404, detail=f"Книга с id={book_id} не найдена")
    del books_db[book_id]


# ---------- КАТЕГОРИИ ----------

@categories_router.get("/", response_model=List[CategoryOut], summary="Список категорий")
def list_categories():
    return list(categories_db.values())


@categories_router.get("/{cat_id}", response_model=CategoryOut, summary="Получить категорию по ID")
def get_category(cat_id: int):
    if cat_id not in categories_db:
        raise HTTPException(status_code=404, detail=f"Категория с id={cat_id} не найдена")
    return categories_db[cat_id]


@categories_router.post("/", response_model=CategoryOut, status_code=201, summary="Создать категорию")
def create_category(cat: CategoryCreate):
    cid = _next_cat_id()
    record = {"id": cid, **cat.model_dump()}
    categories_db[cid] = record
    return record


@categories_router.put("/{cat_id}", response_model=CategoryOut, summary="Обновить категорию")
def update_category(cat_id: int, cat: CategoryCreate):
    if cat_id not in categories_db:
        raise HTTPException(status_code=404, detail=f"Категория с id={cat_id} не найдена")
    record = {"id": cat_id, **cat.model_dump()}
    categories_db[cat_id] = record
    return record


@categories_router.delete("/{cat_id}", status_code=204, summary="Удалить категорию")
def delete_category(cat_id: int):
    if cat_id not in categories_db:
        raise HTTPException(status_code=404, detail=f"Категория с id={cat_id} не найдена")
    del categories_db[cat_id]


# ---------- ЗАКАЗЫ ----------

@orders_router.get("/", response_model=List[OrderOut], summary="Список заказов")
def list_orders(customer_name: Optional[str] = Query(None, description="Фильтр по имени покупателя")):
    result = list(orders_db.values())
    if customer_name:
        result = [o for o in result if customer_name.lower() in o["customer_name"].lower()]
    return result


@orders_router.get("/{order_id}", response_model=OrderOut, summary="Получить заказ по ID")
def get_order(order_id: int):
    if order_id not in orders_db:
        raise HTTPException(status_code=404, detail=f"Заказ с id={order_id} не найден")
    return orders_db[order_id]


@orders_router.post("/", response_model=OrderOut, status_code=201, summary="Создать заказ")
def create_order(order: OrderCreate):
    if order.book_id not in books_db:
        raise HTTPException(status_code=404, detail=f"Книга с id={order.book_id} не найдена")
    book = books_db[order.book_id]
    if book["stock_count"] < order.quantity:
        raise HTTPException(
            status_code=400,
            detail=f"Недостаточно товара на складе. В наличии: {book['stock_count']}"
        )
    # Уменьшаем остаток
    book["stock_count"] -= order.quantity

    oid = _next_order_id()
    record = {
        "id": oid,
        "book_id": order.book_id,
        "customer_name": order.customer_name,
        "quantity": order.quantity,
        "order_date": date.today(),
    }
    orders_db[oid] = record
    return record


@orders_router.delete("/{order_id}", status_code=204, summary="Отменить заказ")
def delete_order(order_id: int):
    if order_id not in orders_db:
        raise HTTPException(status_code=404, detail=f"Заказ с id={order_id} не найден")
    # Возвращаем товар на склад
    order = orders_db[order_id]
    if order["book_id"] in books_db:
        books_db[order["book_id"]]["stock_count"] += order["quantity"]
    del orders_db[order_id]


# ============================================================
#  ПРИЛОЖЕНИЕ
# ============================================================

app = FastAPI(
    title="Bookstore API",
    description="REST API для книжного магазина. Лабораторная работа №6.",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
)

app.include_router(books_router)
app.include_router(categories_router)
app.include_router(orders_router)


@app.get("/", tags=["Root"])
def root():
    return {
        "message": "Bookstore API",
        "docs": "/docs",
        "version": "1.0.0"
    }


# ============================================================
#  ТОЧКА ВХОДА
# ============================================================

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=False)
