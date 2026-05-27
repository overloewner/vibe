"""
Лабораторная работа №7 — Тестирование безопасности
Предметная область: Защищённый книжный магазин (Secured Bookstore)
Инструменты: FastAPI, JWT (HS256 HMAC — встроенный hmac+hashlib),
             bcrypt (passlib), Pydantic, CORS, Rate Limiting
Дисциплина: Скриптовые языки программирования

Примечание: JWT реализован вручную через hmac + hashlib (HS256),
поскольку python-jose требует библиотеку cryptography, которая
не всегда доступна в учебной среде.
"""

from fastapi import FastAPI, APIRouter, HTTPException, Depends, Request, status
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field, field_validator
from typing import Optional, List
from datetime import date, datetime, timedelta, timezone
import bcrypt as _bcrypt_lib
import re
import time
import hmac
import hashlib
import base64
import json
import uvicorn

# ============================================================
#  КОНФИГУРАЦИЯ БЕЗОПАСНОСТИ
# ============================================================

SECRET_KEY = b"super-secret-key-for-lab7-ngtu-2026"  # В prod хранить в .env!
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 30
REFRESH_TOKEN_EXPIRE_DAYS = 7

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/auth/token")


# ============================================================
#  ЧИСТАЯ РЕАЛИЗАЦИЯ JWT (HS256) БЕЗ ВНЕШНИХ ЗАВИСИМОСТЕЙ
# ============================================================

def _b64url_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode()


def _b64url_decode(s: str) -> bytes:
    padding = 4 - len(s) % 4
    if padding != 4:
        s += "=" * padding
    return base64.urlsafe_b64decode(s)


def jwt_encode(payload: dict, key: bytes) -> str:
    """Кодирует JWT токен с алгоритмом HS256."""
    header = _b64url_encode(json.dumps({"alg": "HS256", "typ": "JWT"}, separators=(",", ":")).encode())
    body = _b64url_encode(json.dumps(payload, separators=(",", ":"), default=str).encode())
    msg = f"{header}.{body}".encode()
    sig = hmac.new(key, msg, hashlib.sha256).digest()
    return f"{header}.{body}.{_b64url_encode(sig)}"


def jwt_decode(token: str, key: bytes) -> dict:
    """Декодирует и проверяет JWT токен."""
    try:
        parts = token.split(".")
        if len(parts) != 3:
            raise ValueError("Неверная структура токена")
        header_b, body_b, sig_b = parts
        msg = f"{header_b}.{body_b}".encode()
        expected_sig = hmac.new(key, msg, hashlib.sha256).digest()
        actual_sig = _b64url_decode(sig_b)
        if not hmac.compare_digest(expected_sig, actual_sig):
            raise ValueError("Неверная подпись токена")
        payload = json.loads(_b64url_decode(body_b))
        exp = payload.get("exp")
        if exp is not None and time.time() > exp:
            raise ValueError("Токен истёк")
        return payload
    except (ValueError, KeyError, json.JSONDecodeError) as e:
        raise HTTPException(status_code=401, detail=f"Недействительный токен: {e}")


# ============================================================
#  PYDANTIC-МОДЕЛИ С РАСШИРЕННОЙ ВАЛИДАЦИЕЙ
# ============================================================

class Token(BaseModel):
    access_token: str
    refresh_token: str
    token_type: str = "bearer"


class TokenData(BaseModel):
    username: Optional[str] = None
    role: Optional[str] = None


class UserCreate(BaseModel):
    username: str = Field(..., min_length=3, max_length=50,
                          description="Имя пользователя (3-50 символов, только буквы/цифры/._-)")
    password: str = Field(..., min_length=8, max_length=128,
                          description="Пароль (минимум 8 символов)")
    role: str = Field(default="user", pattern="^(admin|user)$")

    @field_validator("username")
    @classmethod
    def validate_username(cls, v: str) -> str:
        if not re.match(r"^[a-zA-Z0-9._\-]+$", v):
            raise ValueError("Имя пользователя может содержать только буквы, цифры, '.', '_', '-'")
        return v

    @field_validator("password")
    @classmethod
    def validate_password(cls, v: str) -> str:
        if not any(c.isupper() for c in v):
            raise ValueError("Пароль должен содержать хотя бы одну заглавную букву")
        if not any(c.isdigit() for c in v):
            raise ValueError("Пароль должен содержать хотя бы одну цифру")
        return v


class UserOut(BaseModel):
    username: str
    role: str


class CategoryBase(BaseModel):
    name: str = Field(..., min_length=1, max_length=100)
    description: Optional[str] = Field(None, max_length=500)


class CategoryCreate(CategoryBase):
    pass


class CategoryOut(CategoryBase):
    id: int
    model_config = {"from_attributes": True}


class BookBase(BaseModel):
    title: str = Field(..., min_length=1, max_length=200,
                       description="Название книги")
    author: str = Field(..., min_length=2, max_length=100,
                        description="Имя автора")
    genre: str = Field(..., min_length=1, max_length=60)
    year: int = Field(..., ge=1000, le=2100, description="Год издания")
    price: float = Field(..., gt=0, le=100000, description="Цена в рублях")
    stock_count: int = Field(default=0, ge=0, le=10000)

    @field_validator("title", "author")
    @classmethod
    def no_script_injection(cls, v: str) -> str:
        forbidden = ["<script>", "</script>", "javascript:", "onload=", "onerror="]
        for pattern in forbidden:
            if pattern.lower() in v.lower():
                raise ValueError(f"Недопустимый контент в поле: {pattern}")
        return v.strip()


class BookCreate(BookBase):
    pass


class BookUpdate(BaseModel):
    title: Optional[str] = Field(None, min_length=1, max_length=200)
    author: Optional[str] = Field(None, min_length=2, max_length=100)
    genre: Optional[str] = Field(None, min_length=1, max_length=60)
    year: Optional[int] = Field(None, ge=1000, le=2100)
    price: Optional[float] = Field(None, gt=0, le=100000)
    stock_count: Optional[int] = Field(None, ge=0, le=10000)


class BookOut(BookBase):
    id: int
    model_config = {"from_attributes": True}


class OrderBase(BaseModel):
    book_id: int = Field(..., gt=0)
    customer_name: str = Field(..., min_length=2, max_length=100,
                               description="ФИО покупателя")
    quantity: int = Field(..., ge=1, le=100)


class OrderCreate(OrderBase):
    pass


class OrderOut(OrderBase):
    id: int
    order_date: date
    model_config = {"from_attributes": True}


# ============================================================
#  ХРАНИЛИЩЕ ПОЛЬЗОВАТЕЛЕЙ
# ============================================================

users_db: dict[str, dict] = {}

# Администратор по умолчанию
_admin_hash = _bcrypt_lib.hashpw(b"Admin1234", _bcrypt_lib.gensalt())
users_db["admin"] = {
    "username": "admin",
    "hashed_password": _admin_hash,
    "role": "admin",
}

# Тестовый пользователь
_user_hash = _bcrypt_lib.hashpw(b"User1234!", _bcrypt_lib.gensalt())
users_db["testuser"] = {
    "username": "testuser",
    "hashed_password": _user_hash,
    "role": "user",
}


# ============================================================
#  IN-MEMORY ХРАНИЛИЩЕ ТОВАРОВ
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
#  RATE LIMITING (простой счётчик в памяти)
# ============================================================

_request_counts: dict[str, list] = {}
RATE_LIMIT = 60         # запросов
RATE_WINDOW = 60        # в секундах


def check_rate_limit(ip: str) -> bool:
    """Возвращает True если запрос разрешён, False если лимит превышен."""
    now = time.time()
    if ip not in _request_counts:
        _request_counts[ip] = []
    _request_counts[ip] = [t for t in _request_counts[ip] if now - t < RATE_WINDOW]
    if len(_request_counts[ip]) >= RATE_LIMIT:
        return False
    _request_counts[ip].append(now)
    return True


# ============================================================
#  JWT УТИЛИТЫ
# ============================================================

def verify_password(plain: str, hashed) -> bool:
    """Проверяет пароль с bcrypt хешем."""
    if isinstance(hashed, str):
        hashed = hashed.encode()
    return _bcrypt_lib.checkpw(plain.encode(), hashed)


def get_password_hash(password: str) -> bytes:
    """Хэширует пароль с bcrypt."""
    return _bcrypt_lib.hashpw(password.encode(), _bcrypt_lib.gensalt())


def create_access_token(data: dict, expires_delta: Optional[timedelta] = None) -> str:
    to_encode = data.copy()
    expire = time.time() + (expires_delta or timedelta(minutes=15)).total_seconds()
    to_encode.update({"exp": expire, "type": "access"})
    return jwt_encode(to_encode, SECRET_KEY)


def create_refresh_token(data: dict) -> str:
    to_encode = data.copy()
    expire = time.time() + timedelta(days=REFRESH_TOKEN_EXPIRE_DAYS).total_seconds()
    to_encode.update({"exp": expire, "type": "refresh"})
    return jwt_encode(to_encode, SECRET_KEY)


def decode_token(token: str) -> TokenData:
    payload = jwt_decode(token, SECRET_KEY)
    username = payload.get("sub")
    role = payload.get("role")
    if username is None:
        raise HTTPException(status_code=401, detail="Недействительный токен")
    return TokenData(username=username, role=role)


# ============================================================
#  ЗАВИСИМОСТИ FASTAPI
# ============================================================

async def get_current_user(token: str = Depends(oauth2_scheme)) -> dict:
    token_data = decode_token(token)
    user = users_db.get(token_data.username)
    if user is None:
        raise HTTPException(status_code=401, detail="Пользователь не найден")
    return user


async def require_admin(current_user: dict = Depends(get_current_user)) -> dict:
    if current_user["role"] != "admin":
        raise HTTPException(
            status_code=403,
            detail="Недостаточно прав. Требуется роль администратора."
        )
    return current_user


# ============================================================
#  MIDDLEWARE — Security Headers + Rate Limiting
# ============================================================

async def security_headers_middleware(request: Request, call_next):
    client_ip = request.client.host if request.client else "unknown"
    if not check_rate_limit(client_ip):
        return JSONResponse(
            status_code=429,
            content={"detail": "Слишком много запросов. Попробуйте позже."}
        )
    response = await call_next(request)
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-XSS-Protection"] = "1; mode=block"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'"
    )
    return response


# ============================================================
#  РОУТЕРЫ
# ============================================================

auth_router = APIRouter(prefix="/auth", tags=["Аутентификация"])
books_router = APIRouter(prefix="/books", tags=["Книги"])
categories_router = APIRouter(prefix="/categories", tags=["Категории"])
orders_router = APIRouter(prefix="/orders", tags=["Заказы"])
admin_router = APIRouter(prefix="/admin", tags=["Администрирование"])


# ---------- AUTH ----------

@auth_router.post("/register", response_model=UserOut, status_code=201,
                  summary="Регистрация пользователя")
def register(user_data: UserCreate):
    if user_data.username in users_db:
        raise HTTPException(status_code=400, detail="Пользователь уже существует")
    hashed = _bcrypt_lib.hashpw(user_data.password.encode(), _bcrypt_lib.gensalt())
    users_db[user_data.username] = {
        "username": user_data.username,
        "hashed_password": hashed,
        "role": user_data.role,
    }
    return UserOut(username=user_data.username, role=user_data.role)


@auth_router.post("/token", response_model=Token, summary="Получение JWT токенов")
def login(form_data: OAuth2PasswordRequestForm = Depends()):
    user = users_db.get(form_data.username)
    if not user or not verify_password(form_data.password, user["hashed_password"]):
        raise HTTPException(
            status_code=401,
            detail="Неверное имя пользователя или пароль",
            headers={"WWW-Authenticate": "Bearer"},
        )
    token_data = {"sub": user["username"], "role": user["role"]}
    access_token = create_access_token(token_data, timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES))
    refresh_token = create_refresh_token(token_data)
    return Token(access_token=access_token, refresh_token=refresh_token)


@auth_router.post("/refresh", response_model=Token, summary="Обновление токена доступа")
def refresh_token_endpoint(refresh: str):
    payload = jwt_decode(refresh, SECRET_KEY)
    if payload.get("type") != "refresh":
        raise HTTPException(status_code=401, detail="Ожидается refresh токен")
    username = payload.get("sub")
    role = payload.get("role")
    user = users_db.get(username)
    if not user:
        raise HTTPException(status_code=401, detail="Пользователь не найден")
    token_data = {"sub": username, "role": role}
    new_access = create_access_token(token_data, timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES))
    new_refresh = create_refresh_token(token_data)
    return Token(access_token=new_access, refresh_token=new_refresh)


@auth_router.get("/me", response_model=UserOut, summary="Текущий пользователь")
def get_me(current_user: dict = Depends(get_current_user)):
    return UserOut(username=current_user["username"], role=current_user["role"])


# ---------- КНИГИ ----------

@books_router.get("/", response_model=List[BookOut], summary="Список всех книг")
def list_books(
    genre: Optional[str] = None,
    search: Optional[str] = None,
    _user: dict = Depends(get_current_user),
):
    result = list(books_db.values())
    if genre:
        result = [b for b in result if b["genre"].lower() == genre.lower()]
    if search:
        result = [b for b in result if search.lower() in b["title"].lower()]
    return result


@books_router.get("/{book_id}", response_model=BookOut, summary="Получить книгу по ID")
def get_book(book_id: int, _user: dict = Depends(get_current_user)):
    if book_id not in books_db:
        raise HTTPException(status_code=404, detail=f"Книга с id={book_id} не найдена")
    return books_db[book_id]


@books_router.post("/", response_model=BookOut, status_code=201,
                   summary="Добавить книгу (только admin)")
def create_book(book: BookCreate, _admin: dict = Depends(require_admin)):
    bid = _next_book_id()
    record = {"id": bid, **book.model_dump()}
    books_db[bid] = record
    return record


@books_router.put("/{book_id}", response_model=BookOut,
                  summary="Обновить книгу (только admin)")
def update_book(book_id: int, book: BookUpdate, _admin: dict = Depends(require_admin)):
    if book_id not in books_db:
        raise HTTPException(status_code=404, detail=f"Книга с id={book_id} не найдена")
    existing = books_db[book_id]
    existing.update(book.model_dump(exclude_unset=True))
    books_db[book_id] = existing
    return existing


@books_router.delete("/{book_id}", status_code=204,
                     summary="Удалить книгу (только admin)")
def delete_book(book_id: int, _admin: dict = Depends(require_admin)):
    if book_id not in books_db:
        raise HTTPException(status_code=404, detail=f"Книга с id={book_id} не найдена")
    del books_db[book_id]


# ---------- КАТЕГОРИИ ----------

@categories_router.get("/", response_model=List[CategoryOut], summary="Список категорий")
def list_categories(_user: dict = Depends(get_current_user)):
    return list(categories_db.values())


@categories_router.get("/{cat_id}", response_model=CategoryOut)
def get_category(cat_id: int, _user: dict = Depends(get_current_user)):
    if cat_id not in categories_db:
        raise HTTPException(status_code=404, detail="Категория не найдена")
    return categories_db[cat_id]


@categories_router.post("/", response_model=CategoryOut, status_code=201,
                         summary="Создать категорию (только admin)")
def create_category(cat: CategoryCreate, _admin: dict = Depends(require_admin)):
    cid = _next_cat_id()
    record = {"id": cid, **cat.model_dump()}
    categories_db[cid] = record
    return record


@categories_router.put("/{cat_id}", response_model=CategoryOut,
                        summary="Обновить категорию (только admin)")
def update_category(cat_id: int, cat: CategoryCreate, _admin: dict = Depends(require_admin)):
    if cat_id not in categories_db:
        raise HTTPException(status_code=404, detail="Категория не найдена")
    record = {"id": cat_id, **cat.model_dump()}
    categories_db[cat_id] = record
    return record


@categories_router.delete("/{cat_id}", status_code=204,
                           summary="Удалить категорию (только admin)")
def delete_category(cat_id: int, _admin: dict = Depends(require_admin)):
    if cat_id not in categories_db:
        raise HTTPException(status_code=404, detail="Категория не найдена")
    del categories_db[cat_id]


# ---------- ЗАКАЗЫ ----------

@orders_router.get("/", response_model=List[OrderOut],
                   summary="Список заказов (только admin)")
def list_orders(_admin: dict = Depends(require_admin)):
    return list(orders_db.values())


@orders_router.get("/{order_id}", response_model=OrderOut, summary="Получить заказ по ID")
def get_order(order_id: int, _user: dict = Depends(get_current_user)):
    if order_id not in orders_db:
        raise HTTPException(status_code=404, detail="Заказ не найден")
    return orders_db[order_id]


@orders_router.post("/", response_model=OrderOut, status_code=201,
                    summary="Создать заказ (авторизованный пользователь)")
def create_order(order: OrderCreate, _user: dict = Depends(get_current_user)):
    if order.book_id not in books_db:
        raise HTTPException(status_code=404, detail=f"Книга с id={order.book_id} не найдена")
    book = books_db[order.book_id]
    if book["stock_count"] < order.quantity:
        raise HTTPException(
            status_code=400,
            detail=f"Недостаточно товара. В наличии: {book['stock_count']}"
        )
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


@orders_router.delete("/{order_id}", status_code=204,
                       summary="Отменить заказ (только admin)")
def delete_order(order_id: int, _admin: dict = Depends(require_admin)):
    if order_id not in orders_db:
        raise HTTPException(status_code=404, detail="Заказ не найден")
    order = orders_db[order_id]
    if order["book_id"] in books_db:
        books_db[order["book_id"]]["stock_count"] += order["quantity"]
    del orders_db[order_id]


# ---------- ADMIN ----------

@admin_router.get("/users", response_model=List[UserOut],
                  summary="Список всех пользователей")
def list_users(_admin: dict = Depends(require_admin)):
    return [UserOut(username=u["username"], role=u["role"]) for u in users_db.values()]


@admin_router.delete("/users/{username}", status_code=204,
                     summary="Удалить пользователя")
def delete_user(username: str, _admin: dict = Depends(require_admin)):
    if username == "admin":
        raise HTTPException(status_code=400, detail="Нельзя удалить основного администратора")
    if username not in users_db:
        raise HTTPException(status_code=404, detail="Пользователь не найден")
    del users_db[username]


# ============================================================
#  ПРИЛОЖЕНИЕ
# ============================================================

app = FastAPI(
    title="Secured Bookstore API",
    description=(
        "Защищённый REST API для книжного магазина. "
        "Лабораторная работа №7 — Тестирование безопасности."
    ),
    version="2.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
)

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000", "http://localhost:8080"],
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE"],
    allow_headers=["Authorization", "Content-Type"],
)

# Security Headers + Rate Limiting middleware
app.middleware("http")(security_headers_middleware)

# Роутеры
app.include_router(auth_router)
app.include_router(books_router)
app.include_router(categories_router)
app.include_router(orders_router)
app.include_router(admin_router)


@app.get("/", tags=["Root"])
def root():
    return {
        "message": "Secured Bookstore API v2.0",
        "docs": "/docs",
        "features": [
            "JWT Auth (HS256)",
            "RBAC (admin/user)",
            "bcrypt password hashing",
            "Rate Limiting (60 req/min)",
            "Security Headers",
            "CORS",
            "Pydantic input validation",
        ]
    }


# ============================================================
#  ТОЧКА ВХОДА
# ============================================================

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8001, reload=False)
