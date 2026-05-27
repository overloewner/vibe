"""
Лабораторная работа №4 — Совместная работа с Git и инспекцией кода
Дисциплина: Скриптовые языки программирования

Файл: main.py — модуль математических утилит,
который использовался как проект при изучении Git.
"""


def add(a: float, b: float) -> float:
    """Возвращает сумму двух чисел."""
    return a + b


def subtract(a: float, b: float) -> float:
    """Возвращает разность двух чисел."""
    return a - b


def multiply(a: float, b: float) -> float:
    """Возвращает произведение двух чисел."""
    return a * b


def divide(a: float, b: float) -> float:
    """Возвращает частное двух чисел. Вызывает ValueError при делении на ноль."""
    if b == 0:
        raise ValueError("Деление на ноль недопустимо")
    return a / b


def power(base: float, exp: int) -> float:
    """Возвращает основание, возведённое в степень."""
    return base ** exp


def factorial(n: int) -> int:
    """Вычисляет факториал неотрицательного целого числа."""
    if not isinstance(n, int) or n < 0:
        raise ValueError("Факториал определён только для неотрицательных целых чисел")
    result = 1
    for i in range(2, n + 1):
        result *= i
    return result


if __name__ == "__main__":
    print("=== Математические утилиты ===")
    print(f"add(3, 5)        = {add(3, 5)}")
    print(f"subtract(10, 4)  = {subtract(10, 4)}")
    print(f"multiply(6, 7)   = {multiply(6, 7)}")
    print(f"divide(15, 3)    = {divide(15, 3)}")
    print(f"power(2, 8)      = {power(2, 8)}")
    print(f"factorial(6)     = {factorial(6)}")
