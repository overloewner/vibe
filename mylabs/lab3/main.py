"""
Лабораторная работа №3 — Разработка через тестирование (TDD)
Тема: Тестирование класса Stack (Стек)
Дисциплина: Скриптовые языки программирования
"""
import unittest


# ============================================================
#  ТЕСТИРУЕМЫЙ МОДУЛЬ — класс Stack
# ============================================================

class Stack:
    """Реализация структуры данных «стек» (LIFO)."""

    def __init__(self):
        self._items = []

    def push(self, item) -> None:
        """Добавляет элемент на вершину стека."""
        self._items.append(item)

    def pop(self):
        """Удаляет и возвращает элемент с вершины стека."""
        if self.is_empty():
            raise IndexError("Стек пуст — операция pop невозможна")
        return self._items.pop()

    def peek(self):
        """Возвращает элемент с вершины стека без удаления."""
        if self.is_empty():
            raise IndexError("Стек пуст — операция peek невозможна")
        return self._items[-1]

    def is_empty(self) -> bool:
        """Возвращает True, если стек пустой."""
        return len(self._items) == 0

    def size(self) -> int:
        """Возвращает текущее количество элементов в стеке."""
        return len(self._items)

    def clear(self) -> None:
        """Удаляет все элементы из стека."""
        self._items = []

    def __repr__(self) -> str:
        return f"Stack({self._items})"


# ============================================================
#  МОДУЛЬНЫЕ ТЕСТЫ
# ============================================================

class TestStack(unittest.TestCase):

    def setUp(self):
        """Создаёт новый экземпляр стека перед каждым тестом."""
        self.stack = Stack()

    def tearDown(self):
        """Очищает стек после каждого теста."""
        self.stack.clear()

    # ТЕСТ 1: Добавление одного элемента
    def test_push_single_element(self):
        """Проверяет, что после push стек не пустой и размер равен 1."""
        self.stack.push(42)
        self.assertFalse(self.stack.is_empty())
        self.assertEqual(self.stack.size(), 1)

    # ТЕСТ 2: Добавление нескольких элементов
    def test_push_multiple_elements(self):
        """Проверяет корректный подсчёт размера при последовательных push."""
        for value in [10, 20, 30, 40]:
            self.stack.push(value)
        self.assertEqual(self.stack.size(), 4)

    # ТЕСТ 3: Порядок извлечения (LIFO)
    def test_pop_returns_last_pushed(self):
        """Проверяет принцип LIFO: последний добавленный — первый извлечённый."""
        self.stack.push("первый")
        self.stack.push("второй")
        self.stack.push("третий")
        self.assertEqual(self.stack.pop(), "третий")
        self.assertEqual(self.stack.pop(), "второй")
        self.assertEqual(self.stack.pop(), "первый")

    # ТЕСТ 4: Исключение при pop из пустого стека
    def test_pop_empty_stack_raises_index_error(self):
        """Проверяет, что pop на пустом стеке вызывает IndexError."""
        with self.assertRaises(IndexError):
            self.stack.pop()

    # ТЕСТ 5: Просмотр вершины без удаления
    def test_peek_returns_top_without_removing(self):
        """Проверяет, что peek возвращает вершину, не изменяя стек."""
        self.stack.push(100)
        self.stack.push(200)
        result = self.stack.peek()
        self.assertEqual(result, 200)
        self.assertEqual(self.stack.size(), 2)  # размер не изменился

    # ТЕСТ 6: Исключение при peek из пустого стека
    def test_peek_empty_stack_raises_index_error(self):
        """Проверяет, что peek на пустом стеке вызывает IndexError."""
        with self.assertRaises(IndexError):
            self.stack.peek()

    # ТЕСТ 7: Проверка is_empty для нового стека
    def test_is_empty_on_new_stack(self):
        """Новый стек должен быть пустым."""
        self.assertTrue(self.stack.is_empty())

    # ТЕСТ 8: Проверка is_empty после добавления и удаления
    def test_is_empty_after_push_and_pop(self):
        """Стек должен стать пустым после удаления всех элементов."""
        self.stack.push(1)
        self.stack.push(2)
        self.stack.pop()
        self.stack.pop()
        self.assertTrue(self.stack.is_empty())

    # ТЕСТ 9: Очистка стека методом clear
    def test_clear_empties_the_stack(self):
        """Метод clear должен полностью опустошить стек."""
        for i in range(5):
            self.stack.push(i)
        self.stack.clear()
        self.assertTrue(self.stack.is_empty())
        self.assertEqual(self.stack.size(), 0)

    # ТЕСТ 10: Разнотипные элементы
    def test_push_mixed_types(self):
        """Стек должен корректно хранить элементы разных типов."""
        self.stack.push(1)
        self.stack.push("текст")
        self.stack.push(3.14)
        self.stack.push([1, 2, 3])
        self.assertEqual(self.stack.size(), 4)
        self.assertEqual(self.stack.pop(), [1, 2, 3])
        self.assertAlmostEqual(self.stack.pop(), 3.14)


# ============================================================
#  ЗАПУСК ТЕСТОВ
# ============================================================

def run_tests():
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    suite.addTests(loader.loadTestsFromTestCase(TestStack))

    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    print("\n" + "=" * 60)
    print("ИТОГИ ТЕСТИРОВАНИЯ")
    print("=" * 60)
    print(f"  Всего тестов : {result.testsRun}")
    print(f"  Успешно      : {result.testsRun - len(result.failures) - len(result.errors)}")
    print(f"  Провалов     : {len(result.failures)}")
    print(f"  Ошибок       : {len(result.errors)}")
    print("=" * 60)
    return result.wasSuccessful()


if __name__ == "__main__":
    success = run_tests()
    exit(0 if success else 1)
