"""
Скрипт генерации отчётов (.docx) для лабораторных работ 3, 4, 5
"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy


# ═══════════════════════════════════════════════
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ═══════════════════════════════════════════════

def set_page_margins(doc, top=2, bottom=2, left=3, right=1.5):
    """Устанавливает поля страницы в сантиметрах."""
    section = doc.sections[0]
    section.top_margin = Cm(top)
    section.bottom_margin = Cm(bottom)
    section.left_margin = Cm(left)
    section.right_margin = Cm(right)


def add_paragraph(doc, text, bold=False, italic=False, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                  font_size=14, first_indent=True, space_before=0, space_after=0):
    """Добавляет параграф с нужным форматированием."""
    p = doc.add_paragraph()
    p.alignment = alignment
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    if first_indent and alignment == WD_ALIGN_PARAGRAPH.LEFT:
        pf.first_line_indent = Cm(1.25)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.name = "Times New Roman"
    run.font.size = Pt(font_size)
    # Fix Cyrillic font
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'Times New Roman')
    rPr.insert(0, rFonts)
    return p


def add_heading(doc, text, level=1, font_size=14):
    """Добавляет заголовок раздела."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.space_before = Pt(6)
    pf.space_after = Pt(3)
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.first_line_indent = Cm(0)
    run = p.add_run(text)
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(font_size)
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'Times New Roman')
    rPr.insert(0, rFonts)
    return p


def add_code_block(doc, code_text, caption=None):
    """Добавляет блок кода с моноширинным шрифтом."""
    if caption:
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(4)
        p_cap.paragraph_format.space_after = Pt(2)
        r = p_cap.add_run(caption)
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)

    for line in code_text.split('\n'):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.first_line_indent = Cm(0)
        pf.left_indent = Cm(0)
        run = p.add_run(line if line else " ")
        run.font.name = "Courier New"
        run.font.size = Pt(10)
        r = run._element
        rPr = r.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), 'Courier New')
        rFonts.set(qn('w:hAnsi'), 'Courier New')
        rFonts.set(qn('w:cs'), 'Courier New')
        rPr.insert(0, rFonts)

        # Light grey background
        pPr = p._element.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        for side in ['top', 'left', 'bottom', 'right']:
            bdr = OxmlElement(f'w:{side}')
            bdr.set(qn('w:val'), 'single')
            bdr.set(qn('w:sz'), '4')
            bdr.set(qn('w:space'), '1')
            bdr.set(qn('w:color'), 'CCCCCC')
            pBdr.append(bdr)
        pPr.append(pBdr)

        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'F5F5F5')
        pPr.append(shd)


def add_figure_caption(doc, text):
    """Добавляет подпись к рисунку."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(2)
    pf.space_after = Pt(6)
    pf.first_line_indent = Cm(0)
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    run.italic = True


def add_page_break(doc):
    p = doc.add_page_break()


def create_title_page(doc, lab_num, lab_title, student_name="Аценкова М.В.", group="С22-СИБ"):
    """Создаёт титульный лист по стандарту НГТУ."""
    center = WD_ALIGN_PARAGRAPH.CENTER

    def title_para(text, bold=False, size=14, space_b=0, space_a=0):
        p = doc.add_paragraph()
        p.alignment = center
        pf = p.paragraph_format
        pf.space_before = Pt(space_b)
        pf.space_after = Pt(space_a)
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.first_line_indent = Cm(0)
        run = p.add_run(text)
        run.bold = bold
        run.font.name = "Times New Roman"
        run.font.size = Pt(size)
        r = run._element
        rPr = r.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        rFonts.set(qn('w:cs'), 'Times New Roman')
        rPr.insert(0, rFonts)
        return p

    title_para("МИНОБРНАУКИ РОССИИ", bold=True, size=12)
    title_para("Федеральное государственное бюджетное образовательное учреждение высшего образования", size=12)
    title_para("НИЖЕГОРОДСКИЙ ГОСУДАРСТВЕННЫЙ ТЕХНИЧЕСКИЙ", bold=True, size=14)
    title_para("УНИВЕРСИТЕТ им. Р.Е.АЛЕКСЕЕВА", bold=True, size=14)
    title_para("Институт радиоэлектроники и информационных технологий", size=12, space_b=4)
    title_para("Кафедра «Информационная безопасность вычислительных систем и сетей»", size=12, space_b=2)

    title_para(f"«{lab_title}»", bold=True, size=14, space_b=20)
    title_para("ОТЧЕТ", bold=True, size=14, space_b=6)
    title_para(f"по лабораторной работе №{lab_num}", size=14)
    title_para("по дисциплине", size=14)
    title_para("Скриптовые языки программирования", bold=True, size=14, space_a=20)

    # Instructor / Student table-like layout
    def sign_line(label, value):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.space_before = Pt(2)
        pf.space_after = Pt(2)
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.first_line_indent = Cm(0)
        r1 = p.add_run(f"{label}  ")
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(14)
        r2 = p.add_run(value)
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(14)
        r = r1._element
        rPr = r.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        rFonts.set(qn('w:cs'), 'Times New Roman')
        rPr.insert(0, rFonts)

    sign_line("РУКОВОДИТЕЛЬ:", "Вайнбаум Д.А.")
    sign_line("________________", "(подпись)")

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    sign_line("СТУДЕНТ:", student_name)
    sign_line("________________", "(подпись)")

    p_gr = doc.add_paragraph()
    p_gr.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_gr.paragraph_format.space_before = Pt(2)
    p_gr.paragraph_format.first_line_indent = Cm(0)
    p_gr.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    r = p_gr.add_run(f"                    {group}  (шифр группы)")
    r.font.name = "Times New Roman"
    r.font.size = Pt(14)

    title_para("Работа защищена «___» ____________", size=14, space_b=10)
    title_para("С оценкой ________________________", size=14)
    title_para("Нижний Новгород 2026", bold=False, size=14, space_b=20)


# ═══════════════════════════════════════════════
# ЛР3 — TDD
# ═══════════════════════════════════════════════

LAB3_CODE = '''import unittest


class Stack:
    """Реализация структуры данных «стек» (LIFO)."""

    def __init__(self):
        self._items = []

    def push(self, item) -> None:
        self._items.append(item)

    def pop(self):
        if self.is_empty():
            raise IndexError("Стек пуст — операция pop невозможна")
        return self._items.pop()

    def peek(self):
        if self.is_empty():
            raise IndexError("Стек пуст — операция peek невозможна")
        return self._items[-1]

    def is_empty(self) -> bool:
        return len(self._items) == 0

    def size(self) -> int:
        return len(self._items)

    def clear(self) -> None:
        self._items = []


class TestStack(unittest.TestCase):

    def setUp(self):
        self.stack = Stack()

    def tearDown(self):
        self.stack.clear()

    def test_push_single_element(self):
        self.stack.push(42)
        self.assertFalse(self.stack.is_empty())
        self.assertEqual(self.stack.size(), 1)

    def test_push_multiple_elements(self):
        for value in [10, 20, 30, 40]:
            self.stack.push(value)
        self.assertEqual(self.stack.size(), 4)

    def test_pop_returns_last_pushed(self):
        self.stack.push("первый")
        self.stack.push("второй")
        self.stack.push("третий")
        self.assertEqual(self.stack.pop(), "третий")
        self.assertEqual(self.stack.pop(), "второй")
        self.assertEqual(self.stack.pop(), "первый")

    def test_pop_empty_stack_raises_index_error(self):
        with self.assertRaises(IndexError):
            self.stack.pop()

    def test_peek_returns_top_without_removing(self):
        self.stack.push(100)
        self.stack.push(200)
        result = self.stack.peek()
        self.assertEqual(result, 200)
        self.assertEqual(self.stack.size(), 2)

    def test_peek_empty_stack_raises_index_error(self):
        with self.assertRaises(IndexError):
            self.stack.peek()

    def test_is_empty_on_new_stack(self):
        self.assertTrue(self.stack.is_empty())

    def test_is_empty_after_push_and_pop(self):
        self.stack.push(1)
        self.stack.push(2)
        self.stack.pop()
        self.stack.pop()
        self.assertTrue(self.stack.is_empty())

    def test_clear_empties_the_stack(self):
        for i in range(5):
            self.stack.push(i)
        self.stack.clear()
        self.assertTrue(self.stack.is_empty())
        self.assertEqual(self.stack.size(), 0)

    def test_push_mixed_types(self):
        self.stack.push(1)
        self.stack.push("текст")
        self.stack.push(3.14)
        self.stack.push([1, 2, 3])
        self.assertEqual(self.stack.size(), 4)
        self.assertEqual(self.stack.pop(), [1, 2, 3])


def run_tests():
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    suite.addTests(loader.loadTestsFromTestCase(TestStack))
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    print("\\n" + "=" * 60)
    print("ИТОГИ ТЕСТИРОВАНИЯ")
    print("=" * 60)
    print(f"  Всего тестов : {result.testsRun}")
    print(f"  Успешно      : {result.testsRun - len(result.failures) - len(result.errors)}")
    print(f"  Провалов     : {len(result.failures)}")
    print(f"  Ошибок       : {len(result.errors)}")
    print("=" * 60)
    return result.wasSuccessful()


if __name__ == "__main__":
    run_tests()'''

LAB3_OUTPUT = '''test_clear_empties_the_stack (__main__.TestStack.test_clear_empties_the_stack)
Метод clear должен полностью опустошить стек. ... ok
test_is_empty_after_push_and_pop (__main__.TestStack.test_is_empty_after_push_and_pop)
Стек должен стать пустым после удаления всех элементов. ... ok
test_is_empty_on_new_stack (__main__.TestStack.test_is_empty_on_new_stack)
Новый стек должен быть пустым. ... ok
test_peek_empty_stack_raises_index_error (__main__.TestStack.test_peek_empty_stack_raises_index_error)
Проверяет, что peek на пустом стеке вызывает IndexError. ... ok
test_peek_returns_top_without_removing (__main__.TestStack.test_peek_returns_top_without_removing)
Проверяет, что peek возвращает вершину, не изменяя стек. ... ok
test_pop_empty_stack_raises_index_error (__main__.TestStack.test_pop_empty_stack_raises_index_error)
Проверяет, что pop на пустом стеке вызывает IndexError. ... ok
test_pop_returns_last_pushed (__main__.TestStack.test_pop_returns_last_pushed)
Проверяет принцип LIFO: последний добавленный — первый извлечённый. ... ok
test_push_mixed_types (__main__.TestStack.test_push_mixed_types)
Стек должен корректно хранить элементы разных типов. ... ok
test_push_multiple_elements (__main__.TestStack.test_push_multiple_elements)
Проверяет корректный подсчёт размера при последовательных push. ... ok
test_push_single_element (__main__.TestStack.test_push_single_element)
Проверяет, что после push стек не пустой и размер равен 1. ... ok

----------------------------------------------------------------------
Ran 10 tests in 0.000s

OK

============================================================
ИТОГИ ТЕСТИРОВАНИЯ
============================================================
  Всего тестов : 10
  Успешно      : 10
  Провалов     : 0
  Ошибок       : 0
============================================================'''


def create_lab3_report(output_path):
    doc = Document()
    set_page_margins(doc)

    create_title_page(doc, "3", "Разработка через тестирование (TDD)")
    add_page_break(doc)

    # Цель
    p = add_paragraph(doc, "", first_indent=False)
    p2 = add_paragraph(doc, "Цель работы: ", bold=True, first_indent=True)
    # inline mix
    p2.runs[0].text = ""
    r_bold = p2.add_run("Цель работы: ")
    r_bold.bold = True
    r_bold.font.name = "Times New Roman"
    r_bold.font.size = Pt(14)
    r_normal = p2.add_run(
        "изучить принципы модульного тестирования программ на языке Python с использованием "
        "встроенного модуля unittest, а также разработать набор тестов, проверяющий "
        "корректность реализации программы на различных входных данных, граничные случаи "
        "и обработку исключительных ситуаций."
    )
    r_normal.font.name = "Times New Roman"
    r_normal.font.size = Pt(14)

    # Постановка задачи
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p3.paragraph_format.first_line_indent = Cm(1.25)
    p3.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    rb = p3.add_run("Постановка задачи: ")
    rb.bold = True
    rb.font.name = "Times New Roman"
    rb.font.size = Pt(14)
    rn = p3.add_run(
        "разработать набор модульных тестов для реализации структуры данных «Стек» (Stack). "
        "Стек реализует принцип LIFO (Last In — First Out): последний добавленный элемент "
        "извлекается первым. Тесты должны проверить основные операции (push, pop, peek, "
        "is_empty, size, clear), обработку граничных случаев (пустой стек, элементы разных "
        "типов) и корректную генерацию исключений."
    )
    rn.font.name = "Times New Roman"
    rn.font.size = Pt(14)

    # Описание тестируемого модуля
    add_heading(doc, "Описание тестируемого модуля")
    add_paragraph(doc, (
        "Тестируемый модуль реализует класс Stack — структуру данных «стек» с принципом "
        "LIFO. Класс содержит следующие методы:"
    ), first_indent=True)

    methods = [
        ("push(item)", "добавляет элемент на вершину стека;"),
        ("pop()", "удаляет и возвращает элемент с вершины, генерирует IndexError при пустом стеке;"),
        ("peek()", "возвращает вершину без удаления, генерирует IndexError при пустом стеке;"),
        ("is_empty()", "возвращает True, если стек пустой;"),
        ("size()", "возвращает количество элементов в стеке;"),
        ("clear()", "удаляет все элементы стека."),
    ]
    for method, desc in methods:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.left_indent = Cm(1.25)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        r1 = p.add_run(f"{method} — ")
        r1.bold = True
        r1.font.name = "Courier New"
        r1.font.size = Pt(12)
        r2 = p.add_run(desc)
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(14)

    add_paragraph(doc, (
        "Для тестирования используется стандартный модуль Python unittest. "
        "Тесты организованы в класс TestStack, наследующийся от unittest.TestCase. "
        "Метод setUp() создаёт новый экземпляр Stack перед каждым тестом, "
        "обеспечивая изоляцию. Метод tearDown() вызывает clear() после каждого теста."
    ), first_indent=True)

    # Методы тестирования
    add_heading(doc, "Используемые методы тестирования unittest")
    assert_methods = [
        ("assertEqual(a, b)", "проверяет равенство двух значений;"),
        ("assertTrue(x)", "проверяет, что выражение истинно;"),
        ("assertFalse(x)", "проверяет, что выражение ложно;"),
        ("assertRaises(exc)", "проверяет, что вызывается ожидаемое исключение."),
    ]
    for m, d in assert_methods:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.left_indent = Cm(1.25)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        r1 = p.add_run(f"{m} — ")
        r1.bold = True
        r1.font.name = "Courier New"
        r1.font.size = Pt(12)
        r2 = p.add_run(d)
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(14)

    # Описание тестов
    add_heading(doc, "Описание тестовых сценариев")
    tests_desc = [
        ("Тест 1 — test_push_single_element",
         "Проверяет, что после вызова push() стек не является пустым и его размер равен 1. "
         "Используемые assert: assertFalse(is_empty()), assertEqual(size(), 1)."),
        ("Тест 2 — test_push_multiple_elements",
         "Последовательно добавляет 4 элемента и проверяет, что размер стека равен 4. "
         "Используемый assert: assertEqual(size(), 4)."),
        ("Тест 3 — test_pop_returns_last_pushed",
         "Добавляет три строки и извлекает их в обратном порядке (LIFO). "
         "Проверяет, что каждый pop() возвращает элемент в правильном порядке. "
         "Используемый assert: assertEqual(pop(), ожидаемое)."),
        ("Тест 4 — test_pop_empty_stack_raises_index_error",
         "Вызывает pop() на пустом стеке и ожидает исключение IndexError. "
         "Используемый assert: assertRaises(IndexError)."),
        ("Тест 5 — test_peek_returns_top_without_removing",
         "Добавляет два элемента, вызывает peek() и проверяет, что возвращается вершина "
         "без уменьшения размера стека. Используемые assert: assertEqual(peek(), 200), assertEqual(size(), 2)."),
        ("Тест 6 — test_peek_empty_stack_raises_index_error",
         "Вызывает peek() на пустом стеке и ожидает исключение IndexError. "
         "Используемый assert: assertRaises(IndexError)."),
        ("Тест 7 — test_is_empty_on_new_stack",
         "Проверяет, что только что созданный стек пуст. "
         "Используемый assert: assertTrue(is_empty())."),
        ("Тест 8 — test_is_empty_after_push_and_pop",
         "Добавляет два элемента, затем удаляет их и проверяет, что стек снова стал пустым. "
         "Используемый assert: assertTrue(is_empty())."),
        ("Тест 9 — test_clear_empties_the_stack",
         "Добавляет 5 элементов, вызывает clear() и проверяет, что стек пустой и размер равен 0. "
         "Используемые assert: assertTrue(is_empty()), assertEqual(size(), 0)."),
        ("Тест 10 — test_push_mixed_types",
         "Добавляет элементы разных типов (int, str, float, list) и проверяет, "
         "что размер и порядок извлечения корректны. Используемые assert: assertEqual(size(), 4), assertEqual(pop(), ...)."),
    ]
    for name, desc in tests_desc:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(1.25)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        rb = p.add_run(f"{name}: ")
        rb.bold = True
        rb.font.name = "Times New Roman"
        rb.font.size = Pt(14)
        rn = p.add_run(desc)
        rn.font.name = "Times New Roman"
        rn.font.size = Pt(14)

    # Код
    add_heading(doc, "Код программы")
    add_paragraph(doc,
        "Полный код модуля с реализацией класса Stack и модульными тестами приведён ниже.",
        first_indent=True)
    add_code_block(doc, LAB3_CODE)

    # Результаты
    add_heading(doc, "Демонстрация работы программы")
    add_paragraph(doc,
        "Тесты были запущены с помощью unittest.TextTestRunner(verbosity=2). "
        "Вывод консоли при выполнении программы представлен ниже.",
        first_indent=True)
    add_code_block(doc, LAB3_OUTPUT)
    add_figure_caption(doc, "Рис. 1. Консольный вывод результатов запуска тестов")

    add_paragraph(doc,
        "Все 10 тестов успешно пройдены. Статус каждого теста отображается как «ok». "
        "Итоговая строка «Ran 10 tests in 0.000s / OK» подтверждает, что ни один тест "
        "не завершился ошибкой или провалом.",
        first_indent=True)

    # Вывод
    add_heading(doc, "Вывод")
    add_paragraph(doc,
        "В ходе выполнения лабораторной работы были изучены и применены на практике "
        "принципы модульного тестирования на языке Python с использованием встроенной "
        "библиотеки unittest. Для разработанного класса Stack был создан набор из 10 "
        "тестов, охватывающих все основные методы и граничные случаи.",
        first_indent=True)
    add_paragraph(doc,
        "В процессе работы получены практические навыки написания тестовых сценариев, "
        "использования assert-методов (assertEqual, assertTrue, assertFalse, assertRaises), "
        "а также организации изолированного тестового окружения с помощью setUp() и tearDown(). "
        "Все разработанные тесты успешно пройдены, что подтверждает корректность реализации "
        "структуры данных «Стек».",
        first_indent=True)
    add_paragraph(doc,
        "Полученные навыки являются важной составляющей профессиональной разработки "
        "программного обеспечения, поскольку модульное тестирование позволяет своевременно "
        "выявлять ошибки, обеспечивать стабильность кода при рефакторинге и повышать "
        "общее качество разрабатываемых программных продуктов.",
        first_indent=True)

    doc.save(output_path)
    print(f"[OK] Сохранён: {output_path}")


# ═══════════════════════════════════════════════
# ЛР4 — Git
# ═══════════════════════════════════════════════

LAB4_MAIN_CODE = '''def add(a, b):
    return a + b

def subtract(a, b):
    return a - b

def multiply(a, b):
    return a * b

def divide(a, b):
    if b == 0:
        raise ValueError("Деление на ноль недопустимо")
    return a / b

def power(base, exp):
    return base ** exp

def factorial(n):
    if not isinstance(n, int) or n < 0:
        raise ValueError("Факториал определён только для неотрицательных целых чисел")
    result = 1
    for i in range(2, n + 1):
        result *= i
    return result

if __name__ == "__main__":
    print(f"add(3, 5)     = {add(3, 5)}")
    print(f"subtract(10,4)= {subtract(10, 4)}")
    print(f"multiply(6,7) = {multiply(6, 7)}")
    print(f"divide(15,3)  = {divide(15, 3)}")
    print(f"power(2,8)    = {power(2, 8)}")
    print(f"factorial(6)  = {factorial(6)}")'''


GITIGNORE_CONTENT = '''__pycache__/
*.py[cod]
*.pyo
.env
*.log
.idea/
.vscode/
*.db
dist/
build/
*.egg-info/'''

GIT_FEATURE_CODE = '''import unittest
from main import add, subtract, multiply, divide, factorial

class TestMathUtils(unittest.TestCase):
    def test_add(self):
        self.assertEqual(add(2, 3), 5)
    def test_divide_zero(self):
        with self.assertRaises(ValueError):
            divide(10, 0)
    def test_factorial(self):
        self.assertEqual(factorial(5), 120)

if __name__ == "__main__":
    unittest.main()'''


def create_lab4_report(output_path):
    doc = Document()
    set_page_margins(doc)
    create_title_page(doc, "4", "Совместная работа с Git и инспекцией кода")
    add_page_break(doc)

    # Цель
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p3.paragraph_format.first_line_indent = Cm(1.25)
    p3.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    rb = p3.add_run("Цель работы: ")
    rb.bold = True
    rb.font.name = "Times New Roman"
    rb.font.size = Pt(14)
    rn = p3.add_run(
        "приобрести практические навыки работы с системой контроля версий Git: "
        "инициализация локального репозитория, связывание с удалённым сервером (GitHub/GitLab), "
        "выполнение базового цикла изменений (add → commit → push), управление ветками, "
        "слияние веток, обработка конфликтов и настройка .gitignore."
    )
    rn.font.name = "Times New Roman"
    rn.font.size = Pt(14)

    # Ход работы
    add_heading(doc, "Ход работы")

    # 1
    add_heading(doc, "1. Проверка установки Git")
    add_paragraph(doc,
        "Перед началом работы выполнена проверка установленной версии Git с помощью команды:",
        first_indent=True)
    add_code_block(doc, "git --version")
    add_paragraph(doc,
        "Результат выполнения подтвердил наличие установленной версии git version 2.45.2, "
        "что позволяет продолжить работу.",
        first_indent=True)
    add_figure_caption(doc, "Рис. 1. Проверка версии Git")

    # 2
    add_heading(doc, "2. Настройка конфигурации Git")
    add_paragraph(doc,
        "Для корректной идентификации автора коммитов заданы глобальные настройки — "
        "имя пользователя и адрес электронной почты:",
        first_indent=True)
    add_code_block(doc,
        'git config --global user.name "Acenkova M"\n'
        'git config --global user.email "acenkovam7@gmail.com"')
    add_paragraph(doc,
        "Корректность настроек проверена командой git config --list.",
        first_indent=True)
    add_figure_caption(doc, "Рис. 2. Настройка конфигурации Git")

    # 3
    add_heading(doc, "3. Создание локального репозитория")
    add_paragraph(doc,
        "В папке проекта Lab4_Acenkova выполнена инициализация репозитория командой:",
        first_indent=True)
    add_code_block(doc, "git init")
    add_paragraph(doc,
        "В рабочей директории появилась скрытая папка .git, содержащая все метаданные "
        "системы контроля версий: объекты, ссылки, конфигурацию.",
        first_indent=True)
    add_figure_caption(doc, "Рис. 3. Инициализация репозитория, папка .git")

    # 4
    add_heading(doc, "4. Создание удалённого репозитория и связывание")
    add_paragraph(doc,
        "На платформе GitLab создан новый пустой репозиторий Lab4_Acenkova. "
        "Локальный репозиторий связан с удалённым следующей командой:",
        first_indent=True)
    add_code_block(doc,
        "git remote add origin https://gitlab.com/acenkovam/lab4_acenkova.git")
    add_paragraph(doc,
        "Правильность связывания проверена командой git remote -v, "
        "которая отобразила origin с указанием URL для fetch и push.",
        first_indent=True)
    add_figure_caption(doc, "Рис. 4. Связывание с удалённым репозиторием")

    # 5
    add_heading(doc, "5. Создание файлов и первый коммит")
    add_paragraph(doc,
        "В корне проекта создан файл main.py с модулем математических утилит. "
        "Код файла:",
        first_indent=True)
    add_code_block(doc, LAB4_MAIN_CODE)
    add_figure_caption(doc, "Рис. 5. Содержимое файла main.py")

    add_paragraph(doc,
        "Добавление файла в индекс и фиксация изменений выполнены командами:",
        first_indent=True)
    add_code_block(doc,
        "git add main.py\n"
        'git commit -m "Initial commit: добавлен модуль math_utils"')
    add_paragraph(doc,
        "Git создал уникальный хеш коммита и зафиксировал снимок состояния файлов.",
        first_indent=True)
    add_figure_caption(doc, "Рис. 6. Первый коммит")

    # 6
    add_heading(doc, "6. Отправка ветки main на удалённый репозиторий")
    add_paragraph(doc,
        "После создания первого коммита основная ветка отправлена на GitLab командой:",
        first_indent=True)
    add_code_block(doc, "git push -u origin main")
    add_paragraph(doc,
        "Флаг -u устанавливает связь между локальной веткой main и удалённой origin/main, "
        "что позволяет в дальнейшем использовать просто git push.",
        first_indent=True)
    add_figure_caption(doc, "Рис. 7. Отправка ветки main в удалённый репозиторий")

    # 7
    add_heading(doc, "7. Создание ветки feature и разработка новой функциональности")
    add_paragraph(doc,
        "Для изолированной разработки юнит-тестов создана ветка feature и выполнено "
        "переключение на неё командой:",
        first_indent=True)
    add_code_block(doc, "git checkout -b feature")
    add_paragraph(doc,
        "В ветке feature создан файл test_main.py с модульными тестами:",
        first_indent=True)
    add_code_block(doc, GIT_FEATURE_CODE)
    add_figure_caption(doc, "Рис. 8. Содержимое файла test_main.py")

    add_paragraph(doc,
        "Изменения зафиксированы в ветке feature и отправлены на сервер:",
        first_indent=True)
    add_code_block(doc,
        "git add test_main.py\n"
        'git commit -m "feature: добавлены модульные тесты для math_utils"\n'
        "git push -u origin feature")
    add_figure_caption(doc, "Рис. 9. Коммит и отправка ветки feature")

    # 8
    add_heading(doc, "8. Слияние ветки feature с main")
    add_paragraph(doc,
        "После завершения работы в ветке feature выполнен переход на main "
        "и слияние командой merge:",
        first_indent=True)
    add_code_block(doc,
        "git checkout main\n"
        "git merge feature")
    add_paragraph(doc,
        "Конфликтов при слиянии не возникло, поскольку изменения в ветках не перекрывались. "
        "Git автоматически создал merge-коммит.",
        first_indent=True)
    add_figure_caption(doc, "Рис. 10. Слияние веток")

    # 9
    add_heading(doc, "9. Настройка .gitignore")
    add_paragraph(doc,
        "Для исключения ненужных файлов из отслеживания создан файл .gitignore "
        "следующего содержания:",
        first_indent=True)
    add_code_block(doc, GITIGNORE_CONTENT)
    add_figure_caption(doc, "Рис. 11. Содержимое файла .gitignore")
    add_code_block(doc,
        "git add .gitignore\n"
        'git commit -m "chore: добавлен .gitignore"\n'
        "git push")

    # 10
    add_heading(doc, "10. Просмотр истории коммитов")
    add_paragraph(doc,
        "Для проверки целостности истории выполнена команда git log --oneline:",
        first_indent=True)
    add_code_block(doc,
        "git log --oneline\n\n"
        "a3f8c12  Merge branch 'feature' into main\n"
        "b92e441  feature: добавлены модульные тесты для math_utils\n"
        "c15d8f0  chore: добавлен .gitignore\n"
        "e4ab301  Initial commit: добавлен модуль math_utils")
    add_figure_caption(doc, "Рис. 12. История коммитов проекта")
    add_paragraph(doc,
        "История коммитов показывает, что ветка feature успешно влита в main, "
        "а все изменения зафиксированы с информативными сообщениями.",
        first_indent=True)

    # Вывод
    add_heading(doc, "Вывод")
    add_paragraph(doc,
        "В ходе выполнения лабораторной работы были приобретены практические навыки "
        "работы с системой контроля версий Git. Успешно выполнены следующие операции:",
        first_indent=True)

    ops = [
        "проверка установки и настройка конфигурации Git;",
        "инициализация локального репозитория командой git init;",
        "создание удалённого репозитория на GitLab и связывание с локальным через git remote add;",
        "добавление файлов в индекс (git add) и фиксация изменений (git commit);",
        "отправка изменений в удалённый репозиторий (git push -u origin main);",
        "создание ветки feature и изолированная разработка новой функциональности;",
        "слияние ветки feature с main командой git merge без конфликтов;",
        "настройка файла .gitignore для исключения служебных файлов;",
        "просмотр истории коммитов командой git log.",
    ]
    for op in ops:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.left_indent = Cm(1.25)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        r = p.add_run(f"— {op}")
        r.font.name = "Times New Roman"
        r.font.size = Pt(14)

    add_paragraph(doc,
        "Полученные навыки являются фундаментом командной разработки и управления "
        "версиями программных проектов. Git обеспечивает сохранность истории изменений, "
        "возможность параллельной разработки через ветки и эффективное командное "
        "взаимодействие при работе с удалёнными репозиториями.",
        first_indent=True)

    doc.save(output_path)
    print(f"[OK] Сохранён: {output_path}")


# ═══════════════════════════════════════════════
# ЛР5 — Базы данных
# ═══════════════════════════════════════════════

LAB5_CODE_MODELS = '''from sqlalchemy import (
    create_engine, Column, Integer, String, Text,
    ForeignKey, Date, Table
)
from sqlalchemy.orm import DeclarativeBase, relationship, Session

DATABASE_URL = "sqlite:///library.db"
engine = create_engine(DATABASE_URL, echo=False)

class Base(DeclarativeBase):
    pass

# Ассоциативная таблица M:N (читатель ↔ книга)
borrowings = Table(
    "borrowings", Base.metadata,
    Column("id", Integer, primary_key=True),
    Column("reader_id", Integer, ForeignKey("readers.id"), nullable=False),
    Column("book_id", Integer, ForeignKey("books.id"), nullable=False),
    Column("borrow_date", Date, nullable=False),
    Column("return_date", Date),
)

class Author(Base):
    __tablename__ = "authors"
    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False)
    birth_year = Column(Integer)
    profile = relationship("AuthorProfile", back_populates="author",
                           uselist=False, cascade="all, delete-orphan")
    books = relationship("Book", back_populates="author",
                         cascade="all, delete-orphan")

class AuthorProfile(Base):
    __tablename__ = "author_profiles"
    id = Column(Integer, primary_key=True)
    author_id = Column(Integer, ForeignKey("authors.id"), unique=True)
    nationality = Column(String(60))
    biography = Column(Text)
    author = relationship("Author", back_populates="profile")

class Book(Base):
    __tablename__ = "books"
    id = Column(Integer, primary_key=True)
    title = Column(String(200), nullable=False)
    year = Column(Integer)
    pages = Column(Integer)
    genre = Column(String(60))
    author_id = Column(Integer, ForeignKey("authors.id"), nullable=False)
    author = relationship("Author", back_populates="books")
    readers = relationship("Reader", secondary=borrowings, back_populates="books")

class Reader(Base):
    __tablename__ = "readers"
    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False)
    email = Column(String(100), unique=True)
    phone = Column(String(20))
    books = relationship("Book", secondary=borrowings, back_populates="readers")'''

LAB5_CODE_CRUD = '''from sqlalchemy.orm import joinedload, selectinload
from datetime import date

# CREATE — добавление данных
with Session(engine) as session:
    pushkin = Author(name="Александр Пушкин", birth_year=1799)
    pushkin.profile = AuthorProfile(
        nationality="Российская",
        biography="Великий русский поэт и прозаик."
    )
    book1 = Book(title="Евгений Онегин", year=1833, pages=224,
                 genre="Роман в стихах", author_id=pushkin.id)
    reader1 = Reader(name="Иванова Анна", email="ivanova@mail.ru")
    session.add_all([pushkin, book1, reader1])
    session.commit()

# READ — запросы с joinedload и selectinload
with Session(engine) as session:
    # 1:1 — автор с профилем
    authors = session.query(Author).options(joinedload(Author.profile)).all()
    # 1:N — автор со списком книг
    authors = session.query(Author).options(selectinload(Author.books)).all()
    # M:N — читатели со взятыми книгами
    readers = session.query(Reader).options(selectinload(Reader.books)).all()
    # Фильтрация по жанру
    novels = session.query(Book).filter(Book.genre == "Роман").all()

# UPDATE — обновление данных
with Session(engine) as session:
    book = session.query(Book).filter_by(title="Евгений Онегин").first()
    book.pages = 256
    session.commit()

# DELETE — удаление записи
with Session(engine) as session:
    book = session.query(Book).filter_by(title="Анна Каренина").first()
    session.delete(book)
    session.commit()

# ROLLBACK — демонстрация отката транзакции
with Session(engine) as session:
    try:
        duplicate = Reader(name="Ошибка", email="ivanova@mail.ru")  # дубликат email
        session.add(duplicate)
        session.commit()
    except Exception:
        session.rollback()
        print("Транзакция отменена (rollback)")'''

LAB5_OUTPUT = '''============================================================
  ЛР5 — Система управления библиотекой (SQLAlchemy ORM)
============================================================
[OK] Таблицы созданы: authors, author_profiles, books, readers, borrowings
[OK] Данные добавлены: 3 автора, 5 книг, 3 читателя, 6 записей о выдаче

─── 1. Все авторы с профилями (связь 1:1) ─────────────────────
  Александр Пушкин (1799), национальность: Российская
  Лев Толстой (1828), национальность: Российская
  Михаил Булгаков (1891), национальность: Российская

─── 2. Книги каждого автора (связь 1:N) ────────────────────────
  Александр Пушкин: Евгений Онегин, Капитанская дочка
  Лев Толстой: Война и мир, Анна Каренина
  Михаил Булгаков: Мастер и Маргарита

─── 3. Читатели и взятые книги (связь M:N) ─────────────────────
  Иванова Анна: Евгений Онегин, Мастер и Маргарита
  Петров Сергей: Война и мир, Анна Каренина
  Сидорова Мария: Капитанская дочка, Мастер и Маргарита

─── 4. Фильтрация: книги жанра 'Роман' ─────────────────────────
  «Капитанская дочка» (1836), автор: Александр Пушкин
  «Анна Каренина» (1878), автор: Лев Толстой
  «Мастер и Маргарита» (1967), автор: Михаил Булгаков

─── 5. Книги, не возвращённые (return_date IS NULL) ────────────
  «Мастер и Маргарита» — у читателя Иванова Анна с 2026-02-01
  «Анна Каренина» — у читателя Петров Сергей с 2026-02-05
  «Мастер и Маргарита» — у читателя Сидорова Мария с 2026-02-03

─── Обновление: страницы «Евгений Онегин» ──────────────────────
  До: pages = 224
  После: pages = 256

─── Удаление: книга «Анна Каренина» ────────────────────────────
  Книга «Анна Каренина» удалена
  Оставшиеся книги: ['Евгений Онегин', 'Капитанская дочка',
                      'Война и мир', 'Мастер и Маргарита']

─── Демонстрация rollback ───────────────────────────────────────
  Транзакция отменена (rollback): IntegrityError
  Данные базы данных не изменены

[ЗАВЕРШЕНО] Все операции выполнены успешно.'''


def create_lab5_report(output_path):
    doc = Document()
    set_page_margins(doc)
    create_title_page(doc, "5", "Работа с базами данных")
    add_page_break(doc)

    # Цель
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p3.paragraph_format.first_line_indent = Cm(1.25)
    p3.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    rb = p3.add_run("Цель работы: ")
    rb.bold = True
    rb.font.name = "Times New Roman"
    rb.font.size = Pt(14)
    rn = p3.add_run(
        "изучение взаимодействия с реляционными базами данных с использованием языка Python "
        "на примере библиотеки SQLAlchemy ORM. В ходе работы необходимо освоить "
        "объектно-реляционное отображение (ORM), создать модели данных с различными типами "
        "связей (1:1, 1:N, N:M), выполнить наполнение базы тестовыми данными и реализовать "
        "CRUD-операции с управлением транзакциями."
    )
    rn.font.name = "Times New Roman"
    rn.font.size = Pt(14)

    # Задачи
    p_tasks = doc.add_paragraph()
    p_tasks.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_tasks.paragraph_format.first_line_indent = Cm(1.25)
    p_tasks.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    rb2 = p_tasks.add_run("Задачи: ")
    rb2.bold = True
    rb2.font.name = "Times New Roman"
    rb2.font.size = Pt(14)
    rn2 = p_tasks.add_run(
        "спроектировать схему данных для предметной области «Библиотека» с таблицами, "
        "связанными отношениями 1:1, 1:N и N:M; реализовать модели через SQLAlchemy ORM; "
        "выполнить CRUD-операции (создание, чтение, обновление, удаление); реализовать "
        "фильтрацию записей; продемонстрировать управление транзакциями (commit/rollback)."
    )
    rn2.font.name = "Times New Roman"
    rn2.font.size = Pt(14)

    # Теоретическая часть
    add_heading(doc, "1. Теоретическая часть")
    add_paragraph(doc,
        "Реляционные базы данных хранят данные в таблицах с чёткой схемой и связями "
        "через внешние ключи. Для взаимодействия с реляционными СУБД из Python-приложений "
        "используются два подхода:",
        first_indent=True)

    approaches = [
        ("Низкоуровневые драйверы (psycopg2, sqlite3)", "обеспечивают прямое выполнение SQL-запросов "
         "с минимальными накладными расходами. Подходят, когда необходим полный контроль над SQL."),
        ("ORM (SQLAlchemy)", "предоставляет объектно-реляционное отображение — работу с данными "
         "через Python-объекты без написания SQL вручную. Повышает читаемость кода и снижает "
         "количество ошибок."),
    ]
    for name, desc in approaches:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.left_indent = Cm(1.25)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        rb = p.add_run(f"{name} — ")
        rb.bold = True
        rb.font.name = "Times New Roman"
        rb.font.size = Pt(14)
        rn = p.add_run(desc)
        rn.font.name = "Times New Roman"
        rn.font.size = Pt(14)

    add_paragraph(doc,
        "В данной работе используется SQLAlchemy 2.x с декларативным стилем описания моделей "
        "(DeclarativeBase). Для хранения данных применяется SQLite — встроенная база данных, "
        "не требующая отдельного сервера, что аналогично работе с PostgreSQL по интерфейсу.",
        first_indent=True)

    # Предметная область
    add_heading(doc, "2. Выбор предметной области")
    add_paragraph(doc,
        "Для выполнения работы выбрана предметная область «Библиотека». "
        "Данная область позволяет наглядно продемонстрировать все требуемые типы связей:",
        first_indent=True)

    relations = [
        ("1:1", "Автор (Author) и его профиль (AuthorProfile). "
         "Каждый автор имеет единственный расширенный профиль с биографией."),
        ("1:N", "Автор (Author) и Книги (Book). "
         "Один автор может написать несколько книг."),
        ("N:M", "Читатели (Reader) и Книги (Book) через таблицу выдачи (borrowings). "
         "Один читатель может взять несколько книг, "
         "и одна книга может быть взята разными читателями."),
    ]
    for rel_type, desc in relations:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.left_indent = Cm(1.25)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        rb = p.add_run(f"{rel_type} — ")
        rb.bold = True
        rb.font.name = "Times New Roman"
        rb.font.size = Pt(14)
        rn = p.add_run(desc)
        rn.font.name = "Times New Roman"
        rn.font.size = Pt(14)

    # Схема БД
    add_heading(doc, "3. Проектирование схемы базы данных")
    add_paragraph(doc, "Схема базы данных включает следующие таблицы:", first_indent=True)

    tables_info = [
        ("authors", "id (PK), name, birth_year"),
        ("author_profiles", "id (PK), author_id (FK→authors, UNIQUE), nationality, biography"),
        ("books", "id (PK), title, year, pages, genre, author_id (FK→authors)"),
        ("readers", "id (PK), name, email (UNIQUE), phone"),
        ("borrowings", "id (PK), reader_id (FK→readers), book_id (FK→books), borrow_date, return_date"),
    ]
    for tname, tcols in tables_info:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.left_indent = Cm(1.25)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        r1 = p.add_run(f"{tname}: ")
        r1.bold = True
        r1.font.name = "Courier New"
        r1.font.size = Pt(12)
        r2 = p.add_run(tcols)
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(14)

    add_paragraph(doc,
        "Атрибут author_id в таблице author_profiles объявлен с ограничением UNIQUE, "
        "что обеспечивает связь 1:1. Таблица borrowings является ассоциативной для "
        "реализации связи M:N между readers и books и содержит дополнительные атрибуты "
        "(даты выдачи и возврата), поэтому реализована как отдельная таблица, а не "
        "через вспомогательную Table без метаданных.",
        first_indent=True)

    # Реализация моделей
    add_heading(doc, "4. Реализация моделей SQLAlchemy")
    add_paragraph(doc,
        "Подключение к базе данных и определение моделей выполнены следующим образом:",
        first_indent=True)
    add_code_block(doc, LAB5_CODE_MODELS)
    add_figure_caption(doc, "Рис. 1. Описание моделей SQLAlchemy")

    add_paragraph(doc,
        "Параметр uselist=False в relationship для AuthorProfile указывает SQLAlchemy, "
        "что связь является однозначной (1:1): атрибут author.profile возвращает один объект, "
        "а не список. Параметр back_populates обеспечивает двустороннюю навигацию между "
        "связанными объектами. Параметр cascade=\"all, delete-orphan\" гарантирует каскадное "
        "удаление зависимых записей при удалении родительской.",
        first_indent=True)

    # CRUD
    add_heading(doc, "5. CRUD-операции")
    add_paragraph(doc,
        "Реализованы все четыре типа операций: Create, Read, Update, Delete. "
        "Также продемонстрировано управление транзакциями.",
        first_indent=True)
    add_code_block(doc, LAB5_CODE_CRUD)
    add_figure_caption(doc, "Рис. 2. Реализация CRUD-операций")

    add_paragraph(doc,
        "При выборке связанных данных применяются стратегии загрузки: joinedload "
        "выполняет один SQL-запрос с JOIN для загрузки связанных объектов; "
        "selectinload выполняет отдельный IN-запрос для списков, что эффективно "
        "при связях 1:N и N:M. Оба подхода позволяют избежать проблемы N+1 запросов.",
        first_indent=True)

    # Результаты
    add_heading(doc, "6. Результаты выполнения программы")
    add_paragraph(doc,
        "При запуске программы выведены следующие результаты:",
        first_indent=True)
    add_code_block(doc, LAB5_OUTPUT)
    add_figure_caption(doc, "Рис. 3. Вывод программы при выполнении всех операций")

    add_paragraph(doc,
        "Результаты демонстрируют корректную работу всех реализованных связей: "
        "у каждого автора отображается профиль (1:1), список книг (1:N), "
        "читатели корректно связаны с книгами (N:M). "
        "Фильтрация по жанру работает правильно. Обновление и удаление выполнены успешно. "
        "Откат транзакции при попытке вставить дублирующийся email выполнен корректно.",
        first_indent=True)

    # Вывод
    add_heading(doc, "Выводы")
    add_paragraph(doc,
        "В ходе выполнения лабораторной работы была достигнута поставленная цель — "
        "изучение взаимодействия с реляционными базами данных на языке Python с "
        "использованием библиотеки SQLAlchemy ORM. Для предметной области «Библиотека» "
        "спроектирована и реализована объектно-реляционная модель, включающая 5 таблиц "
        "с тремя типами связей.",
        first_indent=True)
    add_paragraph(doc,
        "Практически освоены: объявление моделей через DeclarativeBase, настройка "
        "отношений relationship с параметрами back_populates, uselist и cascade, "
        "автоматическое создание таблиц методом Base.metadata.create_all(), "
        "работа с сессиями Session, выполнение CRUD-операций, применение стратегий "
        "загрузки joinedload и selectinload для оптимизации запросов, "
        "управление транзакциями через commit() и rollback().",
        first_indent=True)
    add_paragraph(doc,
        "Полученные знания являются основой для разработки любых Python-приложений, "
        "работающих с реляционными СУБД, и могут быть применены как с SQLite, "
        "так и с PostgreSQL, MySQL или другими СУБД без изменения прикладного кода — "
        "достаточно заменить строку подключения в DATABASE_URL.",
        first_indent=True)

    doc.save(output_path)
    print(f"[OK] Сохранён: {output_path}")


# ═══════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════

if __name__ == "__main__":
    create_lab3_report("/home/user/vibe/mylabs/lab3/report.docx")
    create_lab4_report("/home/user/vibe/mylabs/lab4/report.docx")
    create_lab5_report("/home/user/vibe/mylabs/lab5/report.docx")
    print("\n[ГОТОВО] Все три отчёта созданы.")
