import os
import random
import tempfile
import unittest
from datetime import datetime

from openpyxl import load_workbook

from tickets import (ExcelStorage, JournalStorage, NameValidator, RandomTicketGenerator, Record,
                     StorageBusyError, TicketApp, UserInterface, ValidationError, HEADER)


class FakeUI(UserInterface):
    def __init__(self, lines, retry=True):
        self.lines, self.retry, self.shown = list(lines), retry, []

    def read_line(self, prompt):
        return self.lines.pop(0) if self.lines else None  # кончились ответы = ESC

    def show(self, message):
        self.shown.append(message)

    def confirm_retry(self, message):
        self.shown.append(message)
        return self.retry


class MemoryStorage(JournalStorage):
    def __init__(self, busy_times=0):
        self.records, self.busy = [], busy_times

    def append(self, record):
        if self.busy:
            self.busy -= 1
            raise StorageBusyError("busy")
        self.records.append(record)


class FixedTicket:
    def next_ticket(self):
        return 7


NOW = lambda: datetime(2026, 1, 1, 12, 0, 0)


def app(ui, storage):
    return TicketApp(ui, NameValidator(), FixedTicket(), storage, NOW)


class TestValidator(unittest.TestCase):
    v = NameValidator()

    def test_valid(self):
        for ok in ("Иванов", "  Петров ", "Анна-Мария", "O'Brien", "Smith"):
            self.assertEqual(self.v.validate(ok), ok.strip())

    def test_digits_rejected(self):
        for bad in ("Иванов1", "123", "1Иван", "Ivan2", "Иван ٣"):
            with self.assertRaisesRegex(ValidationError, "Цифры"):
                self.v.validate(bad)

    def test_empty_and_symbols_rejected(self):
        for bad in ("", "   ", None, "Иван!", "a@b", "---", "Иван_"):
            with self.assertRaises(ValidationError, msg=repr(bad)):
                self.v.validate(bad)


class TestApp(unittest.TestCase):
    def test_happy_path_and_esc(self):
        st = MemoryStorage()
        app(FakeUI([" Иванов ", "Иван", "Петров", "Пётр"]), st).run()
        self.assertEqual([(r.last_name, r.first_name, r.ticket) for r in st.records],
                         [("Иванов", "Иван", 7), ("Петров", "Пётр", 7)])

    def test_invalid_input_reprompts(self):
        st = MemoryStorage()
        ui = FakeUI(["", "Иванов1", "Иванов", "   ", "Иван5", "Иван"])
        app(ui, st).run()
        self.assertEqual(len(st.records), 1)
        self.assertIn("Цифры вводить нельзя.", ui.shown)
        self.assertIn("Поле не может быть пустым.", ui.shown)

    def test_esc_on_second_prompt_saves_nothing(self):
        st = MemoryStorage()
        app(FakeUI(["Иванов"]), st).run()
        self.assertEqual(st.records, [])

    def test_busy_storage_retries(self):
        st = MemoryStorage(busy_times=2)
        app(FakeUI(["Иванов", "Иван"], retry=True), st).run()
        self.assertEqual(len(st.records), 1)

    def test_busy_storage_cancel(self):
        st = MemoryStorage(busy_times=5)
        ui = FakeUI(["Иванов", "Иван"], retry=False)
        app(ui, st).run()
        self.assertEqual(st.records, [])
        self.assertIn("Запись не сохранена.\n", ui.shown)


class TestGenerator(unittest.TestCase):
    def test_range(self):
        g = RandomTicketGenerator(1, 20, random.Random(0))
        self.assertTrue(all(1 <= g.next_ticket() <= 20 for _ in range(200)))


class TestExcelStorage(unittest.TestCase):
    def setUp(self):
        d = tempfile.TemporaryDirectory()
        self.addCleanup(d.cleanup)
        self.path = os.path.join(d.name, "journal.xlsx")

    def rows(self):
        return [list(r) for r in load_workbook(self.path).active.iter_rows(values_only=True)]

    def rec(self, name, n):
        return Record(name, "И", n, datetime(2026, 1, 1, 10, 0, 0))

    def test_header_and_append_keeps_old_rows(self):
        s = ExcelStorage(self.path)
        for n in ("А", "Б", "В"):
            s.append(self.rec(n, 1))
        first_three = self.rows()
        self.assertEqual(first_three[0], HEADER)
        s2 = ExcelStorage(self.path)  # "новый запуск"
        s2.append(self.rec("Г", 2))
        s2.append(self.rec("Д", 3))
        after = self.rows()
        self.assertEqual(len(after), 6)
        self.assertEqual(after[:4], first_three)

    def test_locked_file_raises_busy(self):
        from unittest import mock
        s = ExcelStorage(self.path)
        s.append(self.rec("А", 1))
        before = self.rows()
        with mock.patch("tickets.os.replace", side_effect=PermissionError):
            with self.assertRaises(StorageBusyError):
                s.append(self.rec("Б", 2))
        self.assertEqual(self.rows(), before)
        self.assertFalse(os.path.exists(self.path + ".tmp"))


if __name__ == "__main__":
    unittest.main()
