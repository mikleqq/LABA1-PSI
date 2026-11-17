"""Генератор экзаменационных билетов (Лаб. 1) - версия с соблюдением SOLID.

Установка:  pip install openpyxl
Запуск:     python tickets.py
Тесты:      python -m unittest -v

Структура (каждый класс отвечает за одно):
  NameValidator          - проверка фамилии/имени (без цифр)
  RandomTicketGenerator  - выбор номера билета
  ExcelStorage           - запись в journal.xlsx
  ConsoleUI              - ввод/вывод в консоли, чтение клавиш (ESC)
  TicketApp              - сценарий работы, знает только абстракции
  main()                 - единственное место, где собираются конкретные классы
"""
import os
import random
from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import datetime
from typing import Callable, Optional

ESC = "\x1b"
HEADER = ["Last name", "First name", "Номер билета", "Дата и время"]


# ============================================================ модель и ошибки
@dataclass(frozen=True)
class Record:
    last_name: str
    first_name: str
    ticket: int
    created_at: datetime


class ValidationError(ValueError):
    """Введённое значение не прошло проверку."""


class StorageError(Exception):
    """Не удалось записать в хранилище."""


class StorageBusyError(StorageError):
    """Хранилище занято (например, файл открыт в Excel)."""


# ================================================================ абстракции
class Validator(ABC):
    @abstractmethod
    def validate(self, text: str) -> str:
        """Возвращает очищенное значение или бросает ValidationError."""


class TicketGenerator(ABC):
    @abstractmethod
    def next_ticket(self) -> int: ...


class JournalStorage(ABC):
    @abstractmethod
    def append(self, record: Record) -> None:
        """Дописывает запись. Бросает StorageBusyError / StorageError."""


class UserInterface(ABC):
    @abstractmethod
    def read_line(self, prompt: str) -> Optional[str]:
        """Читает строку; None - пользователь нажал ESC."""

    @abstractmethod
    def show(self, message: str) -> None: ...

    @abstractmethod
    def confirm_retry(self, message: str) -> bool:
        """Спрашивает, повторить ли попытку. False - отказаться."""


# ============================================================== реализации
class NameValidator(Validator):
    """Фамилия/имя: только буквы, пробел, дефис и апостроф. Цифры запрещены."""

    ALLOWED_EXTRA = " -'’"

    def validate(self, text: str) -> str:
        value = (text or "").strip()
        if not value:
            raise ValidationError("Поле не может быть пустым.")
        if any(ch.isdigit() for ch in value):
            raise ValidationError("Цифры вводить нельзя.")
        if not all(ch.isalpha() or ch in self.ALLOWED_EXTRA for ch in value):
            raise ValidationError("Допустимы только буквы, пробел, дефис и апостроф.")
        if not any(ch.isalpha() for ch in value):
            raise ValidationError("Введите хотя бы одну букву.")
        return value


class RandomTicketGenerator(TicketGenerator):
    def __init__(self, first: int = 1, last: int = 20, rng: Optional[random.Random] = None):
        self._first, self._last = first, last
        self._rng = rng or random.Random()

    def next_ticket(self) -> int:
        return self._rng.randint(self._first, self._last)


class ExcelStorage(JournalStorage):
    """journal.xlsx: создаётся с шапкой, строки только дописываются.

    Сохранение через временный файл + os.replace: файл не портится при сбое,
    а если он открыт в Excel - замена не удастся и будет StorageBusyError.
    """

    def __init__(self, path: str):
        self._path = path

    def append(self, record: Record) -> None:
        from openpyxl import Workbook, load_workbook  # импорт здесь: модуль легко тестировать без Excel

        tmp = self._path + ".tmp"
        try:
            if os.path.exists(self._path):
                wb = load_workbook(self._path)
                ws = wb.active
            else:
                wb = Workbook()
                ws = wb.active
                ws.append(HEADER)
                for col, width in zip("ABCD", (20, 20, 14, 20)):
                    ws.column_dimensions[col].width = width
            ws.append([record.last_name, record.first_name, record.ticket, record.created_at])
            ws.cell(row=ws.max_row, column=4).number_format = "DD.MM.YYYY HH:MM:SS"
            wb.save(tmp)
            os.replace(tmp, self._path)
        except PermissionError as e:
            self._cleanup(tmp)
            raise StorageBusyError(str(e)) from e
        except Exception as e:
            self._cleanup(tmp)
            raise StorageError(str(e)) from e

    @staticmethod
    def _cleanup(tmp: str) -> None:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass


def _make_key_reader() -> Callable[[], str]:
    """Возвращает функцию чтения одной клавиши для текущей ОС."""
    if os.name == "nt":
        import msvcrt

        def getkey() -> str:
            ch = msvcrt.getwch()
            if ch in ("\x00", "\xe0"):  # стрелки, F-клавиши
                msvcrt.getwch()
                return ""
            return ch
    else:
        import select
        import sys
        import termios
        import tty

        def getkey() -> str:
            fd = sys.stdin.fileno()
            old = termios.tcgetattr(fd)
            try:
                tty.setraw(fd)
                ch = sys.stdin.read(1)
                if ch == ESC and select.select([sys.stdin], [], [], 0.05)[0]:
                    while select.select([sys.stdin], [], [], 0.01)[0]:
                        sys.stdin.read(1)
                    return ""  # стрелка/спец. клавиша, а не одиночный ESC
                return ch
            finally:
                termios.tcsetattr(fd, termios.TCSADRAIN, old)
    return getkey


class ConsoleUI(UserInterface):
    def __init__(self, key_reader: Optional[Callable[[], str]] = None):
        self._getkey = key_reader or _make_key_reader()

    def read_line(self, prompt: str) -> Optional[str]:
        print(prompt, end="", flush=True)
        buf = []
        while True:
            ch = self._getkey()
            if ch in (ESC, "\x03"):  # ESC или Ctrl+C
                print()
                return None
            if ch in ("\r", "\n"):
                print()
                return "".join(buf)
            if ch in ("\x08", "\x7f"):
                if buf:
                    buf.pop()
                    print("\b \b", end="", flush=True)
            elif ch and ch.isprintable():
                buf.append(ch)
                print(ch, end="", flush=True)

    def show(self, message: str) -> None:
        print(message)

    def confirm_retry(self, message: str) -> bool:
        print(message)
        print("Закройте файл и нажмите любую клавишу для повтора (ESC - отменить запись).")
        return self._getkey() != ESC


# ================================================================== сценарий
class TicketApp:
    """Сценарий работы. Зависит только от абстракций (DIP)."""

    def __init__(self, ui: UserInterface, validator: Validator, generator: TicketGenerator,
                 storage: JournalStorage, clock: Callable[[], datetime] = datetime.now):
        self._ui, self._validator, self._generator = ui, validator, generator
        self._storage, self._clock = storage, clock

    def run(self) -> None:
        self._ui.show("Для выхода нажмите ESC.")
        while True:
            last = self._ask("Last name: ")
            if last is None:
                break
            first = self._ask("First name: ")
            if first is None:
                break
            ticket = self._generator.next_ticket()
            self._ui.show(f"Билет № {ticket}")
            record = Record(last, first, ticket, self._clock().replace(microsecond=0))
            self._ui.show("Запись сохранена.\n" if self._save(record) else "Запись не сохранена.\n")
        self._ui.show("Работа завершена.")

    def _ask(self, prompt: str) -> Optional[str]:
        while True:
            raw = self._ui.read_line(prompt)
            if raw is None:
                return None
            try:
                return self._validator.validate(raw)
            except ValidationError as e:
                self._ui.show(str(e))

    def _save(self, record: Record) -> bool:
        while True:
            try:
                self._storage.append(record)
                return True
            except StorageBusyError:
                msg = "Не удаётся записать журнал: файл открыт в Excel или другой программе."
            except StorageError as e:
                msg = f"Ошибка при записи в файл: {e}"
            if not self._ui.confirm_retry(msg):
                return False


def main() -> None:
    """Точка сборки: здесь выбираются конкретные реализации."""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "journal.xlsx")
    TicketApp(
        ui=ConsoleUI(),
        validator=NameValidator(),
        generator=RandomTicketGenerator(1, 20),
        storage=ExcelStorage(path),
    ).run()


if __name__ == "__main__":
    main()
