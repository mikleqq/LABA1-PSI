# Exam Ticket Generator (Lab 1)

A console application that asks for a student's last and first name, draws a random exam ticket number (1–20), and appends a record to an Excel journal.

## Features

- Prompts for `Last name:` and `First name:` (confirm with Enter).
- Names must not be empty and **must not contain digits**. Only letters, spaces, hyphens and apostrophes are accepted (`Anna-Maria`, `O'Brien`). Leading/trailing spaces are trimmed.
- Prints the drawn ticket: `Билет № 7`.
- Appends a row to `journal.xlsx` **immediately after each student**, so no data is lost if the program crashes. Existing rows are never modified.
- If `journal.xlsx` is open in Excel, the app shows a clear message and lets you close the file and retry instead of crashing.
- Press **ESC** at any prompt to exit.

`journal.xlsx` is created automatically with this header:

| Last name | First name | Номер билета | Дата и время |
|---|---|---|---|

## Requirements

- Python 3.9+
- [`openpyxl`](https://pypi.org/project/openpyxl/)

## Run

```bash
pip install openpyxl
python tickets.py
```

Works on Windows, Linux and macOS (single-key input is handled per platform).

## Tests

```bash
python -m unittest -v
```

## Design (SOLID)

| Class | Responsibility |
|---|---|
| `NameValidator` | Validates last/first names (no digits, no empty values) |
| `RandomTicketGenerator` | Draws a ticket number |
| `ExcelStorage` | Appends records to `journal.xlsx` (atomic write via a temp file) |
| `ConsoleUI` | Console input/output and ESC handling |
| `TicketApp` | Application flow; depends only on abstractions |
| `main()` | Composition root — the only place where concrete classes are chosen |

Abstractions (`Validator`, `TicketGenerator`, `JournalStorage`, `UserInterface`) let you swap Excel for CSV or a database, or the console for a GUI, without touching `TicketApp`. Tests use in-memory fakes instead of the console and Excel.

## Project structure

```
tickets.py        application code
test_tickets.py   unit tests
```
