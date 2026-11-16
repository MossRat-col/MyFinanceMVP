import subprocess
import tkinter as tk
from decimal import Decimal, InvalidOperation
from tkinter import messagebox, ttk

import excel_storage
import logic
import settings


TEXT = {
    "ru": {
        "window_title": "Учёт финансов",
        "language": "Язык",
        "currency": "Счёт в валюте",
        "russian": "Русский",
        "english": "English (US)",
        "uzs_account": "Сум (UZS)",
        "usd_account": "Доллар США (USD)",
        "balance": "Баланс счёта",
        "today": "Итог за сегодня: {amount}",
        "amount": "Сумма операции ({currency})",
        "income": "+ Доход",
        "expense": "− Расход",
        "show_table": "Показать в таблице",
        "reset": "Сбросить счёт {currency}",
        "input_error": "Ошибка ввода",
        "enter_amount": "Введите сумму операции.",
        "whole_number": "Введите число без букв и посторонних символов.",
        "uzs_whole_number": "Для счёта UZS введите целое число сумов.",
        "usd_precision": "Для USD можно указать не более двух знаков после запятой.",
        "positive_amount": "Сумма должна быть больше нуля.",
        "error": "Ошибка",
        "transaction_error": "Не удалось обработать операцию:\n{error}",
        "open_error_title": "Ошибка открытия",
        "open_error": "Не удалось открыть Excel-файл:\n{error}",
        "reset_title": "Сброс счёта",
        "reset_question": "Удалить все данные только со счёта {currency}? Это действие нельзя отменить.",
        "reset_done_title": "Сброс выполнен",
        "reset_done": "Данные счёта {currency} удалены.",
        "reset_error_title": "Ошибка сброса",
        "reset_error": "Не удалось сбросить счёт. Возможно, файл открыт в Excel.\n\n{error}",
    },
    "en_US": {
        "window_title": "Personal Finance Tracker",
        "language": "Language",
        "currency": "Account currency",
        "russian": "Русский",
        "english": "English (US)",
        "uzs_account": "Uzbek sum (UZS)",
        "usd_account": "US dollar (USD)",
        "balance": "Account balance",
        "today": "Today's total: {amount}",
        "amount": "Transaction amount ({currency})",
        "income": "+ Income",
        "expense": "− Expense",
        "show_table": "Open spreadsheet",
        "reset": "Reset {currency} account",
        "input_error": "Input error",
        "enter_amount": "Enter a transaction amount.",
        "whole_number": "Enter a number using digits and a decimal separator only.",
        "uzs_whole_number": "Enter a whole number of sum for the UZS account.",
        "usd_precision": "USD amounts can have up to two decimal places.",
        "positive_amount": "The amount must be greater than zero.",
        "error": "Error",
        "transaction_error": "Could not process the transaction:\n{error}",
        "open_error_title": "Could not open file",
        "open_error": "Could not open the Excel file:\n{error}",
        "reset_title": "Reset account",
        "reset_question": "Delete all data from the {currency} account only? This cannot be undone.",
        "reset_done_title": "Reset complete",
        "reset_done": "The {currency} account data has been deleted.",
        "reset_error_title": "Reset error",
        "reset_error": "Could not reset the account. The file may be open in Excel.\n\n{error}",
    },
}


class FinanceApp:
    def __init__(self, root):
        self.root = root
        self.preferences = settings.load_settings()
        self.language = self.preferences["language"]
        self.currency = self.preferences["currency"]

        root.geometry("700x600")
        root.minsize(600, 500)
        root.resizable(False, False)

        self.language_label = tk.Label(root)
        self.language_label.pack(anchor="e", padx=20, pady=(10, 0))
        self.language_choice = tk.StringVar()
        self.language_menu = ttk.Combobox(
            root,
            textvariable=self.language_choice,
            values=("Русский", "English (US)"),
            state="readonly",
            width=18,
        )
        self.language_menu.pack(anchor="e", padx=20)
        self.language_menu.bind("<<ComboboxSelected>>", self._language_selected)

        self.currency_label = tk.Label(root)
        self.currency_label.pack(anchor="e", padx=20, pady=(5, 0))
        self.currency_choice = tk.StringVar()
        self.currency_menu = ttk.Combobox(
            root,
            textvariable=self.currency_choice,
            values=("UZS", "USD"),
            state="readonly",
            width=18,
        )
        self.currency_menu.pack(anchor="e", padx=20)
        self.currency_menu.bind("<<ComboboxSelected>>", self._currency_selected)

        self.title_label = tk.Label(root, font=("Arial", 18, "bold"))
        self.title_label.pack(pady=(10, 8))
        self.balance_heading = tk.Label(root, font=("Arial", 11))
        self.balance_heading.pack()
        self.balance_label = tk.Label(
            root, text="0 UZS", font=("Arial", 20, "bold"), fg="#18794e"
        )
        self.balance_label.pack(pady=5)
        self.today_label = tk.Label(root, font=("Arial", 11))
        self.today_label.pack(pady=5)
        self.amount_label = tk.Label(root)
        self.amount_label.pack(pady=(12, 5))

        self.amount_entry = tk.Entry(root, font=("Arial", 14), justify="center")
        self.amount_entry.pack(padx=40, fill="x")
        self.amount_entry.focus_set()

        buttons = tk.Frame(root)
        buttons.pack(pady=16)
        self.income_button = tk.Button(
            buttons,
            width=12,
            bg="#18794e",
            fg="white",
            command=lambda: self.add_transaction("income"),
        )
        self.income_button.pack(side="left", padx=5)
        self.expense_button = tk.Button(
            buttons,
            width=12,
            bg="#b42318",
            fg="white",
            command=lambda: self.add_transaction("expense"),
        )
        self.expense_button.pack(side="left", padx=5)
        self.table_button = tk.Button(root, command=self.open_excel)
        self.table_button.pack(pady=(0, 8))
        self.reset_button = tk.Button(root, command=self.reset_balance)
        self.reset_button.pack(pady=(0, 8))

        self._apply_language()

    def _text(self, key, **values):
        return TEXT[self.language][key].format(**values)

    def _language_selected(self, _event=None):
        self.language = "en_US" if self.language_choice.get() == "English (US)" else "ru"
        self.preferences["language"] = self.language
        self._save_preferences()
        self._apply_language()

    def _currency_selected(self, _event=None):
        self.currency = self.currency_choice.get()
        self.preferences["currency"] = self.currency
        self._save_preferences()
        self._apply_language()

    def _save_preferences(self):
        try:
            settings.save_settings(self.preferences)
        except OSError as error:
            messagebox.showerror(self._text("error"), str(error))

    def _format_amount(self, amount):
        amount = Decimal(str(amount))
        if self.currency == "USD":
            return f"${amount:,.2f}"
        return f"{amount:,.0f} UZS".replace(",", " ")

    def _apply_language(self):
        self.root.title(self._text("window_title"))
        self.language_label.config(text=self._text("language"))
        self.language_choice.set(
            self._text("english") if self.language == "en_US" else self._text("russian")
        )
        self.currency_label.config(text=self._text("currency"))
        self.currency_choice.set(self.currency)
        self.title_label.config(text=self._text("window_title"))
        self.balance_heading.config(text=self._text("balance"))
        self.amount_label.config(text=self._text("amount", currency=self.currency))
        self.income_button.config(text=self._text("income"))
        self.expense_button.config(text=self._text("expense"))
        self.table_button.config(text=self._text("show_table"))
        self.reset_button.config(text=self._text("reset", currency=self.currency))
        self.update_display()

    def update_display(self):
        balance = logic.get_balance(self.currency)
        today_total = logic.get_today_total(self.currency)
        self.balance_label.config(text=self._format_amount(balance))
        self.today_label.config(text=self._text("today", amount=self._format_amount(today_total)))

    def open_excel(self):
        """Open the workbook in the system's default spreadsheet application."""
        try:
            file_path = excel_storage.get_file_path()
            subprocess.Popen(["xdg-open", file_path])
        except OSError as error:
            messagebox.showerror(
                self._text("open_error_title"), self._text("open_error", error=error)
            )

    def add_transaction(self, transaction_type):
        text = self.amount_entry.get().strip()
        try:
            if not text:
                raise ValueError(self._text("enter_amount"))
            try:
                amount = Decimal(text.replace(",", "."))
            except InvalidOperation:
                raise ValueError(self._text("whole_number"))

            if not amount.is_finite():
                raise ValueError(self._text("whole_number"))
            if amount <= 0:
                raise ValueError(self._text("positive_amount"))
            if self.currency == "UZS" and amount != amount.to_integral_value():
                raise ValueError(self._text("uzs_whole_number"))
            if self.currency == "USD" and amount != amount.quantize(Decimal("0.01")):
                raise ValueError(self._text("usd_precision"))

            logic.add_transaction(amount, transaction_type, self.currency)
            self.update_display()
            self.amount_entry.delete(0, tk.END)
            self.amount_entry.focus_set()
        except ValueError as error:
            messagebox.showerror(self._text("input_error"), str(error))
        except Exception as error:
            messagebox.showerror(
                self._text("error"), self._text("transaction_error", error=error)
            )

    def reset_balance(self):
        currency_name = self._account_name()
        answer = messagebox.askyesno(
            self._text("reset_title"),
            self._text("reset_question", currency=currency_name),
        )
        if not answer:
            return

        try:
            excel_storage.reset_data(self.currency)
            self.update_display()
            messagebox.showinfo(
                self._text("reset_done_title"),
                self._text("reset_done", currency=currency_name),
            )
        except OSError as error:
            messagebox.showerror(
                self._text("reset_error_title"), self._text("reset_error", error=error)
            )

    def _account_name(self):
        key = "usd_account" if self.currency == "USD" else "uzs_account"
        return self._text(key)


def run_app():
    root = tk.Tk()
    FinanceApp(root)
    root.mainloop()
