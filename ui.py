
import os
import excel_storage
import tkinter as tk
from tkinter import messagebox

import logic


class FinanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Учёт финансов")
        self.root.geometry("400x350")
        self.root.resizable(False, False)

        tk.Label(
            root,
            text="Учёт финансов",
            font=("Arial", 18, "bold")
        ).pack(pady=(20, 10))

        tk.Label(
            root,
            text="Баланс за всё время",
            font=("Arial", 11)
        ).pack()

        self.balance_label = tk.Label(
            root,
            text="0 UZS",
            font=("Arial", 20, "bold"),
            fg="#18794e"
        )
        self.balance_label.pack(pady=5)

        self.today_label = tk.Label(
            root,
            text="Итог за сегодня: 0 UZS",
            font=("Arial", 11)
        )
        self.today_label.pack(pady=5)

        tk.Label(
            root,
            text="Сумма операции (UZS)"
        ).pack(pady=(15, 5))

        self.amount_entry = tk.Entry(
            root,
            font=("Arial", 14),
            justify="center"
        )
        self.amount_entry.pack(padx=40, fill="x")
        self.amount_entry.focus_set()

        buttons = tk.Frame(root)
        buttons.pack(pady=20)

        tk.Button(
            buttons,
            text="+ Доход",
            width=12,
            bg="#18794e",
            fg="white",
            command=lambda: self.add_transaction("income")
        ).pack(side="left", padx=5)

        tk.Button(
            buttons,
            text="- Расход",
            width=12,
            bg="#b42318",
            fg="white",
            command=lambda: self.add_transaction("expense")
        ).pack(side="left", padx=5)


        tk.Button(
            root,
            text="Показать в таблице",
            command=self.open_excel
        ).pack(pady=(0, 10))


        self.update_display()

    def update_display(self):
        balance = logic.get_balance()
        today_total = logic.get_today_total()

        self.balance_label.config(
            text=f"{balance:,} UZS".replace(",", " ")
        )

        self.today_label.config(
            text=f"Итог за сегодня: {today_total:,} UZS".replace(",", " ")
        )

    #Функция открытия Excel файла, который хранит всю историю финансовых операций

    def open_excel(self):
        """Открывает Excel-файл в стандартной программе."""

        try:
            file_path = excel_storage.get_file_path()
            os.startfile(file_path)

        except OSError as error:
            messagebox.showerror(
                "Ошибка открытия",
                f"Не удалось открыть Excel-файл:\n{error}"
            )

    def add_transaction(self, transaction_type):
        text = self.amount_entry.get().strip()

        try:
            amount = int(text)

            result = logic.add_transaction(
                amount,
                transaction_type
            )

            self.update_display()
            self.amount_entry.delete(0, tk.END)
            self.amount_entry.focus_set()

        except ValueError as error:
            messagebox.showerror("Ошибка ввода", str(error))

        except Exception as error:
            messagebox.showerror(
                "Ошибка",
                f"Не удалось обработать операцию:\n{error}"
            )


def run_app():
    root = tk.Tk()
    FinanceApp(root)
    root.mainloop()

#Буду честен, я и сам не совсем понимаю что здесь происходит, но вроде как это должно работать