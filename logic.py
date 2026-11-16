
from datetime import date
from decimal import Decimal

import excel_storage


def get_today():
    """Возвращает сегодняшнюю дату в формате YYYY-MM-DD."""

    return date.today().isoformat()


def get_balance(currency="UZS"):
    """Calculate the all-time balance for one currency account."""

    history = excel_storage.load_history()

    balance = sum(
        record["daily_total"]
        for record in history
        if record["currency"] == currency
    )

    return balance


def get_today_total(currency="UZS"):
    """Return today's total for one currency account."""

    today = get_today()

    return excel_storage.get_daily_total(today, currency)


def add_transaction(amount, transaction_type, currency="UZS"):
    """
    Изменяет дневной итог и сохраняет его в Excel.

    amount — сумма операции.
    transaction_type — 'income' для дохода
                       или 'expense' для расхода.
    """

    try:
        amount = Decimal(str(amount))
    except Exception as error:
        raise ValueError("Сумма должна быть числом.") from error

    if not amount.is_finite():
        raise ValueError("Сумма должна быть конечным числом.")

    if amount <= 0:
        raise ValueError("Сумма должна быть больше нуля.")

    if transaction_type not in ("income", "expense"):
        raise ValueError("Неизвестный тип операции.")

    # Получаем текущий итог за сегодня
    if currency not in ("UZS", "USD"):
        raise ValueError("Неизвестная валюта.")

    if currency == "UZS" and amount != amount.to_integral_value():
        raise ValueError("Для UZS сумма должна быть целым числом.")
    if currency == "USD" and amount != amount.quantize(Decimal("0.01")):
        raise ValueError("Для USD допускается не более двух знаков после запятой.")

    today = get_today()
    daily_total = excel_storage.get_daily_total(today, currency)

    # Добавляем доход или вычитаем расход
    if transaction_type == "income":
        daily_total += amount
    else:
        daily_total -= amount

    # Сохраняем обновлённый итог
    excel_storage.save_daily_total(today, daily_total, currency)

    # Возвращаем обновлённые данные для интерфейса
    return {
        "daily_total": daily_total,
        "balance": get_balance(currency)
    }
#Мне начинает это нравится
