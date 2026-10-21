
from datetime import date

import excel_storage


def get_today():
    """Возвращает сегодняшнюю дату в формате YYYY-MM-DD."""

    return date.today().isoformat()


def get_balance():
    """Рассчитывает общий баланс за всё время."""

    history = excel_storage.load_history()

    balance = sum(
        record["daily_total"]
        for record in history
    )

    return balance


def get_today_total():
    """Возвращает текущий итог за сегодня."""

    today = get_today()

    return excel_storage.get_daily_total(today)


def add_transaction(amount, transaction_type):
    """
    Изменяет дневной итог и сохраняет его в Excel.

    amount — сумма операции.
    transaction_type — 'income' для дохода
                       или 'expense' для расхода.
    """

    if amount <= 0:
        raise ValueError("Сумма должна быть больше нуля.")

    if transaction_type not in ("income", "expense"):
        raise ValueError("Неизвестный тип операции.")

    today = get_today()

    # Получаем текущий итог за сегодня
    daily_total = excel_storage.get_daily_total(today)

    # Добавляем доход или вычитаем расход
    if transaction_type == "income":
        daily_total += amount
    else:
        daily_total -= amount

    # Сохраняем обновлённый итог
    excel_storage.save_daily_total(today, daily_total)

    # Возвращаем обновлённые данные для интерфейса
    return {
        "daily_total": daily_total,
        "balance": get_balance()
    }
#Мне начинает это нравится