import datetime as dt

from main import DATE_FORMAT, HEIGHT


class Accrual:
    def __init__(self, year: int, accrued: int, payment_amount: int, due_date: str, debt: int,):
        self.year = year  # год начислений
        self.accrued = accrued  # сумма начислений
        self.payment_amount = payment_amount  # сумма оплаты
        self.due_date = due_date  # срок оплаты
        self.debt = debt  # задолженность

    def __str__(self):
        return f'Accrual: year = {self.year}, accrued = {self.accrued}, due_date = {self.due_date}'

    def __repr__(self):
        return f'Accrual: year = {self.year}, accrued = {self.accrued}, due_date = {self.due_date}'


class Box:
    def __init__(self, number: int):
        self.number = number
        self.accruals = []

    def __str__(self):
        return f'Class box №{self.number}'

    def __repr__(self):
        return f'Class box №{self.number}'


class People:
    def __init__(self, name, address: str, phone: str):
        self.name = name
        self.address = address
        self.boxes = []
        self.phone = phone

    def get_height(self, height_bottom_line: int) -> int:
        """Вернуть высоту, которую будет занимать человек в квитанции."""
        result = 10 * HEIGHT  # верх квитанции
        for box in self.boxes:
            result += HEIGHT * 3  # заголовок таблицы взносов
            for _ in box.accruals:
                result += HEIGHT
            result += HEIGHT  # пропуск после таблицы взносов
        result += HEIGHT * 3 + height_bottom_line  # низ квитанции
        return result

    def get_sum_accrued(self) -> str:
        """Вернуть сумму начислений."""
        result = 0
        for box in self.boxes:
            for accrual in box.accruals:
                result += accrual.accrued
        return str(result)

    def get_sum_debt(self) -> int:
        """Вернуть суммарную задолженность."""
        result = 0
        for box in self.boxes:
            for accrual in box.accruals:
                result += accrual.debt
        return result
