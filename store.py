from dataclasses import dataclass
from typing import List, Optional


@dataclass
class InitialDataItem:
    value: str
    row: int
    col: int

    @property
    def dict(self):
        return self.__dict__

    def __getitem__(self, item):
        return getattr(self, item)


class Item:
    def __init__(
        self,
        cell: int,
        name: str,
        # Исходные данные
        diameter_list: List[float],
        rn_sqrt_list: List[float],
        # Расчетные данные
        slope: float,
        intercept: float,
        drift: float,
        rns: float,
        drift_error: float,
        rns_error: float,
        # Исходная таблица
        initial_data: List[InitialDataItem],
    ):
        self.cell = cell
        self.name = name

        self.diameter = diameter_list
        self.rn_sqrt = rn_sqrt_list

        self.slope = slope
        self.intercept = intercept
        self.drift = drift
        self.rns = rns
        self.drift_error = drift_error
        self.rns_error = rns_error

        self.initial_data = initial_data
        self.is_plot = False


class ItemsList(list):
    def _filter(self, **kwargs) -> filter:
        def _filter(item):
            for key, value in kwargs.items():
                if not getattr(item, key, None) == value:
                    return False
            return True

        return filter(_filter, self)

    def _exclude(self, **kwargs) -> filter:
        def _exclude(item):
            for key, value in kwargs.items():
                if not getattr(item, key, None) != value:
                    return False
            return True

        return filter(_exclude, self)

    def filter(self, **kwargs) -> "ItemsList":
        return self.__class__(self._filter(**kwargs))

    def exclude(self, **kwargs):
        return self.__class__(self._exclude(**kwargs))

    def get(self, **kwargs) -> Optional["Item"]:
        try:
            return self.filter(**kwargs)[0]
        except IndexError:
            return None

    def exists(self):
        return len(self)


class Store:
    data: ItemsList = ItemsList()

    @classmethod
    def update_or_create_item(cls, cell: int, **kwargs):
        item = cls.data.get(cell=cell)
        if item:
            for k, v in kwargs.items():
                setattr(item, k, v)
        else:
            item = Item(cell, **kwargs)
            cls.data.append(item)
