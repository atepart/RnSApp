from typing import List, Optional


class Item:
    def __init__(
        self,
        cell: int,
        diameter: List[float],
        rn_sqrt: List[float],
        drift: float,
        slope: float,
        intercept: float,
    ):
        self.cell = cell
        self.diameter = diameter
        self.rn_sqrt = rn_sqrt
        self.drift = drift
        self.slope = slope
        self.intercept = intercept
        self.is_plot = False


class ItemsList(list):
    def _filter(self, **kwargs) -> filter:
        def _filter(item):
            for key, value in kwargs.items():
                if not getattr(item, key, None) == value:
                    return False
            return True

        return filter(_filter, self)

    def filter(self, **kwargs) -> "ItemsList":
        return self.__class__(self._filter(**kwargs))

    def get(self, **kwargs) -> Optional["Item"]:
        try:
            return self.filter(**kwargs)[0]
        except IndexError:
            return None


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
