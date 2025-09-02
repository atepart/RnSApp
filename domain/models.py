from dataclasses import dataclass
from typing import Any, Optional


class BaseList(list):
    def _filter(self, **kwargs) -> filter:
        def _filter(item):
            for key, value in kwargs.items():
                if getattr(item, key, None) != value:
                    return False
            return True

        return filter(_filter, self)

    def _exclude(self, **kwargs) -> filter:
        def _exclude(item):
            for key, value in kwargs.items():
                if getattr(item, key, None) == value:
                    return False
            return True

        return filter(_exclude, self)

    def filter(self, **kwargs) -> "BaseList":
        return self.__class__(self._filter(**kwargs))

    def exclude(self, **kwargs):
        return self.__class__(self._exclude(**kwargs))

    def get(self, **kwargs) -> Any:
        try:
            return self.filter(**kwargs)[0]
        except IndexError:
            return None

    def exists(self):
        return len(self)


@dataclass
class InitialDataItem:
    value: Any
    row: int
    col: int

    @property
    def dict(self):
        return self.__dict__

    def __getitem__(self, item):
        return getattr(self, item)


class InitialDataItemList(BaseList):
    ...


class Item:
    def __init__(
        self,
        cell: int,
        name: str,
        diameter_list,
        rn_sqrt_list,
        slope: float,
        intercept: float,
        drift: float,
        rns: float,
        drift_error: float,
        rns_error: float,
        initial_data: InitialDataItemList,
        rn_consistent: float = 0,
        allowed_error: float = 0,
    ) -> None:
        self.cell = cell
        self.name = name
        self.diameter_list = diameter_list
        self.rn_sqrt_list = rn_sqrt_list
        self.slope = slope
        self.intercept = intercept
        self.drift = drift
        self.rns = rns
        self.drift_error = drift_error
        self.rns_error = rns_error
        self.rn_consistent = rn_consistent
        self.allowed_error = allowed_error
        self.initial_data = initial_data
        self.is_plot = False


class ItemsList(BaseList):
    def get(self, **kwargs) -> Optional["Item"]:
        return super().get(**kwargs)
