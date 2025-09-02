from typing import Iterable, Optional

from domain.models import Item, ItemsList
from domain.ports import CellRepository


class InMemoryCellRepository(CellRepository):
    def __init__(self) -> None:
        self._data: ItemsList[Item] = ItemsList()

    def update_or_create_item(self, cell: int, **kwargs) -> Item:
        item = self._data.get(cell=cell)
        if item:
            for k, v in kwargs.items():
                setattr(item, k, v)
        else:
            item = Item(cell, **kwargs)
            self._data.append(item)
        return item

    def get(self, **kwargs) -> Optional[Item]:
        return self._data.get(**kwargs)

    def clear(self) -> None:
        self._data = ItemsList()

    def __iter__(self) -> Iterable[Item]:
        return iter(self._data)
