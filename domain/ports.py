from typing import Iterable, Optional, Protocol

from .models import Item


class CellRepository(Protocol):
    def update_or_create_item(self, cell: int, **kwargs) -> Item:
        ...

    def get(self, **kwargs) -> Optional[Item]:
        ...

    def clear(self) -> None:
        ...

    def __iter__(self) -> Iterable[Item]:
        ...
