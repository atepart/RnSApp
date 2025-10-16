from typing import Any, Dict, Iterable, List, Optional, Protocol, Tuple

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


class CellDataIO(Protocol):
    """Port for importing/exporting cell data to external storage (e.g., XLSX).

    Clean-architecture friendly interface to decouple UI from persistence details.
    """

    def save(
        self,
        file_name: str,
        cell_grid_values: List[Tuple[str, str, str]],
        repo: CellRepository,
    ) -> None:
        ...

    def load(self, file_name: str) -> Tuple[List[Dict[str, Any]], List[str]]:
        """Load items and return (items, errors).

        items are dicts accepted by CellRepository.update_or_create_item.
        """
        ...
