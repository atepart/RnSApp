from enum import EnumMeta, Enum


class TableColumnsMeta(EnumMeta):
    def __new__(mcs, cls, bases, classdict):
        enum_class = super().__new__(mcs, cls, bases, classdict)
        for index, member in enumerate(enum_class):
            member._index = index
        return enum_class


class TableColumns(Enum, metaclass=TableColumnsMeta):
    def __init__(self, name, dtype):
        self._name = name
        self._dtype = dtype

    @property
    def name(self):
        return self._name

    @property
    def dtype(self):
        return self._dtype

    @property
    def index(self):
        return self._index

    @classmethod
    def get_all_names(cls):
        return [member.name for member in cls]


class DataTableColumns(TableColumns, metaclass=TableColumnsMeta):
    NUMBER = ("№", int)
    NAME = ("Имя", str)
    DIAMETER = ("Диаметр ACAD (μm)", float)
    RESISTANCE = ("Rn (Ω)", float)
    RNS = ("RnS", float)
    DRIFT = ("Суммарный Уход (μm)", float)
    SQUARE = ("Площадь (μm^2)", float)
    RN = ("Rn^-0.5", float)


class ParamTableColumns(TableColumns, metaclass=TableColumnsMeta):
    SLOPE = ("Наклон", float)
    INTERCEPT = ("Пересечение", float)
    DRIFT = ("Суммарный Уход", float)
    RNS = ("RnS", float)
    DRIFT_ERROR = ("Ошибка ухода", float)
    RNS_ERROR = ("Ошибка RnS", float)


PLOT_COLORS = [
    "#FF5733",  # Ярко-оранжевый
    "#33FF57",  # Ярко-зеленый
    "#3357FF",  # Ярко-синий
    "#FF33A1",  # Ярко-розовый
    "#A133FF",  # Ярко-фиолетовый
    "#FFFF33",  # Ярко-желтый
    "#33FFFF",  # Ярко-бирюзовый
    "#FF3333",  # Ярко-красный
    "#33FFA1",  # Ярко-голубой
    "#A1FF33",  # Ярко-лаймовый
    "#FFA133",  # Ярко-апельсиновый
    "#33A1FF",  # Ярко-небесно-синий
    "#A133A1",  # Ярко-пурпурный
    "#33A133",  # Ярко-темно-зеленый
    "#A1A133",  # Ярко-оливковый
    "#33A1A1",  # Ярко-бирюзово-зеленый
]
