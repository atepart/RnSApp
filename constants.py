from enum import EnumMeta, Enum


class TableColumnsMeta(EnumMeta):
    def __new__(mcs, cls, bases, classdict):
        enum_class = super().__new__(mcs, cls, bases, classdict)
        for index, member in enumerate(enum_class):
            member._index = index
        return enum_class


class TableColumns(Enum, metaclass=TableColumnsMeta):
    def __init__(self, name, dtype, slug=None):
        self._name = name
        self._dtype = dtype
        self._slug = slug

    @property
    def name(self):
        return self._name

    @property
    def dtype(self):
        return self._dtype

    @property
    def index(self):
        return self._index

    @property
    def slug(self):
        return self._slug

    @classmethod
    def get_all_names(cls):
        return [member.name for member in cls]

    @classmethod
    def get_by_index(cls, index: int):
        for i, item in enumerate(cls):
            if i == index:
                return item


class DataTableColumns(TableColumns, metaclass=TableColumnsMeta):
    NUMBER = ("№", int)
    NAME = ("Имя", str)
    SELECT = ("✓", bool)
    DIAMETER = ("Диаметр ACAD (μm)", float)
    RESISTANCE = ("Rn (Ω)", float)
    RNS = ("RnS", float)
    RNS_ERROR = ("Ошибка RnS", float)
    DRIFT = ("Суммарный Уход (μm)", float)
    SQUARE = ("Площадь (μm^2)", float)
    RN_SQRT = ("Rn^-0.5", float)


class ParamTableColumns(TableColumns, metaclass=TableColumnsMeta):
    SLOPE = ("Наклон", float)
    INTERCEPT = ("Пересечение", float)
    DRIFT = ("Суммарный Уход", float)
    RNS = ("RnS", float)
    DRIFT_ERROR = ("Ошибка ухода", float)
    RNS_ERROR = ("Ошибка RnS", float)


class MetaTableColumns(TableColumns, metaclass=TableColumnsMeta):
    RN_CONSISTENT = ("Последовательное Rn", float, "rn_consistent")
    ALLOWED_ERROR = ("Разрешенная ошибка", float, "allowed_error")


PLOT_COLORS = [
    "#3357FF",  # Ярко-синий
    "#FF3333",  # Ярко-красный
    "#33FF57",  # Ярко-зеленый
    "#FF33A1",  # Ярко-розовый
    "#00FFFF",  # Циановый
    "#FFFF33",  # Ярко-желтый
    "#333333",  # Серый
    "#33A1FF",  # Ярко-небесно-синий
    "#F34949",  # Ярко-красно-розовый
    "#6AF16C",  # Ярко-светло-зелёный
    "#E66BAD",  # Ярко-светло-розовый
    "#8CE8E8",  # Тускло-светло-голубой
    "#E0E05C",  # Тускло-светло-жёлтый
    "#666666",  # Светло-серый
    "#05FFA2",  # Ярко-аквамариновый
    "#FF5B2E",  # Ярко-рыже-красный
]


RNS_ERROR_COLOR = "#F56B6B"
WHITE = "#FFFFFF"
BLACK = "#000000"
