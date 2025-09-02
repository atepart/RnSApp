from enum import Enum, EnumMeta


class TableColumnsMeta(EnumMeta):
    def __new__(mcs, cls, bases, classdict):
        enum_class = super().__new__(mcs, cls, bases, classdict)
        for index, member in enumerate(enum_class):
            member._index = index
        return enum_class


class TableColumns(Enum, metaclass=TableColumnsMeta):
    def __init__(self, name, dtype, slug=None) -> None:
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
    def get_all_slugs(cls):
        return [member.slug for member in cls]

    @classmethod
    def get_by_index(cls, index: int):
        for i, item in enumerate(cls):
            if i == index:
                return item
        return None


class DataTableColumns(TableColumns, metaclass=TableColumnsMeta):
    NUMBER = ("№", int, "№")
    NAME = ("Имя", str, "Имя")
    SELECT = ("", bool, "✓")
    DIAMETER = ("Диаметр ACAD (μm)", float, "Диаметр ACAD (μm)")
    RESISTANCE = ("Rn (Ω)", float, "Rn (Ω)")
    RNS = ("RnS", float, "RnS")
    RNS_ERROR = ("Ошибка RnS", float, "Ошибка RnS")
    DRIFT = ("Суммарный Уход (μm)", float, "Суммарный Уход (μm)")
    SQUARE = ("Площадь (μm^2)", float, "Площадь (μm^2)")
    RN_SQRT = ("Rn^-0.5", float, "Rn^-0.5")


class ParamTableColumns(TableColumns, metaclass=TableColumnsMeta):
    SLOPE = ("Наклон", float, "slope")
    INTERCEPT = ("Пересечение", float, "intercept")
    DRIFT = ("Суммарный Уход", float, "drift")
    RNS = ("RnS", float, "rns")
    DRIFT_ERROR = ("Ошибка ухода", float, "drift_error")
    RNS_ERROR = ("Ошибка RnS", float, "rns_error")
    RN_CONSISTENT = ("Последовательное Rn", float, "rn_consistent")
    ALLOWED_ERROR = ("Разрешенная ошибка", float, "allowed_error")


PLOT_COLORS = [
    "#3357FF",
    "#FF3333",
    "#33FF57",
    "#FF33A1",
    "#00FFFF",
    "#FFFF33",
    "#333333",
    "#33A1FF",
    "#F34949",
    "#6AF16C",
    "#E66BAD",
    "#8CE8E8",
    "#E0E05C",
    "#666666",
    "#05FFA2",
    "#FF5B2E",
]


RNS_ERROR_COLOR = "#F56B6B"
WHITE = "#FFFFFF"
BLACK = "#000000"
BLUE = "#a2d2ff"
