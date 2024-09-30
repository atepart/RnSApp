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
    DRIFT = ("Уход", float)
    RNS = ("RnS", float)
    DIAMETER = ("Диаметр ACAD (μm)", float)
    RESISTANCE = ("Сопротивление (Ω)", float)
    RN = ("Rn^-0.5", float)


class ParamTableColumns(TableColumns, metaclass=TableColumnsMeta):
    SLOPE = ("Наклон", float)
    INTERCEPT = ("Пересечение", float)
    DRIFT = ("Уход", float)
    RNS = ("RnS", float)
    DRIFT_ERROR = ("Ошибка ухода", float)
    RNS_ERROR = ("Ошибка RnS", float)
