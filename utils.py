import numpy as np


def linear_fit(x, y):
    def mean(xs):
        return sum(xs) / len(xs)

    m_x = mean(x)
    m_y = mean(y)

    def std(xs, m):
        normalizer = len(xs) - 1
        return np.sqrt(sum((pow(x1 - m, 2) for x1 in xs)) / normalizer)

    def pearson_r(xs, ys):
        sum_xy = 0
        sum_sq_v_x = 0
        sum_sq_v_y = 0

        for x1, y2 in zip(xs, ys):
            var_x = x1 - m_x
            var_y = y2 - m_y
            sum_xy += var_x * var_y
            sum_sq_v_x += pow(var_x, 2)
            sum_sq_v_y += pow(var_y, 2)
        return sum_xy / np.sqrt(sum_sq_v_x * sum_sq_v_y)

    r = pearson_r(x, y)

    slope = r * (std(y, m_y) / std(x, m_x))
    intercept = m_y - slope * m_x

    return slope, intercept


def linear(x, b, a):
    return x * b + a


def calculate_drift(diameter: float, resistance: float, rns: float):
    """Рассчет ухода для образца"""
    return diameter - np.sqrt((4 * rns / resistance) / np.pi)


def calculate_rns(slope: float):
    """Рассчет RnS для набора образцов"""
    return np.pi * 0.25 / (slope**2)


def calculate_rns_per_sample(resistance: float, diameter: float, zero_x: float):
    """Рассчет RnS для одного образца"""
    return resistance * 0.25 * np.pi * (diameter - zero_x) ** 2
