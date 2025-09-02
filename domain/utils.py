import numpy as np

from .errors import ListsNotSameLength


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


def linear(x: float, b: float, a: float):
    return x * b + a


def drop_nans(arr1: list, arr2: list):
    if len(arr1) != len(arr2):
        raise ListsNotSameLength
    return np.array([arr for arr in np.array([arr1, arr2]).T if all(arr)], dtype=float).T


def calculate_drift(slope: float, intercept: float):
    return -intercept / slope


def calculate_drift_per_sample(diameter: float, resistance: float, rns: float, rn_persistent: float):
    return diameter - np.sqrt((4 * rns / (resistance + rn_persistent)) / np.pi)


def calculate_rns(slope: float):
    return np.pi * 0.25 / (slope**2)


def calculate_rns_per_sample(resistance: float, diameter: float, drift: float, rn_persistent: float):
    return (resistance + rn_persistent) * 0.25 * np.pi * (diameter - drift) ** 2


def calculate_rn_sqrt(resistance: float, rn_consistent: float):
    return 1 / np.sqrt(resistance + rn_consistent)


def calculate_square(diameter: float, drift: float):
    return (diameter - drift) ** 2 * np.pi / 4


def calculate_rns_error_per_sample(rns_i: float, rns: float):
    return np.abs(rns_i - rns)


def calculate_allowed_rns_error(rns_error: float, allowed_error: float):
    return rns_error * (1 + allowed_error)


def calculate_rns_error_diff(rns_error_per_sample: float, rns_error: float, allowed_error: float):
    return rns_error_per_sample - calculate_allowed_rns_error(rns_error=rns_error, allowed_error=allowed_error)
