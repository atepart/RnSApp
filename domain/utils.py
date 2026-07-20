import numpy as np

from .errors import ListsNotSameLength


def weighted_linear_fit(x, y, residual_weights):
    """Fit ``y = slope*x + intercept`` using residual multipliers.

    ``residual_weights`` are the multipliers applied to residuals, therefore
    the minimized objective is ``sum((weight * residual) ** 2)``.  This is the
    same convention used by ``numpy.polyfit(..., w=...)``.
    """

    x_values = np.asarray(x, dtype=float)
    y_values = np.asarray(y, dtype=float)
    weights = np.asarray(residual_weights, dtype=float)
    if x_values.ndim != 1 or y_values.ndim != 1 or weights.ndim != 1:
        raise ValueError("Fit inputs must be one-dimensional")
    if not (len(x_values) == len(y_values) == len(weights)):
        raise ListsNotSameLength

    finite = np.isfinite(x_values) & np.isfinite(y_values) & np.isfinite(weights) & (weights > 0)
    x_values = x_values[finite]
    y_values = y_values[finite]
    weights = weights[finite]
    if len(x_values) < 2 or np.unique(x_values).size < 2:
        raise ValueError("At least two points with distinct x values are required")

    objective_weights = weights**2
    total_weight = np.sum(objective_weights)
    x_mean = np.sum(objective_weights * x_values) / total_weight
    y_mean = np.sum(objective_weights * y_values) / total_weight
    centered_x = x_values - x_mean
    denominator = np.sum(objective_weights * centered_x**2)
    if denominator == 0:
        raise ValueError("Cannot fit data with zero weighted x variance")

    slope = np.sum(objective_weights * centered_x * (y_values - y_mean)) / denominator
    intercept = y_mean - slope * x_mean
    return float(slope), float(intercept)


def linear_fit(x, y):
    """Unweighted least-squares fit kept for callers that explicitly need OLS."""

    return weighted_linear_fit(x, y, np.ones(len(x), dtype=float))


def inverse_diameter_linear_fit(diameters, y):
    """Fit using ``1 / diameter`` residual multipliers.

    Consequently the coefficient in the weighted least-squares sum is
    ``1 / diameter**2``. Empty, non-finite and non-positive diameters are not
    valid measurements for this fit and are excluded together with their y.
    """

    diameter_values = np.asarray(diameters, dtype=float)
    y_values = np.asarray(y, dtype=float)
    if diameter_values.ndim != 1 or y_values.ndim != 1:
        raise ValueError("Fit inputs must be one-dimensional")
    if len(diameter_values) != len(y_values):
        raise ListsNotSameLength
    valid = np.isfinite(diameter_values) & np.isfinite(y_values) & (diameter_values > 0)
    diameter_values = diameter_values[valid]
    y_values = y_values[valid]
    return weighted_linear_fit(diameter_values, y_values, 1.0 / diameter_values)


def linear(x: float, b: float, a: float):
    return x * b + a


def drop_nans(arr1: list, arr2: list):
    """Return two float arrays with rows where both inputs are truthy; never fail unpacking.

    If no valid pairs, returns two empty arrays with shape (0,).
    """
    if len(arr1) != len(arr2):
        raise ListsNotSameLength
    pairs = np.array([arr1, arr2], dtype=object).T
    kept = [pair for pair in pairs if all(pair)]
    if not kept:
        return np.array([], dtype=float), np.array([], dtype=float)
    kept_arr = np.array(kept, dtype=float).T
    return kept_arr[0], kept_arr[1]


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


def calculate_diameter_from_area(area: float):
    return np.sqrt(4 * area / np.pi)


def calculate_area_from_diameter(diameter: float):
    return diameter**2 * np.pi / 4


def calculate_rns_over_rn(rns: float, resistance: float):
    return rns / resistance


def calculate_rns_error_per_sample(rns_i: float, rns: float):
    return np.abs(rns_i - rns)


def calculate_allowed_rns_error(rns_error: float, allowed_error: float):
    return rns_error * (1 + allowed_error)


def calculate_rns_error_diff(rns_error_per_sample: float, rns_error: float, allowed_error: float):
    return rns_error_per_sample - calculate_allowed_rns_error(rns_error=rns_error, allowed_error=allowed_error)


def calculate_real_area(area_nominal: float, drift: float) -> float:
    """Calculate real circular area for a nominal area with given drift.

    A_real = pi * (sqrt(4*A_nom/pi) - drift)^2 / 4
    """
    d0 = np.sqrt(4 * area_nominal / np.pi)
    return np.pi * (d0 - drift) ** 2 / 4


def calculate_real_custom_area(area_nominal: float | None, planned_drift: float, drift: float) -> float:
    """Calculate real area for custom S* with planned drift d and actual drift l.

    D* = sqrt(4/π * S*); S_real = π/4 * (D* - l + d)^2
    """
    with np.errstate(all="ignore"):
        if area_nominal in (None, "", 0):
            return 0.0
        try:
            d_star = float(np.sqrt(4 * float(area_nominal) / np.pi))
            return float(np.pi * (d_star - float(drift) + float(planned_drift)) ** 2 / 4)
        except Exception:
            return 0.0
