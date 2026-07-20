import unittest

import numpy as np

from domain.errors import ListsNotSameLength
from domain.utils import inverse_diameter_linear_fit, linear_fit, weighted_linear_fit


class WeightedLinearFitTests(unittest.TestCase):
    def test_matches_direct_weighted_normal_equations(self):
        x = np.array([1.0, 2.0, 4.0, 8.0])
        y = np.array([1.1, 2.2, 3.7, 9.0])
        residual_weights = np.array([1.0, 0.5, 0.25, 0.125])

        slope, intercept = weighted_linear_fit(x, y, residual_weights)

        design = np.column_stack((x, np.ones_like(x)))
        weighted_design = design * residual_weights[:, None]
        weighted_y = y * residual_weights
        expected_slope, expected_intercept = np.linalg.lstsq(weighted_design, weighted_y, rcond=None)[0]
        self.assertAlmostEqual(slope, expected_slope, places=12)
        self.assertAlmostEqual(intercept, expected_intercept, places=12)

    def test_equal_weights_match_unweighted_fit(self):
        x = [1.0, 2.0, 3.0, 5.0]
        y = [0.5, 2.5, 2.8, 7.0]
        assert np.allclose(weighted_linear_fit(x, y, [1.0] * len(x)), linear_fit(x, y), rtol=0, atol=1e-12)

    def test_common_weight_scale_does_not_change_fit(self):
        x = [1.0, 2.0, 4.0, 8.0]
        y = [1.0, 2.1, 3.9, 9.0]
        weights = np.array([1.0, 0.5, 0.25, 0.125])
        assert np.allclose(
            weighted_linear_fit(x, y, weights), weighted_linear_fit(x, y, weights * 17.0), rtol=0, atol=1e-12
        )

    def test_inverse_diameter_fit_uses_one_over_d_as_residual_multiplier(self):
        diameters = np.array([1.0, 2.0, 4.0, 8.0])
        y = np.array([1.0, 2.0, 4.0, 12.0])
        expected = weighted_linear_fit(diameters, y, 1.0 / diameters)
        actual = inverse_diameter_linear_fit(diameters, y)
        assert np.allclose(actual, expected, rtol=0, atol=1e-12)
        assert not np.allclose(actual, linear_fit(diameters, y), rtol=0, atol=1e-06)

    def test_invalid_diameters_are_excluded(self):
        actual = inverse_diameter_linear_fit([0.0, 1.0, -2.0, 2.0, np.nan], [99.0, 2.0, 88.0, 4.0, 77.0])
        assert np.allclose(actual, (2.0, 0.0), rtol=0, atol=1e-12)

    def test_rejects_mismatched_lengths_and_degenerate_x(self):
        with self.assertRaises(ListsNotSameLength):
            weighted_linear_fit([1.0, 2.0], [1.0], [1.0, 1.0])
        with self.assertRaises(ValueError):
            inverse_diameter_linear_fit([1.0, 1.0], [2.0, 3.0])


if __name__ == "__main__":
    unittest.main()
