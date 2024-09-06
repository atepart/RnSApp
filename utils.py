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

    b = r * (std(y, m_y) / std(x, m_x))
    a = m_y - b * m_x

    return b, a

def linear(x, b, a):
    return x * b + a