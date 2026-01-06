import math
import pytest

from tt1 import power, exp, log


def test_power_basic():
    assert power(2, 3) == 8
    assert power(5, 0) == 1
    assert pytest.approx(power(4, 0.5)) == 2


def test_power_negative_exponent():
    assert pytest.approx(power(2, -1)) == 0.5


def test_exp():
    assert pytest.approx(exp(0)) == 1
    assert pytest.approx(exp(1), rel=1e-12) == math.e


def test_log_natural():
    assert pytest.approx(log(math.e), rel=1e-12) == 1
    assert pytest.approx(log(1)) == 0


def test_log_base_10():
    assert pytest.approx(log(100, 10)) == 2
    assert pytest.approx(log(1000, 10)) == 3


def test_log_custom_base():
    assert pytest.approx(log(8, 2)) == 3
    assert pytest.approx(log(27, 3)) == 3


def test_log_errors():
    with pytest.raises(ValueError):
        log(0)
    with pytest.raises(ValueError):
        log(-1)
    with pytest.raises(ValueError):
        log(10, 1)
    with pytest.raises(ValueError):
        log(10, -2)
