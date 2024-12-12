# test_math_operations.py (test file)
import pytest
from numpy.ma.core import divide
from math_operations import add, subtract


@pytest.mark.skip(reason='my wish')
def test_add():
    assert add(2, 3) == 5  # Test passes
    assert add(-1, 1) == 0  # Test passes

@pytest.mark.sub
def test_subtract():
    assert subtract(5, 3) == 2  # Test passes
    assert subtract(5, 5) == 0  # Test passes
    assert subtract(5, 7) == -2  # Test passes
    print("true")

@pytest.mark.div
def test_mul():
    assert divide(10, 2) == 5


