# test_math_operations.py (test file)
import pytest
from numpy.ma.core import divide
from math_operations import add, subtract


@pytest.mark.add
def test_add():
    assert add(2, 3) == 5  # Test passes
    assert add(-1, 1) == 0  # Test passes

@pytest.mark.sub
def test_subtract():
    assert subtract(5, 3) == 2  # Test passes
    assert subtract(5, 5) == 0  # Test passes
    assert subtract(5, 7) == -2  # Test passes

@pytest.mark.div
def test_mul():
    assert divide(10, 2) == 9


@pytest.mark.example
def test_example(setup_and_teardown):
    print("Test executed")
    assert True
#
#
# @pytest.mark.array
# @pytest.mark.parametrize("num1,num2,num3",[(1,2,3),(4,5,6)])
# def test_arry(num1,num2,num3):
#     print(num1,num2)


def test_given(getData):
    print("vijaya")