import pytest
import sys

@pytest.mark.skip
def test_mul():
    print("mul function")

@pytest.mark.skipif(sys.version_info<(3,8), reason = "python version not supported")
def test_add():
    print("add function")

@pytest.mark.xfail
def test_sub():
    assert True
    print("sub function")

@pytest.mark.xfail
def test_subb():
    assert False
    print("subb function")


def test_addubb():
    assert True
    print("test_addubb function")

@pytest.mark.parametrize("username,password",
                         [
                             ("grsvk", 211),
                             ("dfg", 234),
                             ("asd", 345)
                         ])
def test_adbb(username,password):
    # assert True
    # print("test_addubb function")
    print(username,password)