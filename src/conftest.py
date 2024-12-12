import pytest

@pytest.fixture
def setup_and_teardown():
    # Setup code
    print("\nSetup code executed")
    yield
    # Teardown code
    print("\nTeardown code executed")



@pytest.fixture
def getData():
    name = input("Enter the entary name: ")
    print(name)


@pytest.fixture(scope = 'session', autouse=True)
def myfixtures():
    print("conftest fixtures is called")