"""Test package for unittest discovery."""

import os
import warnings


warnings.simplefilter("error", ResourceWarning)


def load_tests(loader, tests, pattern):
    # Enables: python -m unittest -v tests
    if pattern is None:
        pattern = "test_*.py"
    this_dir = os.path.dirname(__file__)
    return loader.discover(start_dir=this_dir, pattern=pattern)
