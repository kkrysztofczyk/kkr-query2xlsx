"""
Bootstraps unittest discovery so startup policies (e.g., warnings filters)
are applied before other test modules execute.

`python -m unittest discover -s tests -p 'test_*.py'` imports test modules as
top-level modules and does not import `tests/__init__.py` automatically.
Importing the package here forces `tests/__init__.py` to run early.
"""

import tests  # noqa: F401
