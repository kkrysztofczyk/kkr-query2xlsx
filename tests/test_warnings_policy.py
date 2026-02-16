import unittest
import warnings

import tests


class WarningPolicySmokeTests(unittest.TestCase):
    def test_resourcewarning_is_error(self):
        # tests/__init__.py configures the filter; importing `tests` above ensures this
        # smoke test works both for `python -m unittest -v tests` and plain discover.
        self.assertTrue(
            any(
                action == "error" and category is ResourceWarning
                for action, _message, category, *_ in warnings.filters
            )
        )
