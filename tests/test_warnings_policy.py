import unittest
import warnings


class WarningPolicySmokeTests(unittest.TestCase):
    def test_resourcewarning_is_error(self):
        self.assertTrue(
            any(
                action == "error" and category is ResourceWarning
                for action, _message, category, *_ in warnings.filters
            )
        )
