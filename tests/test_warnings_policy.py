import unittest

import tests


class WarningPolicySmokeTests(unittest.TestCase):
    def test_resourcewarning_is_error(self):
        self.assertTrue(tests.RESOURCEWARNING_AS_ERROR_CONFIGURED)
