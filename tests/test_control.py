from __future__ import annotations

import unittest

from src.control import DEFAULT_CONTROL_CODE, check_software_control


class ControlTests(unittest.TestCase):
    def test_non_windows_environment_does_not_require_dll(self) -> None:
        result = check_software_control(DEFAULT_CONTROL_CODE)

        self.assertTrue(result.ok)


if __name__ == "__main__":
    unittest.main()
