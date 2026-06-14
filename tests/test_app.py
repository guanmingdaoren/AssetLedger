from pathlib import Path
import sys
import unittest
from unittest.mock import patch

from asset_ledger.app import application_base_dir, default_workbook_path, resource_path


class AppPathTests(unittest.TestCase):
    def test_frozen_application_uses_executable_directory(self) -> None:
        executable = Path("C:/Program Files/AssetLedger/设备资产台账.exe")

        with (
            patch.object(sys, "frozen", True, create=True),
            patch.object(sys, "executable", str(executable)),
        ):
            self.assertEqual(application_base_dir(), executable.parent)

    def test_default_workbook_is_stored_outside_executable(self) -> None:
        executable = Path("C:/Portable/AssetLedger/设备资产台账.exe")

        with (
            patch.object(sys, "frozen", True, create=True),
            patch.object(sys, "executable", str(executable)),
        ):
            self.assertEqual(
                default_workbook_path(),
                executable.parent / "data" / "设备资产台账.xlsx",
            )

    def test_source_application_icon_exists_at_resource_path(self) -> None:
        self.assertTrue(resource_path("assets/app_icon.ico").exists())


if __name__ == "__main__":
    unittest.main()
