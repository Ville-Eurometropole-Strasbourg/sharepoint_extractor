import contextlib
import io
import os
import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import site_urls as su

HOST_ROOT = "https://intranet.example.org"
PATH_ROOT = "https://sp.example.org/sites/hr/"


class GetRootNameTest(unittest.TestCase):
    def run_with(self, value):
        env = {} if value is None else {su.ROOT_NAME_ENV: value}
        with mock.patch.dict(os.environ, env, clear=True):
            with contextlib.redirect_stdout(io.StringIO()):
                return su.get_root_name()

    def test_valid_names(self):
        self.assertEqual(self.run_with("hr"), "hr")
        self.assertEqual(self.run_with(" hr_2024-a "), "hr_2024-a")

    def test_missing_or_empty(self):
        for value in (None, "", "   "):
            with self.assertRaises(SystemExit):
                self.run_with(value)

    def test_invalid_characters(self):
        for value in ("sites/hr", "../hr", "h r", "rh.example"):
            with self.assertRaises(SystemExit):
                self.run_with(value)


class UrlTest(unittest.TestCase):
    def test_site_origin(self):
        self.assertEqual(su.site_origin(PATH_ROOT), "https://sp.example.org")
        self.assertEqual(su.site_origin(HOST_ROOT), HOST_ROOT)

    def test_root_path(self):
        self.assertEqual(su.root_path(PATH_ROOT), "/sites/hr")
        self.assertEqual(su.root_path(HOST_ROOT), "/")
        self.assertEqual(su.root_path(HOST_ROOT + "/"), "/")

    def test_absolute_url_does_not_repeat_the_root_path(self):
        self.assertEqual(
            su.absolute_url(PATH_ROOT, "/sites/hr/team"),
            "https://sp.example.org/sites/hr/team",
        )
        self.assertEqual(
            su.absolute_url(HOST_ROOT, "/service/team"),
            "https://intranet.example.org/service/team",
        )

    def test_absolute_url_adds_missing_slash(self):
        self.assertEqual(
            su.absolute_url(HOST_ROOT, "team"), "https://intranet.example.org/team"
        )
        self.assertEqual(su.absolute_url(HOST_ROOT, ""), HOST_ROOT + "/")

    def test_server_relative_path(self):
        self.assertEqual(
            su.server_relative_path("https://SP.example.org/sites/hr/team", PATH_ROOT),
            "/sites/hr/team",
        )
        self.assertIsNone(
            su.server_relative_path("https://other.example.org/sites/hr", PATH_ROOT)
        )
        self.assertIsNone(su.server_relative_path("https://sp.example.org", PATH_ROOT))


if __name__ == "__main__":
    unittest.main()
