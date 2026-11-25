import contextlib
import io
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src" / "tools"))

import check_broken_inheritance as cbi
import inventory_site


def _principal(title):
    return {"type": "User", "title": title, "login_name": title, "roles": ["Lecture"]}


def _item(item_id, kind):
    return {
        "id": item_id,
        "kind": kind,
        "url": f"/x/{item_id}",
        "permissions": [_principal("A"), _principal("B")],
    }


SITES = {
    "team": {
        "format": 2,
        "site_has_unique_role_assignments": True,
        "site_permissions": [_principal("Owner")],
        "lists_with_broken_inheritance": {},
        "items_with_broken_inheritance": {
            "Documents": [_item(1, "folder"), _item(2, "file"), _item(3, "file")],
            "Tasks": [_item(7, "item")],
        },
        "items_scan_errors": [],
    },
    "notes": {
        "format": 2,
        "site_has_unique_role_assignments": True,
        "site_permissions": [],
        "lists_with_broken_inheritance": {},
        "items_with_broken_inheritance": {},
        "items_scan_errors": [{"list": "Notes", "item_id": None, "error": "500"}],
    },
    "old": {
        "site_has_unique_role_assignments": True,
        "site_permissions": [],
        "lists_with_broken_inheritance": {},
    },
}


def _write_sites(base):
    for slug, permissions in SITES.items():
        site_dir = os.path.join(base, "intranet", slug)
        os.makedirs(site_dir)
        with open(os.path.join(site_dir, "permissions.json"), "w") as f:
            json.dump(permissions, f)
        with open(os.path.join(site_dir, "site_metadata.json"), "w") as f:
            json.dump({"ServerRelativeUrl": f"/{slug}"}, f)


def _output(function, *args):
    out = io.StringIO()
    with contextlib.redirect_stdout(out):
        function(*args)
    return out.getvalue()


class CheckBrokenInheritanceTest(unittest.TestCase):
    def test_items_are_counted_by_kind_and_older_files_are_reported(self):
        with tempfile.TemporaryDirectory() as base:
            _write_sites(base)
            csv_path = os.path.join(base, "out.csv")
            argv = ["check", "intranet", "--base", base, "--csv", csv_path]
            with mock.patch.object(sys, "argv", argv):
                output = _output(cbi.main)
            with open(csv_path, encoding="utf-8") as f:
                rows = f.read().splitlines()

        self.assertIn(
            "permissions.json without item permissions (older format): 1", output
        )
        self.assertIn("sites with at least one item with broken inheritance: 1", output)
        self.assertIn(
            "items with broken inheritance (total): 4 (file=2, folder=1, item=1)",
            output,
        )
        self.assertIn("principal entries on items: 8", output)
        self.assertIn("item scan errors: 1 (sites with at least one: 1)", output)
        self.assertEqual(
            rows[0],
            "slug,level,site_unique,lists_broken,items_broken,items_scan_errors",
        )
        self.assertIn("old,top,True,0,,", rows)
        self.assertIn("team,top,True,0,4,0", rows)


class InventorySitePermissionsTest(unittest.TestCase):
    def section(self, slug):
        with tempfile.TemporaryDirectory() as base:
            _write_sites(base)
            return _output(
                inventory_site.permissions_section,
                os.path.join(base, "intranet", slug),
                set(),
            )

    def test_items_are_counted(self):
        output = self.section("team")
        self.assertIn(
            "items with broken inheritance: 4, file=2, folder=1, item=1", output
        )
        self.assertIn("item scan errors: 0", output)

    def test_older_files_say_items_were_not_extracted(self):
        self.assertIn("not extracted (older permissions.json)", self.section("old"))


if __name__ == "__main__":
    unittest.main()
