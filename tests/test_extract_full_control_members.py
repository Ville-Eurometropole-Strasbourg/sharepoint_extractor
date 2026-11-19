import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src" / "reporting"))

import extract_full_control_members as fc


def entry(roles, role_types=None, kind="User", title="Person A"):
    result = {"type": kind, "title": title, "login_name": "i:0#.w|a", "roles": roles}
    if role_types is not None:
        result["role_types"] = role_types
    return result


class IsFullControlTest(unittest.TestCase):
    def test_built_in_role_in_any_language(self):
        self.assertTrue(fc.is_full_control(entry(["Full Control"], [5])))
        self.assertTrue(fc.is_full_control(entry(["Contrôle total"], [5])))
        self.assertFalse(fc.is_full_control(entry(["Lecture"], [2])))

    def test_custom_roles_by_name(self):
        custom = entry(["Custom full"], [0])
        self.assertFalse(fc.is_full_control(custom))
        self.assertTrue(fc.is_full_control(custom, ["Custom full"]))

    def test_files_without_role_types(self):
        old = entry(["Contrôle total"])
        self.assertFalse(fc.is_full_control(old))
        self.assertTrue(fc.is_full_control(old, ["Contrôle total"]))


class ExtractRowsForSiteTest(unittest.TestCase):
    def test_group_members_and_direct_grants(self):
        group = entry(["Full Control"], [5], kind="SharePointGroup", title="Owners")
        group["members"] = [{"title": "Person B", "login_name": "i:0#.w|b"}]
        permissions = {
            "site_permissions": [
                group,
                entry(["Full Control"], [5]),
                entry(["Read"], [2]),
            ]
        }
        rows = fc.extract_rows_for_site("id", "Site", "https://sp/site", permissions)
        self.assertEqual(
            [(r["group_title"], r["person_name"]) for r in rows],
            [("Owners", "Person B"), ("(direct)", "Person A")],
        )


if __name__ == "__main__":
    unittest.main()
