import json
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src" / "extraction"))


def _network_disabled(*args, **kwargs):
    raise RuntimeError("SharePoint access is disabled in tests")


def _stub(name, *attributes):
    module = types.ModuleType(name)
    for attribute in attributes:
        setattr(module, attribute, _network_disabled)
    return module


with mock.patch.dict(
    sys.modules,
    {
        "sharepoint_client": _stub("sharepoint_client", "get_config", "build_context"),
        "persistent_session": _stub("persistent_session", "enable_persistent_session"),
    },
):
    import extract_metadata as em


class _Query:
    """A chainable client query; execute_query() logs `tag`, then returns or raises."""

    def __init__(self, server, tag, result=None, error=None):
        self.server, self.tag, self.result, self.error = server, tag, result, error

    def select(self, *args, **kwargs):
        return self

    def expand(self, *args, **kwargs):
        return self

    def get(self):
        return self

    def get_all(self, *args, **kwargs):
        return self

    def execute_query(self):
        self.server.requests.append(self.tag)
        if self.error is not None:
            raise self.error
        return self.result


def _principal(kind, title, principal_id):
    return types.SimpleNamespace(
        principal_type_name=kind,
        title=title,
        login_name=f"login:{title}",
        user_principal_name=None,
        id=principal_id,
    )


def _assignment(principal, role="Lecture", role_type=2):
    binding = types.SimpleNamespace(name=role, properties={"RoleTypeKind": role_type})
    return types.SimpleNamespace(member=principal, role_definition_bindings=[binding])


OWNERS = _principal("SharePointGroup", "Owners", 10)
ALICE = _principal("User", "Alice", 20)


class FakeSite:
    """
    One site as the client library sees it. Each server round trip is logged
    in self.requests, so a test can check what was asked and how often.
    """

    def __init__(self):
        self.requests = []
        self.lists = []  # list properties, as returned by web.lists
        self.items = {}  # list id -> item properties
        self.assignments = {("site",): [_assignment(OWNERS)]}  # securable key -> list
        self.groups = {10: [{"title": "Bob", "login_name": "login:Bob"}]}
        self.errors = {}  # tag -> exception raised by that request

    def add_list(self, list_id, title, base_type=1, items=(), **props):
        self.lists.append(
            {
                "Id": list_id,
                "Title": title,
                "Hidden": False,
                "BaseTemplate": 101 if base_type == 1 else 100,
                "BaseType": base_type,
                "EntityTypeName": title.replace(" ", "_x0020_"),
                "ItemCount": len(items),
                "HasUniqueRoleAssignments": False,
                **props,
            }
        )
        self.items[list_id] = list(items)

    def query(self, tag, result=None):
        return _Query(self, tag, result, self.errors.get(tag))

    def securable(self, key):
        return types.SimpleNamespace(
            role_assignments=self.query(("roles",) + key, self.assignments.get(key, []))
        )

    def context(self):
        site = self

        class Items:
            def __init__(self, list_id):
                self.list_id = list_id

            def select(self, fields):
                rows = [
                    types.SimpleNamespace(properties=p)
                    for p in site.items[self.list_id]
                ]
                return site.query(("items", self.list_id), rows)

            def get_by_id(self, item_id):
                return site.securable((self.list_id, item_id))

        class List:
            def __init__(self, list_id):
                self.list_id = list_id
                self.items = Items(list_id)
                self.title = next(p["Title"] for p in site.lists if p["Id"] == list_id)
                self.role_assignments = site.securable((list_id,)).role_assignments

            def get(self):
                return self

            def select(self, fields):
                return self

            def execute_query(self):
                return self

        class Lists:
            def select(self, fields):
                rows = [types.SimpleNamespace(properties=p) for p in site.lists]
                return site.query(("lists",), rows)

            def get_by_id(self, list_id):
                return List(list_id)

        class Groups:
            def get_by_id(self, group_id):
                return types.SimpleNamespace(
                    users=site.query(
                        ("members", group_id),
                        [
                            types.SimpleNamespace(
                                title=m["title"],
                                login_name=m["login_name"],
                                user_principal_name=None,
                            )
                            for m in site.groups.get(group_id, [])
                        ],
                    )
                )

        web = types.SimpleNamespace(
            get=lambda: site.query(
                ("web",), types.SimpleNamespace(has_unique_role_assignments=True)
            ),
            role_assignments=site.securable(("site",)).role_assignments,
            lists=Lists(),
            site_groups=Groups(),
        )
        return types.SimpleNamespace(web=web)


def _item(item_id, file_ref, unique, folder=False):
    return {
        "Id": item_id,
        "FileRef": file_ref,
        "FileSystemObjectType": 1 if folder else 0,
        "HasUniqueRoleAssignments": unique,
    }


class ExtractPermissionsTest(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.output_dir = tmp.name
        self.site = FakeSite()
        self.logger = mock.Mock()

    def extract(self):
        with mock.patch.object(em, "build_context", return_value=self.site.context()):
            summary = em.extract_permissions(
                "https://sp/team",
                "team",
                "user",
                "password",
                None,
                self.output_dir,
                self.logger,
            )
        path = os.path.join(self.output_dir, "team", "permissions.json")
        with open(path, encoding="utf-8") as f:
            return summary, json.load(f)

    def role_reads(self):
        return [r[1:] for r in self.site.requests if r[0] == "roles"]

    def test_only_items_with_their_own_permissions_are_read(self):
        self.site.add_list(
            "L1",
            "Documents",
            items=[
                _item(1, "/team/Docs/Confidentiel", True, folder=True),
                _item(2, "/team/Docs/note.docx", True),
                _item(3, "/team/Docs/public.docx", False),
            ],
        )
        self.site.add_list(
            "L2",
            "Tasks",
            base_type=0,
            items=[_item(7, "/team/Lists/Tasks/7_.000", True)],
        )
        self.site.assignments[("L1", 1)] = [_assignment(ALICE, "Modification", 3)]
        self.site.assignments[("L1", 2)] = [_assignment(ALICE)]
        self.site.assignments[("L2", 7)] = [_assignment(ALICE)]

        summary, permissions = self.extract()

        self.assertEqual(permissions["format"], em.PERMISSIONS_FORMAT)
        items = permissions["items_with_broken_inheritance"]
        self.assertEqual(
            {
                title: [(e["id"], e["kind"], e["url"]) for e in entries]
                for title, entries in items.items()
            },
            {
                "Documents": [
                    (1, "folder", "/team/Docs/Confidentiel"),
                    (2, "file", "/team/Docs/note.docx"),
                ],
                "Tasks": [(7, "item", "/team/Lists/Tasks/7_.000")],
            },
        )
        self.assertEqual(
            items["Documents"][0]["permissions"],
            [
                {
                    "type": "User",
                    "title": "Alice",
                    "login_name": "login:Alice",
                    "members_resolvable": True,
                    "roles": ["Modification"],
                    "role_types": [3],
                }
            ],
        )
        # Role assignments: the site, then only the 3 items flagged.
        self.assertEqual(
            self.role_reads(), [("site",), ("L1", 1), ("L1", 2), ("L2", 7)]
        )
        self.assertEqual(permissions["items_scan_errors"], [])
        self.assertEqual(summary["items_with_broken_inheritance_count"], 3)
        self.assertEqual(summary["items_scan_errors_count"], 0)

    def add_lists_of_every_kind(self, **props):
        """A hidden list, two system lists, a hidden list kept, an empty one."""
        self.site.add_list(
            "L1", "Hidden", items=[_item(1, "/a", True)], Hidden=True, **props
        )
        self.site.add_list(
            "L2", "Style Library", items=[_item(1, "/b", True)], **props
        )  # system list, its content is not extracted
        self.site.add_list(
            "L3",
            "Site Pages",
            items=[_item(1, "/p/home.aspx", True)],
            EntityTypeName="SitePages",
            BaseTemplate=119,
            **props,
        )  # system list, its pages are read by inventory_site_pages.py
        self.site.add_list(
            "L4",
            "Workflow History",
            base_type=0,
            items=[_item(1, "/c", True)],
            Hidden=True,
            BaseTemplate=140,  # hidden but extracted
            **props,
        )
        self.site.add_list("L5", "Empty", **props)

    def test_system_lists_are_scanned_but_skipped_and_empty_lists_are_not(self):
        self.add_lists_of_every_kind()

        _, permissions = self.extract()

        scanned = [r[1] for r in self.site.requests if r[0] == "items"]
        self.assertEqual(scanned, ["L2", "L3", "L4"])
        self.assertEqual(
            {
                title: [(e["kind"], e["url"]) for e in entries]
                for title, entries in permissions[
                    "items_with_broken_inheritance"
                ].items()
            },
            {
                "Style Library": [("file", "/b")],
                "Site Pages": [("file", "/p/home.aspx")],
                "Workflow History": [("item", "/c")],
            },
        )

    def test_list_permissions_follow_the_same_rule_as_item_permissions(self):
        self.add_lists_of_every_kind(HasUniqueRoleAssignments=True)
        for list_id in ("L1", "L2", "L3", "L4", "L5"):
            self.site.assignments[(list_id,)] = [_assignment(ALICE)]

        _, permissions = self.extract()

        self.assertEqual(
            list(permissions["lists_with_broken_inheritance"]),
            ["Style Library", "Site Pages", "Workflow History", "Empty"],
        )
        self.assertNotIn(("L1",), self.role_reads())

    def test_a_list_that_cannot_be_scanned_is_recorded(self):
        self.site.add_list("L1", "Broken", items=[_item(1, "/a", True)])
        self.site.add_list("L2", "Fine", items=[_item(2, "/b", True)])
        self.site.errors[("items", "L1")] = RuntimeError("500 Server Error")

        summary, permissions = self.extract()

        self.assertEqual(list(permissions["items_with_broken_inheritance"]), ["Fine"])
        self.assertEqual(
            permissions["items_scan_errors"],
            [{"list": "Broken", "item_id": None, "error": "500 Server Error"}],
        )
        self.assertEqual(summary["items_scan_errors_count"], 1)

    def test_an_item_without_the_property_is_an_error_not_a_silent_zero(self):
        item = _item(1, "/a", False)
        del item["HasUniqueRoleAssignments"]
        self.site.add_list("L1", "Documents", items=[item])

        summary, permissions = self.extract()

        self.assertEqual(permissions["items_with_broken_inheritance"], {})
        self.assertEqual(
            [e["list"] for e in permissions["items_scan_errors"]], ["Documents"]
        )
        self.assertEqual(summary["items_scan_errors_count"], 1)

    def test_unreadable_item_permissions_are_kept_and_recorded(self):
        self.site.add_list("L1", "Documents", items=[_item(1, "/a", True)])
        self.site.errors[("roles", "L1", 1)] = RuntimeError("403")

        summary, permissions = self.extract()

        entry = permissions["items_with_broken_inheritance"]["Documents"][0]
        self.assertEqual((entry["id"], entry["permissions"]), (1, None))
        self.assertEqual(
            [(e["list"], e["item_id"]) for e in permissions["items_scan_errors"]],
            [("Documents", 1)],
        )
        self.assertEqual(summary["items_scan_errors_count"], 1)

    def test_unreadable_lists_are_recorded(self):
        self.site.errors[("lists",)] = RuntimeError("timeout")

        summary, permissions = self.extract()

        self.assertEqual(
            permissions["items_scan_errors"],
            [{"list": None, "item_id": None, "error": "timeout"}],
        )
        self.assertEqual(summary["items_scan_errors_count"], 1)

    def test_group_members_are_read_once_per_site(self):
        # Owners on the site, on a list with broken inheritance and on an item.
        self.site.add_list(
            "L1",
            "Documents",
            items=[_item(1, "/a", True)],
            HasUniqueRoleAssignments=True,
        )
        self.site.assignments[("L1",)] = [_assignment(OWNERS)]
        self.site.assignments[("L1", 1)] = [_assignment(OWNERS), _assignment(ALICE)]

        summary, permissions = self.extract()

        member_reads = [r for r in self.site.requests if r[0] == "members"]
        self.assertEqual(member_reads, [("members", 10)])
        bob = [{"title": "Bob", "login_name": "login:Bob"}]
        # Members are still written in every group entry.
        self.assertEqual(permissions["site_permissions"][0]["members"], bob)
        self.assertEqual(
            permissions["lists_with_broken_inheritance"]["Documents"][0]["members"],
            bob,
        )
        item = permissions["items_with_broken_inheritance"]["Documents"][0]
        self.assertEqual(item["permissions"][0]["members"], bob)
        self.assertEqual(summary["group_members_requests"], 1)
        self.assertEqual(summary["group_members_reused"], 2)

    def test_a_failed_group_read_is_retried(self):
        self.site.add_list("L1", "Documents", items=[], HasUniqueRoleAssignments=True)
        self.site.assignments[("L1",)] = [_assignment(OWNERS)]
        failures = iter([RuntimeError("timeout")])

        original = self.site.query

        def flaky(tag, result=None):
            query = original(tag, result)
            if tag == ("members", 10):
                query.error = next(failures, None)
            return query

        self.site.query = flaky

        summary, permissions = self.extract()

        member_reads = [r for r in self.site.requests if r[0] == "members"]
        self.assertEqual(member_reads, [("members", 10), ("members", 10)])
        self.assertIsNone(permissions["site_permissions"][0]["members"])
        self.assertEqual(
            permissions["lists_with_broken_inheritance"]["Documents"][0]["members"],
            [{"title": "Bob", "login_name": "login:Bob"}],
        )
        # A failure is not kept: 2 requests, nothing reused.
        self.assertEqual(summary["group_members_requests"], 2)
        self.assertEqual(summary["group_members_reused"], 0)


class RunPermissionsPassTest(unittest.TestCase):
    """Sites whose permissions.json predates the current format are redone."""

    def test_only_up_to_date_permissions_are_skipped(self):
        with tempfile.TemporaryDirectory() as output_dir:
            files = {
                "current": json.dumps({"format": em.PERMISSIONS_FORMAT}),
                "old": json.dumps({"site_permissions": []}),
                "unreadable": "{",
                "missing": None,
            }
            for slug, content in files.items():
                os.makedirs(os.path.join(output_dir, slug))
                if content is not None:
                    with open(
                        os.path.join(output_dir, slug, "permissions.json"), "w"
                    ) as f:
                        f.write(content)
            sites = [{"slug": s, "url_suffix": "/" + s} for s in files]

            with mock.patch.object(
                em, "extract_permissions", return_value={}
            ) as extract:
                em.run_permissions_pass(
                    sites, "u", "p", None, "https://sp", output_dir, mock.Mock()
                )

        self.assertEqual(
            [c.args[1] for c in extract.call_args_list],
            ["old", "unreadable", "missing"],
        )


if __name__ == "__main__":
    unittest.main()
