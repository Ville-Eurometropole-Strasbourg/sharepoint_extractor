import json
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
    import extract_list_versions as elv


class _Query:
    """A chainable client query; execute_query() logs `tag`, then returns or raises."""

    def __init__(self, server, tag, result=None, error=None):
        self.server, self.tag, self.result, self.error = server, tag, result, error

    def select(self, *args, **kwargs):
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


def _version(label, title):
    return types.SimpleNamespace(
        properties={"VersionLabel": label, "Title": title, "IsCurrentVersion": False}
    )


class FakeList:
    def __init__(self, server, list_id, title, labels, **props):
        self.server = server
        # {item id: current label}; a list item's versions are 1.0 .. label.
        self.labels = labels
        self.failing = set()
        self.props = {
            "Id": list_id,
            "Title": title,
            "BaseType": 0,
            "BaseTemplate": 100,
            "Hidden": False,
            "EntityTypeName": title + "List",
            "ItemCount": len(labels),
            "EnableVersioning": True,
            "EnableMinorVersions": False,
            "LastItemModifiedDate": "2026-01-01 10:00:00",
            **props,
        }

    @property
    def items(self):
        fake = self

        class Items:
            def select(self, fields):
                return _Query(
                    fake.server,
                    f"labels:{fake.props['Title']}",
                    [
                        types.SimpleNamespace(
                            properties={"Id": i, elv.LABEL_FIELD: label}
                        )
                        for i, label in fake.labels.items()
                    ],
                )

            def get_by_id(self, item_id):
                label = fake.labels[item_id]
                count = int(float(label)) if label else 0
                return types.SimpleNamespace(
                    versions=_Query(
                        fake.server,
                        f"versions:{fake.props['Title']}:{item_id}",
                        [_version(f"{n}.0", f"v{n}") for n in range(count, 0, -1)],
                        error=(
                            RuntimeError("500") if item_id in fake.failing else None
                        ),
                    )
                )

        return Items()


class FakeServer:
    def __init__(self):
        self.requests = []
        self.lists = {}

    def add(self, list_id, title, labels, **props):
        self.lists[list_id] = FakeList(self, list_id, title, labels, **props)
        return self.lists[list_id]

    @property
    def web(self):
        server = self

        class Lists:
            def select(self, fields):
                return _Query(
                    server,
                    "lists",
                    [
                        types.SimpleNamespace(properties=dict(l.props))
                        for l in server.lists.values()
                    ],
                )

            def get_by_id(self, list_id):
                return server.lists[list_id]

        return types.SimpleNamespace(lists=Lists())


class ExtractSiteListVersionsTest(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.site_dir = Path(tmp.name)
        (self.site_dir / "site_metadata.json").write_text(
            json.dumps({"Url": "https://sp/team"}), encoding="utf-8"
        )
        self.server = FakeServer()
        self.logger = mock.Mock()

    def run_extraction(self, dry_run=False):
        self.server.requests.clear()
        with mock.patch.object(elv, "build_context", return_value=self.server):
            return elv.extract_site_list_versions(
                str(self.site_dir), "user", "password", None, self.logger, dry_run
            )

    def saved(self, title):
        path = self.site_dir / "lists" / title / elv.VERSIONS_FILE
        return json.loads(path.read_text(encoding="utf-8"))

    def versions_read(self):
        return sorted(r for r in self.server.requests if r.startswith("versions:"))

    def test_only_items_with_a_history_are_queried(self):
        self.server.add("L1", "Issues", {1: "1.0", 2: "3.0", 3: "2.0", 4: "0.1"})
        # Versioning turned off later: past versions are still there.
        self.server.add("L2", "Requests", {7: "2.0"}, EnableVersioning=False)

        counts = self.run_extraction()

        self.assertEqual(
            self.versions_read(),
            ["versions:Issues:2", "versions:Issues:3", "versions:Requests:7"],
        )
        issues = self.saved("Issues")
        self.assertEqual(sorted(issues["items"]), ["2", "3"])
        self.assertEqual(issues["items"]["2"]["current_label"], "3.0")
        self.assertEqual(
            [v["VersionLabel"] for v in issues["items"]["2"]["versions"]],
            ["3.0", "2.0", "1.0"],
        )
        self.assertEqual(issues["errors"], [])
        self.assertFalse(self.saved("Requests")["enable_versioning"])
        self.assertEqual(counts["items_read"], 5)
        self.assertEqual(counts["items_with_history"], 3)
        self.assertEqual(counts["items_queried"], 3)
        self.assertEqual(counts["versions"], 7)

    def test_1_0_items_are_queried_with_minor_versions(self):
        self.server.add("L1", "Drafts", {1: "1.0", 2: "0.1"}, EnableMinorVersions=True)

        self.run_extraction()

        self.assertEqual(self.versions_read(), ["versions:Drafts:1"])

    def test_libraries_skipped_hidden_and_empty_lists_are_not_read(self):
        self.server.add("L1", "Docs", {1: "2.0"}, BaseType=1)
        self.server.add("L2", "UserInfo", {1: "2.0"}, Hidden=True, BaseTemplate=112)
        self.server.add("L3", "Empty", {})

        counts = self.run_extraction()

        self.assertEqual(self.server.requests, ["lists"])
        self.assertEqual(counts["lists_checked"], 0)
        self.assertFalse((self.site_dir / "lists").exists())

    def test_items_without_a_label_are_counted_not_queried(self):
        self.server.add("L1", "Issues", {1: None, 2: "2.0"})

        counts = self.run_extraction()

        self.assertEqual(self.versions_read(), ["versions:Issues:2"])
        self.assertEqual(counts["items_without_label"], 1)

    def test_an_unchanged_list_is_skipped(self):
        self.server.add("L1", "Issues", {1: "2.0"})
        self.run_extraction()

        counts = self.run_extraction()

        self.assertEqual(self.server.requests, ["lists"])
        self.assertEqual(counts["lists_unchanged"], 1)

    def test_a_changed_list_reads_only_new_labels_and_drops_deleted_items(self):
        issues = self.server.add("L1", "Issues", {1: "2.0", 2: "2.0", 3: "3.0"})
        self.run_extraction()
        issues.labels = {1: "2.0", 2: "4.0"}  # 2 edited, 3 deleted
        issues.props["LastItemModifiedDate"] = "2026-02-01 10:00:00"

        counts = self.run_extraction()

        self.assertEqual(self.versions_read(), ["versions:Issues:2"])
        saved = self.saved("Issues")["items"]
        self.assertEqual(sorted(saved), ["1", "2"])
        self.assertEqual(len(saved["2"]["versions"]), 4)
        self.assertEqual(counts["items_already_saved"], 1)

    def test_failed_items_are_retried_on_the_next_run(self):
        issues = self.server.add("L1", "Issues", {1: "2.0", 2: "2.0"})
        issues.failing = {2}

        counts = self.run_extraction()

        self.assertEqual(counts["errors"], 1)
        saved = self.saved("Issues")
        self.assertEqual(sorted(saved["items"]), ["1"])
        self.assertEqual(saved["errors"][0]["item_id"], 2)
        self.assertFalse((self.site_dir / "lists" / "Issues" / elv.STATE_FILE).exists())

        issues.failing = set()
        counts = self.run_extraction()

        self.assertEqual(self.versions_read(), ["versions:Issues:2"])
        self.assertEqual(counts["errors"], 0)
        self.assertEqual(sorted(self.saved("Issues")["items"]), ["1", "2"])
        self.assertTrue((self.site_dir / "lists" / "Issues" / elv.STATE_FILE).exists())

    def test_progress_is_saved_during_a_list(self):
        issues = self.server.add("L1", "Issues", {1: "2.0", 2: "2.0", 3: "2.0"})
        reads = []

        def interrupted(sp_list, item_id):
            if len(reads) == 2:
                raise KeyboardInterrupt
            reads.append(item_id)
            return [{"VersionLabel": "2.0"}, {"VersionLabel": "1.0"}]

        with (
            mock.patch.object(elv, "SAVE_EVERY", 1),
            mock.patch.object(elv, "fetch_item_versions", side_effect=interrupted),
            self.assertRaises(KeyboardInterrupt),
        ):
            self.run_extraction()

        self.assertEqual(sorted(self.saved("Issues")["items"]), ["1", "2"])
        self.run_extraction()
        self.assertEqual(self.versions_read(), ["versions:Issues:3"])

    def test_dry_run_counts_without_querying_or_writing(self):
        self.server.add("L1", "Issues", {1: "1.0", 2: "3.0", 3: "2.0"})

        counts = self.run_extraction(dry_run=True)

        self.assertEqual(self.versions_read(), [])
        self.assertEqual(counts["items_to_query"], 2)
        self.assertFalse((self.site_dir / "lists").exists())


if __name__ == "__main__":
    unittest.main()
