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
    import extract_versions as ev


class NeedsHistoryTest(unittest.TestCase):
    def test_first_draft_never_has_history(self):
        for minor_versions in (True, False, None):
            self.assertFalse(ev.needs_history("0.1", minor_versions))

    def test_first_major_version_without_minor_versions(self):
        self.assertFalse(ev.needs_history("1.0", False))

    def test_first_major_version_published_from_drafts(self):
        self.assertTrue(ev.needs_history("1.0", True))

    def test_first_major_version_when_the_setting_is_unknown(self):
        self.assertTrue(ev.needs_history("1.0", None))

    def test_later_versions(self):
        for label in ("0.2", "1.1", "2.0", "12.0"):
            self.assertTrue(ev.needs_history(label, False))


def _write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data), encoding="utf-8")


class ExtractSiteVersionsTest(unittest.TestCase):
    """Which files get a /versions call, with SharePoint replaced at its edge."""

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.site_dir = Path(tmp.name)
        _write_json(self.site_dir / "site_metadata.json", {"Url": "https://sp/team"})
        self.files = {
            # library with minor versions: 1.0 published from drafts
            "Drafts/a.docx": "1.0",
            "Drafts/sub/b.docx": "1.0",
            "Drafts/c.docx": "0.1",
            # library without minor versions
            "Plain/d.docx": "1.0",
            "Plain/e.docx": "3.0",
            # library without a 1.0 file: its setting is never read
            "Later/f.docx": "2.0",
        }
        for relative, label in self.files.items():
            _write_json(
                self.site_dir / "documents" / (relative + ".meta.json"),
                {"FileRef": "/team/" + relative, "OData__UIVersionString": label},
            )
        self.minor_versions = {"Drafts": True, "Plain": False}

    def run_extraction(self, minor_versions):
        def fetch_versions(ctx, file_ref):
            return {"file_ref": file_ref, "history": [], "history_count": 0}

        logger = mock.Mock()
        with (
            mock.patch.object(ev, "build_context", return_value=object()),
            mock.patch.object(
                ev, "fetch_versions", side_effect=fetch_versions
            ) as fetch,
            mock.patch.object(
                ev, "library_minor_versions", side_effect=minor_versions
            ) as lookup,
        ):
            counts = ev.extract_site_versions(
                str(self.site_dir), "user", "password", None, logger
            )
        return counts, fetch, lookup, logger

    def queried(self, fetch_mock):
        return sorted(c.args[1] for c in fetch_mock.call_args_list)

    def looked_up(self, lookup_mock):
        """Libraries whose setting was read, from the file_ref of each call."""
        return sorted(c.args[1].split("/")[2] for c in lookup_mock.call_args_list)

    def test_1_0_files_are_queried_only_in_libraries_with_minor_versions(self):
        def minor_versions(ctx, file_ref):
            return self.minor_versions[file_ref.split("/")[2]]

        counts, fetch, lookup, _ = self.run_extraction(minor_versions)

        self.assertEqual(
            self.queried(fetch),
            [
                "/team/Drafts/a.docx",
                "/team/Drafts/sub/b.docx",
                "/team/Later/f.docx",
                "/team/Plain/e.docx",
            ],
        )
        # Once per library holding a 1.0 file, never for "Later".
        self.assertEqual(self.looked_up(lookup), ["Drafts", "Plain"])
        self.assertEqual(counts["queried"], 4)
        self.assertEqual(counts["skipped_no_history"], 2)
        self.assertTrue(
            (self.site_dir / "documents" / "Drafts" / "a.docx.versions.json").exists()
        )

    def test_1_0_files_are_queried_when_the_setting_cannot_be_read(self):
        def unreadable(ctx, file_ref):
            raise RuntimeError("403")

        counts, fetch, lookup, logger = self.run_extraction(unreadable)

        self.assertEqual(
            self.queried(fetch),
            [
                "/team/Drafts/a.docx",
                "/team/Drafts/sub/b.docx",
                "/team/Later/f.docx",
                "/team/Plain/d.docx",
                "/team/Plain/e.docx",
            ],
        )
        # Not retried per file, and each unreadable library is reported once.
        self.assertEqual(self.looked_up(lookup), ["Drafts", "Plain"])
        warned = " ".join(c.args[0] for c in logger.warning.call_args_list)
        self.assertEqual(logger.warning.call_count, 2)
        self.assertIn("'Drafts'", warned)
        self.assertIn("'Plain'", warned)
        self.assertEqual(counts["skipped_no_history"], 1)

    def test_folders_and_document_sets_are_skipped(self):
        for relative, label in (("Sets/Lot 1", "1.0"), ("Plain/archive", "2.0")):
            _write_json(
                self.site_dir / "documents" / (relative + ".meta.json"),
                {
                    "FileRef": "/team/" + relative,
                    "FileSystemObjectType": 1,
                    "OData__UIVersionString": label,
                },
            )

        def minor_versions(ctx, file_ref):
            return self.minor_versions[file_ref.split("/")[2]]

        counts, fetch, lookup, _ = self.run_extraction(minor_versions)

        self.assertNotIn("/team/Sets/Lot 1", self.queried(fetch))
        self.assertNotIn("/team/Plain/archive", self.queried(fetch))
        self.assertEqual(self.looked_up(lookup), ["Drafts", "Plain"])
        self.assertEqual(counts["skipped_folders"], 2)
        self.assertEqual(counts["queried"], 4)
        self.assertFalse(
            (self.site_dir / "documents" / "Sets" / "Lot 1.versions.json").exists()
        )


if __name__ == "__main__":
    unittest.main()
