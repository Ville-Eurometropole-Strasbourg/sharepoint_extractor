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
    import download_files as df


CONTENT = b"binary"
LIBRARY_ROOT = "/sites/team/Shared Documents"


def _item(item_id, file_ref, folder=False):
    return types.SimpleNamespace(
        properties={
            "Id": item_id,
            "FileRef": file_ref,
            "FileLeafRef": file_ref.rsplit("/", 1)[1],
            "FileSystemObjectType": 1 if folder else 0,
            "FileSizeDisplay": None if folder else str(len(CONTENT)),
        }
    )


def _fake_download(stream):
    stream.write(CONTENT)
    return mock.MagicMock()


class DownloadLibraryTest(unittest.TestCase):
    """
    One library downloaded on its own (`<site path> <library title>`): each
    binary must land where extract_metadata.py put its .meta.json, i.e. under
    the library folder, keeping the subfolders below the library root.
    """

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.lib_dir = os.path.join(tmp.name, "documents", "Documents")
        self.logger = mock.Mock()

    def download(self, file_refs, root_folder_error=None):
        items = [_item(1, LIBRARY_ROOT + "/2019", folder=True)] + [
            _item(item_id, ref) for item_id, ref in enumerate(file_refs, start=2)
        ]
        ctx = mock.MagicMock()
        sp_list = ctx.web.lists.get_by_title.return_value
        read_items = sp_list.items.select.return_value.get_all.return_value
        read_items.execute_query.return_value = items
        read_root = sp_list.root_folder.get.return_value.execute_query
        if root_folder_error:
            read_root.side_effect = root_folder_error
        else:
            read_root.return_value.properties = {"ServerRelativeUrl": LIBRARY_ROOT}
        sp_list.items.get_by_id.return_value.file.download.side_effect = _fake_download
        with mock.patch.object(df, "build_context", return_value=ctx):
            return df.download_library(
                "https://sp/sites/team",
                "user",
                "password",
                None,
                "Documents",
                self.lib_dir,
                self.logger,
                1,
            )

    def files_on_disk(self):
        return sorted(
            os.path.relpath(os.path.join(dirpath, name), self.lib_dir).replace(
                os.sep, "/"
            )
            for dirpath, _, names in os.walk(self.lib_dir)
            for name in names
        )

    def test_files_keep_their_folders_when_the_title_differs_from_the_url(self):
        results = self.download(
            [
                LIBRARY_ROOT + "/2019/report.docx",
                LIBRARY_ROOT + "/2020/report.docx",
                LIBRARY_ROOT + "/readme.txt",
            ]
        )

        self.assertEqual(
            self.files_on_disk(),
            ["2019/report.docx", "2020/report.docx", "readme.txt"],
        )
        self.assertEqual([r["status"] for r in results], ["ok", "ok", "ok"])

    def test_files_are_written_flat_when_the_root_folder_cannot_be_read(self):
        results = self.download(
            [LIBRARY_ROOT + "/2019/report.docx"], root_folder_error=RuntimeError("403")
        )

        self.assertEqual(self.files_on_disk(), ["report.docx"])
        self.assertEqual([r["status"] for r in results], ["ok"])
        self.logger.warning.assert_called()


class DiscoverAllDownloadableFilesTest(unittest.TestCase):
    """--site and --all: the files of every extracted site, from its folder."""

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.extraction_dir = tmp.name
        self.logger = mock.Mock()

    def extracted_site(self, slug, url, file_ref):
        site_dir = os.path.join(self.extraction_dir, "intranet", slug)
        lib_dir = os.path.join(site_dir, "documents", "Documents")
        os.makedirs(lib_dir)
        with open(os.path.join(site_dir, "site_metadata.json"), "w") as f:
            json.dump({"Url": url}, f)
        name = file_ref.rsplit("/", 1)[1]
        with open(os.path.join(lib_dir, name + ".meta.json"), "w") as f:
            json.dump({"FileLeafRef": name, "FileRef": file_ref, "Id": 1}, f)

    def test_the_files_of_a_host_root_site_are_found(self):
        # A host root's Url has no path: https://intranet.example.org
        self.extracted_site(
            "_root", "https://intranet.example.org", "/Documents/a.docx"
        )
        self.extracted_site(
            "team", "https://intranet.example.org/team", "/team/Documents/b.docx"
        )

        tasks = df.discover_all_downloadable_files(self.extraction_dir, self.logger)

        self.assertEqual(
            [(t["site_slug"], t["server_relative_url"], t["name"]) for t in tasks],
            [("_root", "/", "a.docx"), ("team", "/team", "b.docx")],
        )
        self.logger.warning.assert_not_called()


if __name__ == "__main__":
    unittest.main()
