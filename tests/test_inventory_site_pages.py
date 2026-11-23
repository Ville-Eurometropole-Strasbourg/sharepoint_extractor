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
    import inventory_site_pages as isp


class _ServerItems:
    """
    A library's items as the client library reads them: one request returns
    a single page ($top items, 100 by default); get_all follows the pages.
    """

    def __init__(self, items):
        self._items = items
        self._limit = 100

    def select(self, fields):
        return self

    def top(self, count):
        self._limit = count
        return self

    def get(self):
        return self

    def get_all(self, page_size=None):
        self._limit = len(self._items)
        return self

    def execute_query(self):
        return self._items[: self._limit]


def _page(item_id):
    name = f"page-{item_id}.aspx"
    return types.SimpleNamespace(
        properties={
            "Id": item_id,
            "FileLeafRef": name,
            "FileRef": f"/sites/team/Pages/{name}",
            "FileSystemObjectType": 0,
        }
    )


class InventorySitePagesTest(unittest.TestCase):
    def test_every_page_of_a_large_library_is_inventoried(self):
        pages = [_page(i) for i in range(1, 602)]
        library = types.SimpleNamespace(
            properties={"BaseTemplate": 850, "Title": "Pages", "ItemCount": 601},
            items=_ServerItems(pages),
        )
        ctx = mock.MagicMock()
        ctx.web.lists.get.return_value.execute_query.return_value = [library]

        with tempfile.TemporaryDirectory() as output_dir:
            site_dir = os.path.join(output_dir, "extraction", "root", "team")
            os.makedirs(site_dir)
            with open(
                os.path.join(site_dir, "site_metadata.json"), "w", encoding="utf-8"
            ) as f:
                json.dump({"Url": "https://sp/sites/team"}, f)

            with (
                mock.patch.dict(os.environ, {"SPX_OUTPUT_DIR": output_dir}),
                mock.patch.object(
                    sys, "argv", ["inventory_site_pages.py", "root", "team"]
                ),
                mock.patch.object(
                    isp, "get_config", return_value=("https://sp", "u", "p", None)
                ),
                mock.patch.object(isp, "build_context", return_value=ctx),
                mock.patch.object(isp, "fetch_versions", return_value={"history": []}),
                mock.patch("builtins.print"),
            ):
                isp.main()

            with open(
                os.path.join(site_dir, "pages", "pages_inventory.json"),
                encoding="utf-8",
            ) as f:
                inventory = json.load(f)

        self.assertEqual(len(inventory["pages"]), 601)
        self.assertEqual(inventory["errors"], [])


if __name__ == "__main__":
    unittest.main()
