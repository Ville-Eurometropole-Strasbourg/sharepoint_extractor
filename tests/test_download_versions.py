import io
import json
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock

import requests
from office365.sharepoint.client_context import ClientContext

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
    import download_versions as dv


SITE_URL = "https://sp.example/sites/team"
PAGE = "/sites/team/Pages/default.aspx"
PAGE_VERSION = "_vti_history/512/Pages/default.aspx"
PAGE_VTI_URL = SITE_URL + "/_vti_history/512/Pages/default.aspx"
PAGE_REST_URL = (
    SITE_URL + "/_api/Web/getFileByServerRelativeUrl("
    "'%2Fsites%2Fteam%2FPages%2Fdefault.aspx')/versions/getById(512)/$value"
)


class _Server:
    """Answers like SharePoint: content at one URL, the status `elsewhere`
    everywhere else."""

    def __init__(self, url, content, elsewhere):
        self.url = url
        self.content = content
        self.elsewhere = elsewhere
        self.requested = []

    def execute_request_direct(self, request):
        self.requested.append(request.url)
        response = requests.Response()
        response.url = request.url
        if request.url == self.url:
            response.status_code = 200
            response.raw = io.BytesIO(self.content)
        else:
            response.status_code = self.elsewhere
        response.raise_for_status()
        return response


class DownloadSiteVersionsTest(unittest.TestCase):
    def download(self, file_ref, version_url, served_url, content, elsewhere=404):
        """Download version 1.0 of file_ref, listed at version_url, from a
        server that serves content at served_url and answers `elsewhere`
        everywhere else. Returns the server, the counts, and the bytes
        written (None when no file was written)."""
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        site_metadata = os.path.join(tmp.name, "site_metadata.json")
        with open(site_metadata, "w", encoding="utf-8") as f:
            json.dump({"Url": SITE_URL}, f)
        library_path = file_ref[len("/sites/team/") :]
        versions_path = os.path.join(tmp.name, "documents", library_path)
        versions_path += ".versions.json"
        os.makedirs(os.path.dirname(versions_path))
        with open(versions_path, "w", encoding="utf-8") as f:
            history = [{"label": "1.0", "size": len(content), "url": version_url}]
            json.dump({"file_ref": file_ref, "history": history}, f)

        ctx = ClientContext(SITE_URL)
        server = _Server(served_url, content, elsewhere)
        with (
            mock.patch.object(dv, "build_context", return_value=ctx),
            mock.patch.object(
                ctx.pending_request(),
                "execute_request_direct",
                side_effect=server.execute_request_direct,
            ),
        ):
            counts = dv.download_site_versions(tmp.name, "u", "p", None, mock.Mock())

        written = dv._history_path(versions_path, "1.0")
        if not os.path.exists(written):
            return server, counts, None
        with open(written, "rb") as f:
            return server, counts, f.read()

    def test_a_page_version_is_read_through_rest_after_a_404(self):
        # SharePoint 2019 answers 404 on the _vti_history URL of a page.
        server, counts, written = self.download(
            file_ref=PAGE,
            version_url=PAGE_VERSION,
            served_url=PAGE_REST_URL,
            content=b"<html/>",
        )
        self.assertEqual(written, b"<html/>")
        self.assertEqual((counts["downloaded"], counts["errors"]), (1, 0))
        self.assertEqual(server.requested, [PAGE_VTI_URL, PAGE_REST_URL])

    def test_a_version_served_by_vti_history_is_read_there_only(self):
        vti = SITE_URL + "/_vti_history/512/Shared%20Documents/a.docx"
        server, counts, written = self.download(
            file_ref="/sites/team/Shared Documents/a.docx",
            version_url="_vti_history/512/Shared Documents/a.docx",
            served_url=vti,
            content=b"PK\x03\x04doc",
        )
        self.assertEqual(written, b"PK\x03\x04doc")
        self.assertEqual(server.requested, [vti])

    def test_an_error_other_than_404_is_recorded_without_a_rest_attempt(self):
        server, counts, written = self.download(
            file_ref=PAGE,
            version_url=PAGE_VERSION,
            served_url=PAGE_REST_URL,
            content=b"<html/>",
            elsewhere=500,
        )
        self.assertIsNone(written)
        self.assertEqual((counts["downloaded"], counts["errors"]), (0, 1))
        self.assertEqual(server.requested, [PAGE_VTI_URL])


if __name__ == "__main__":
    unittest.main()
