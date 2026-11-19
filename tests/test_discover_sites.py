import sys
import types
import unittest
from pathlib import Path
from unittest import mock
from xml.sax.saxutils import escape

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src" / "discovery"))


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
        "requests_ntlm": _stub("requests_ntlm", "HttpNtlmAuth"),
        "sharepoint_client": _stub("sharepoint_client", "get_config", "build_context"),
        "persistent_session": _stub(
            "persistent_session", "enable_persistent_session", "get_shared_session"
        ),
    },
):
    import discover_sites as ds


def soap_response(inner_xml, last_item=None):
    last = (
        f"<lastItemIdOnPage>{escape(last_item)}</lastItemIdOnPage>"
        if last_item is not None
        else ""
    )
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'
        f'<soap:Body><GetContentResponse xmlns="{ds.SOAP_NS}">'
        f"<GetContentResult>{escape(inner_xml)}</GetContentResult>{last}"
        "</GetContentResponse></soap:Body></soap:Envelope>"
    ).encode("utf-8")


class ParseGetContentTest(unittest.TestCase):
    def test_virtual_server(self):
        inner = (
            '<VirtualServer><Metadata URL="https://sp.example.org/" />'
            '<ContentDatabases><ContentDatabase ID="{A}" />'
            '<ContentDatabase ID="{B}" /></ContentDatabases></VirtualServer>'
        )
        element, last = ds.parse_get_content(soap_response(inner))
        self.assertEqual(ds.content_database_ids(element), ["{A}", "{B}"])
        self.assertIsNone(last)

    def test_content_database_with_paging(self):
        inner = (
            '<ContentDatabase><Sites><Site URL="https://sp.example.org/sites/hr" '
            'ID="{1}" /><Site URL="https://sp.example.org/sites/it" ID="{2}" />'
            "</Sites></ContentDatabase>"
        )
        element, last = ds.parse_get_content(soap_response(inner, "{2}"))
        self.assertEqual(
            ds.site_collection_urls(element),
            ["https://sp.example.org/sites/hr", "https://sp.example.org/sites/it"],
        )
        self.assertEqual(last, "{2}")

    def test_empty_last_item_is_none(self):
        _, last = ds.parse_get_content(soap_response("<ContentDatabase />", ""))
        self.assertIsNone(last)

    def test_missing_result(self):
        with self.assertRaises(ValueError):
            ds.parse_get_content(
                b'<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/'
                b'envelope/"><soap:Body /></soap:Envelope>'
            )


class EnvelopeTest(unittest.TestCase):
    def test_escapes_values(self):
        body = ds.get_content_envelope("ContentDatabase", "{A}", "a<b&c")
        self.assertIn("<objectId>{A}</objectId>", body)
        self.assertIn("<lastItemIdOnPage>a&lt;b&amp;c</lastItemIdOnPage>", body)
        self.assertNotIn("lastItemIdOnPage", ds.get_content_envelope("VirtualServer"))


class CollectionsUnderRootTest(unittest.TestCase):
    URLS = [
        "https://intranet.example.org/",
        "https://intranet.example.org/projets/alpha",
        "https://INTRANET.example.org/projets/Alpha/",
        "https://intranet.example.org/equipe/beta",
        "https://other.example.org/sites/gamma",
        "https://intranet.example.org/sites/hr",
        "https://intranet.example.org/sites/hrx",
    ]

    def test_host_root_keeps_whole_host(self):
        self.assertEqual(
            ds.collections_under_root(self.URLS, "https://intranet.example.org"),
            ["/equipe/beta", "/projets/alpha", "/sites/hr", "/sites/hrx"],
        )

    def test_path_root_keeps_only_its_path(self):
        urls = self.URLS + ["https://intranet.example.org/sites/hr/sub"]
        self.assertEqual(
            ds.collections_under_root(urls, "https://intranet.example.org/sites/hr/"),
            ["/sites/hr/sub"],
        )


if __name__ == "__main__":
    unittest.main()
