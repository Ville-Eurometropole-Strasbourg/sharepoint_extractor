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
    import extract_metadata as em


def column(name, type_, hidden=False):
    return {"internal_name": name, "type": type_, "hidden": hidden}


class ExtraListFieldsTest(unittest.TestCase):
    def test_person_and_lookup_ids(self):
        schema = [
            column("Title", "Text"),
            column("AssignedTo", "User"),
            column("Participants", "UserMulti"),
            column("Projet", "Lookup"),
            column("Predecessors", "LookupMulti"),
            column("Author", "User"),
            column("Editor", "User"),
            column("MyAuthor", "User"),
            column("ItemChildCount", "Lookup"),
            column("HiddenPerson", "User", hidden=True),
        ]
        self.assertEqual(
            em.extra_list_fields(schema),
            ["AssignedToId", "ParticipantsId", "ProjetId", "PredecessorsId"],
        )

    def test_calendar_recurrence_fields_present_in_schema(self):
        schema = [
            column("fRecurrence", "Recurrence", hidden=True),
            column("RecurrenceID", "DateTime", hidden=True),
            column("EventType", "Integer", hidden=True),
            column("RecurrenceData", "Note", hidden=True),
        ]
        self.assertEqual(
            em.extra_list_fields(schema),
            ["RecurrenceData", "EventType", "RecurrenceID"],
        )

    def test_recurrence_fields_need_frecurrence(self):
        schema = [column("RecurrenceData", "Note", hidden=True)]
        self.assertEqual(em.extra_list_fields(schema), [])

    def test_empty_schema(self):
        self.assertEqual(em.extra_list_fields(None), [])
        self.assertEqual(em.extra_list_fields([]), [])


class NormalizeIndexedValuesTest(unittest.TestCase):
    def test_integer_keys_as_returned_by_the_client(self):
        record = {"AssignedToId": {1: 34, 0: 12}}
        self.assertEqual(
            em._normalize_indexed_values(record), {"AssignedToId": [12, 34]}
        )

    def test_string_keys_sorted_numerically(self):
        record = {"Choix": {"10": "c", "2": "b", "0": "a"}}
        self.assertEqual(
            em._normalize_indexed_values(record), {"Choix": ["a", "b", "c"]}
        )

    def test_other_values_untouched(self):
        record = {
            "Empty": {},
            "Mixed": {"0": "x", "Title": "y"},
            "Plain": {"Title": "y"},
            "Number": 3,
            "Text": "0",
            "List": [1, 2],
            "Nothing": None,
        }
        expected = {key: value for key, value in record.items()}
        self.assertEqual(em._normalize_indexed_values(record), expected)

    def test_modifies_and_returns_the_same_record(self):
        record = {"Ids": {0: 1}}
        self.assertIs(em._normalize_indexed_values(record), record)
        self.assertEqual(record["Ids"], [1])


class ExtractBadFieldNameTest(unittest.TestCase):
    def test_french_error_message(self):
        message = "Il n'existe pas de champ ni de propriété « _CopySource »."
        self.assertEqual(em._extract_bad_field_name(message), "_CopySource")

    def test_english_error_message(self):
        message = 'Field or property "_CopySource" does not exist.'
        self.assertEqual(em._extract_bad_field_name(message), "_CopySource")

    def test_no_field_name(self):
        self.assertIsNone(em._extract_bad_field_name("400 Bad Request"))
        self.assertIsNone(em._extract_bad_field_name(None))


def sp_list(entity_name, template=100, hidden=False):
    return {
        "Title": "Any title",
        "EntityTypeName": entity_name,
        "BaseTemplate": template,
        "Hidden": hidden,
    }


class ListSkipReasonTest(unittest.TestCase):
    def test_business_lists_are_kept(self):
        self.assertIsNone(em.list_skip_reason(sp_list("Shared_x0020_Documents", 101)))
        self.assertIsNone(em.list_skip_reason(sp_list("TasksList", 171)))

    def test_system_lists_whatever_their_title(self):
        self.assertEqual(em.list_skip_reason(sp_list("SitePages", 119)), "system_list")
        self.assertEqual(
            em.list_skip_reason(sp_list("Style_x0020_Library", 101)), "system_list"
        )

    def test_hidden_lists(self):
        self.assertEqual(
            em.list_skip_reason(sp_list("UserInfo", 112, hidden=True)), "hidden"
        )
        for template in (117, 140, 10102):
            self.assertIsNone(
                em.list_skip_reason(sp_list("Workflows", template, hidden=True))
            )


class PersonColumnsTest(unittest.TestCase):
    def test_author_editor_first_without_duplicates(self):
        schema = [
            column("Author", "User"),
            column("AssignedTo", "User"),
            column("Participants", "UserMulti"),
            column("MyEditor", "User"),
            column("Hidden", "User", hidden=True),
            column("Projet", "Lookup"),
        ]
        self.assertEqual(
            em.person_columns(schema),
            ["Author", "Editor", "AssignedTo", "Participants"],
        )

    def test_without_schema(self):
        self.assertEqual(em.person_columns(None), ["Author", "Editor"])


class PersonTitlesTest(unittest.TestCase):
    def test_single_person(self):
        self.assertEqual(em._person_titles({"Title": "Personne A"}), "Personne A")

    def test_list_of_people(self):
        value = [{"Title": "Personne A"}, {"Title": "Personne B"}]
        self.assertEqual(em._person_titles(value), ["Personne A", "Personne B"])

    def test_indexed_people(self):
        value = {1: {"Title": "Personne B"}, 0: {"Title": "Personne A"}}
        self.assertEqual(em._person_titles(value), ["Personne A", "Personne B"])

    def test_unknown_shapes(self):
        self.assertIsNone(em._person_titles(None))
        self.assertIsNone(em._person_titles("Personne A"))
        self.assertIsNone(em._person_titles({"Id": 12}))
        self.assertIsNone(em._person_titles({}))


PATH_ROOT = "https://sp.example.org/sites/hr"
HOST_ROOT = "https://intranet.example.org"


def discovered(path, title=None, status="ok"):
    return {"server_relative_url": path, "title": title, "status": status}


class BuildSitesFromDiscoveryTest(unittest.TestCase):
    def test_path_root_is_kept_and_paths_stay_server_relative(self):
        sites = em.build_sites_from_discovery(
            [
                discovered("/sites/hr", "Root"),
                discovered("/sites/hr/team", "Team"),
                discovered("/sites/hr/broken", status="error"),
            ],
            PATH_ROOT,
        )
        self.assertEqual(
            sites,
            [
                {"slug": "_root", "url_suffix": "/sites/hr", "label": "Root"},
                {"slug": "team", "url_suffix": "/sites/hr/team", "label": "Team"},
            ],
        )

    def test_host_root_is_kept_in_a_non_empty_folder(self):
        sites = em.build_sites_from_discovery(
            [discovered("/"), discovered("/service/team")], HOST_ROOT
        )
        self.assertEqual(
            [(s["slug"], s["url_suffix"]) for s in sites],
            [("_root", "/"), ("team", "/service/team")],
        )

    def test_root_is_matched_ignoring_case_and_trailing_slash(self):
        sites = em.build_sites_from_discovery([discovered("/Sites/HR/")], PATH_ROOT)
        self.assertEqual([s["slug"] for s in sites], ["_root"])

    def test_a_subsite_named_like_the_root_keeps_its_folder(self):
        # The root does not take its last segment, so /sites/hr/x/hr stays "hr".
        sites = em.build_sites_from_discovery(
            [discovered("/sites/hr"), discovered("/sites/hr/x/hr")], PATH_ROOT
        )
        self.assertEqual([s["slug"] for s in sites], ["_root", "hr"])

    def test_a_subsite_named_root_folder_uses_its_full_path(self):
        sites = em.build_sites_from_discovery(
            [discovered("/"), discovered("/service/_root")], HOST_ROOT
        )
        self.assertEqual([s["slug"] for s in sites], ["_root", "service-_root"])


class SiteSlugTest(unittest.TestCase):
    def test_root_site(self):
        self.assertEqual(em.site_slug("/", HOST_ROOT), "_root")
        self.assertEqual(em.site_slug("/sites/hr", PATH_ROOT), "_root")
        self.assertEqual(
            em.site_slug("/sites/hr/SitePages/Home.aspx", PATH_ROOT), "_root"
        )

    def test_other_sites(self):
        self.assertEqual(em.site_slug("/sites/hr/team", PATH_ROOT), "team")
        self.assertEqual(em.site_slug("/sites/hr", HOST_ROOT), "hr")


class _Query:
    """A chainable client query whose execute_query() returns `result`."""

    def __init__(self, result):
        self.result, self.selected = result, None

    def select(self, fields):
        self.selected = list(fields)
        return self

    def get(self):
        return self

    def get_all(self, *args, **kwargs):
        return self

    def execute_query(self):
        return self.result


LIBRARY_ROOT = "/sites/team/Marches"
FOLDER_CT = "0x0120002A"
DOCUMENT_SET_CT = "0x0120D520001"
FILE_CT = "0x0101001B"


def _library_item(item_id, relative, content_type, folder=False):
    return types.SimpleNamespace(
        properties={
            "Id": item_id,
            "FileRef": f"{LIBRARY_ROOT}/{relative}",
            "FileLeafRef": relative.rsplit("/", 1)[-1],
            "FileSystemObjectType": 1 if folder else 0,
            "ContentTypeId": content_type,
            "Lot": "A",
        }
    )


class FakeLibrary:
    def __init__(self, items):
        self.fields = _Query(
            [
                types.SimpleNamespace(
                    properties={
                        "Title": "Lot",
                        "InternalName": "Lot",
                        "TypeAsString": "Text",
                        "Hidden": False,
                    }
                )
            ]
        )
        self.items = _Query(items)
        self.root_folder = _Query(
            types.SimpleNamespace(properties={"ServerRelativeUrl": LIBRARY_ROOT})
        )


class DocumentLibraryFoldersTest(unittest.TestCase):
    """Folders and document sets get a .meta.json next to their directory."""

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.output_dir = Path(tmp.name)
        self.library = FakeLibrary(
            [
                _library_item(1, "2024", FOLDER_CT, folder=True),
                _library_item(2, "2024/Lot 1", DOCUMENT_SET_CT, folder=True),
                _library_item(3, "2024/Lot 1/offre.pdf", FILE_CT),
                _library_item(4, "Vide", FOLDER_CT, folder=True),
                _library_item(5, "note.docx", FILE_CT),
            ]
        )
        self.result = em.extract_document_library(
            self.library, "Marches", str(self.output_dir), mock.Mock()
        )
        self.lib_dir = self.output_dir / "documents" / "Marches"

    def meta(self, relative):
        path = self.lib_dir / (relative + ".meta.json")
        return json.loads(path.read_text(encoding="utf-8"))

    def test_every_folder_and_document_set_has_its_meta(self):
        self.assertEqual(self.meta("2024")["ContentTypeId"], FOLDER_CT)
        document_set = self.meta("2024/Lot 1")
        self.assertEqual(document_set["ContentTypeId"], DOCUMENT_SET_CT)
        self.assertEqual(document_set["Lot"], "A")
        self.assertEqual(document_set["FileSystemObjectType"], 1)
        self.assertEqual(document_set["_relative_path"], "2024/Lot 1")

    def test_empty_folders_exist_on_disk(self):
        self.assertTrue((self.lib_dir / "Vide").is_dir())
        self.assertEqual(list((self.lib_dir / "Vide").iterdir()), [])

    def test_files_are_written_as_before(self):
        self.assertEqual(self.meta("2024/Lot 1/offre.pdf")["Id"], 3)
        self.assertEqual(self.meta("note.docx")["Id"], 5)
        self.assertEqual(self.result, {"files": 2, "folders": 3})

    def test_the_content_type_is_read(self):
        self.assertIn("ContentTypeId", self.library.items.selected)


class LibraryStateFormatTest(unittest.TestCase):
    """Libraries extracted before folders had a .meta.json are read again once."""

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.output_dir = tmp.name

    def write_state(self, kind, state):
        path = Path(em._state_path(self.output_dir, kind, "Docs"))
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(state), encoding="utf-8")

    def test_an_older_library_state_is_re_extracted(self):
        self.write_state("library", {"last_item_modified_date": "2026-01-01"})
        self.assertIsNone(em._read_last_extracted(self.output_dir, "library", "Docs"))

    def test_a_current_library_state_is_skipped(self):
        em._write_last_extracted(self.output_dir, "library", "Docs", "2026-01-01")
        self.assertEqual(
            em._read_last_extracted(self.output_dir, "library", "Docs"), "2026-01-01"
        )

    def test_list_states_keep_their_own_format(self):
        self.write_state(
            "list",
            {"last_item_modified_date": "2026-01-01", "format": em.LIBRARY_FORMAT},
        )
        self.assertIsNone(em._read_last_extracted(self.output_dir, "list", "Docs"))
        em._write_last_extracted(self.output_dir, "list", "Docs", "2026-01-01")
        self.assertEqual(
            em._read_last_extracted(self.output_dir, "list", "Docs"), "2026-01-01"
        )


if __name__ == "__main__":
    unittest.main()
