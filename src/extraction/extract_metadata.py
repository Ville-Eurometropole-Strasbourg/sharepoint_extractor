"""
Extract document and list metadata from SharePoint sites.

Usage:
    python extract_metadata.py <site_url_suffix>   # e.g. /sites/hr/team
    python extract_metadata.py --all-from-discovery <sites_index_<root>.json>
    python extract_metadata.py --permissions-all

Notes:
- Libraries vs lists are split on BaseType (1 = DocumentLibrary).
- Each list gets a _schema_columns.json.
- FileLeafRef/FileRef/FileSystemObjectType must be selected explicitly.
- One .meta.json per file.
- Hidden lists are skipped, except KEEP_DESPITE_HIDDEN_TEMPLATES.
- System lists are matched on EntityTypeName and BaseTemplate, which do not
  depend on the site language (titles do).
"""

import os
import sys
import json
import logging
import re
import subprocess

from sharepoint_client import get_config, build_context
from site_urls import absolute_url, get_root_name, root_path
from persistent_session import enable_persistent_session

GENERIC_LIST_FORMAT = 4
PERSON_TYPES = {"User", "UserMulti"}
PERSON_EXCLUDED = {"AppAuthor", "AppEditor", "MyAuthor", "MyEditor"}

# System lists with no business content, by EntityTypeName. It is built from
# the list URL, which SharePoint does not translate (spaces become _x0020_,
# lists under /Lists/ get a "List" suffix, lists at the web root do not).
SKIP_LIST_ENTITY_NAMES = {
    "Style_x0020_Library",
    "FormServerTemplates",
    "SitePages",
    "Reports_x0020_List",  # Content and Structure Reports (publishing)
    "ReusableContent",  # Reusable Content (publishing)
    "SiteCollectionDocuments",
    "SiteCollectionImages",
}

# Hidden lists that still hold business content, by BaseTemplate.
KEEP_DESPITE_HIDDEN_TEMPLATES = {
    117,  # NoCodeWorkflows: workflow definitions
    140,  # WorkflowHistory
    10102,  # IWConvertedForms: converted InfoPath forms
}

LIST_FILTER_FIELDS = ["Title", "Hidden", "BaseTemplate", "EntityTypeName"]

# permissions.json format. 2 adds items_with_broken_inheritance and
# items_scan_errors; older files are redone by run_permissions_pass().
PERMISSIONS_FORMAT = 2
PERMISSION_LIST_FIELDS = LIST_FILTER_FIELDS + [
    "Id",
    "BaseType",
    "ItemCount",
    "HasUniqueRoleAssignments",
]
# Items are only scanned for this flag: one paged read per list, and role
# assignments are then read for the flagged items only.
ITEM_PERMISSION_FIELDS = [
    "Id",
    "FileRef",
    "FileSystemObjectType",
    "HasUniqueRoleAssignments",
]
ITEM_PERMISSIONS_PAGE_SIZE = 1000


def list_skip_reason(props):
    """None to extract the list, else "hidden" or "system_list"."""
    if (
        props.get("Hidden")
        and props.get("BaseTemplate") not in KEEP_DESPITE_HIDDEN_TEMPLATES
    ):
        return "hidden"
    if props.get("EntityTypeName") in SKIP_LIST_ENTITY_NAMES:
        return "system_list"
    return None


def reads_permissions(props):
    """
    Whether a list's permissions and those of its items are read: every list
    but the hidden ones list_skip_reason() skips. System lists are kept, as
    their content is skipped (or read elsewhere, like pages) but not their
    permissions.
    """
    return not (
        props.get("Hidden")
        and props.get("BaseTemplate") not in KEEP_DESPITE_HIDDEN_TEMPLATES
    )


# OData__UIVersionString = _UIVersionString (current version label). Anything
# above "1.0", or "1.0" in a library with minor versions, may have a /versions
# history.
ITEM_SELECT_FIELDS_BASE = [
    "FileLeafRef",
    "FileRef",
    "FileSystemObjectType",
    "Id",
    "GUID",
    "FileSizeDisplay",
    "OData__UIVersionString",
]

MANY_CUSTOM_COLUMNS_WARNING_THRESHOLD = 40


def extract_list_schema(sp_list, logger):
    """Read a list's field definitions raw, for _schema_columns.json."""
    try:
        fields = (
            sp_list.fields.select(
                ["Title", "InternalName", "TypeAsString", "Hidden", "Required"]
            )
            .get()
            .execute_query()
        )
        return [
            {
                "title": f.properties.get("Title"),
                "internal_name": f.properties.get("InternalName"),
                "type": f.properties.get("TypeAsString"),
                "hidden": f.properties.get("Hidden"),
                "required": f.properties.get("Required"),
            }
            for f in fields
        ]
    except Exception as e:
        logger.warning(f"    Failed to read schema: {e}")
        return None


def discover_custom_columns(schema):
    """
    Likely custom columns: not hidden, not a standard system field.
    Types that need $expand (User, Lookup...) are excluded.
    """
    if not schema:
        return []
    standard_fields = {
        "ContentType",
        "Title",
        "Modified",
        "Created",
        "CreatedBy",
        "ModifiedBy",
        "Author",
        "Editor",
        "_UIVersionString",
        "Attachments",
        "Edit",
        "LinkTitleNoMenu",
        "LinkTitle",
        "DocIcon",
        "ItemChildCount",
        "FolderChildCount",
        "AppAuthor",
        "AppEditor",
        "FileLeafRef",
        "FileRef",
        "FileSystemObjectType",
        "ID",
        "GUID",
        "ComplianceAssetId",
    }
    expand_required_types = {
        "User",
        "UserMulti",
        "Lookup",
        "LookupMulti",
        "ModStat",
        "Computed",
    }
    return [
        f
        for f in schema
        if not f.get("hidden")
        and f.get("internal_name") not in standard_fields
        and f.get("type") not in expand_required_types
    ]


ID_FIELD_TYPES = {"User", "UserMulti", "Lookup", "LookupMulti"}
ID_FIELD_EXCLUDED = {
    "Author",
    "Editor",
    "AppAuthor",
    "AppEditor",
    "ItemChildCount",
    "FolderChildCount",
    "MyAuthor",
    "MyEditor",
}
RECURRENCE_FIELDS = [
    "RecurrenceData",
    "EventType",
    "MasterSeriesItemID",
    "RecurrenceID",
]


def extra_list_fields(schema):
    """<Name>Id of visible user/lookup columns, plus calendar recurrence fields."""
    if not schema:
        return []
    names = {f.get("internal_name") for f in schema}
    fields = [
        f"{f['internal_name']}Id"
        for f in schema
        if not f.get("hidden")
        and f.get("type") in ID_FIELD_TYPES
        and f.get("internal_name") not in ID_FIELD_EXCLUDED
    ]
    if "fRecurrence" in names:
        fields += [name for name in RECURRENCE_FIELDS if name in names]
    return fields


QUOTED_NAME = re.compile(r'«\s*([^»]+?)\s*»|"([^"]+)"|“([^”]+)”')


def _extract_bad_field_name(error_message):
    """
    First quoted name in a SharePoint "field does not exist" error, in any
    language: « _CopySource » (fr), "_CopySource" (en), “_CopySource”.
    None if absent.
    """
    match = QUOTED_NAME.search(str(error_message))
    if not match:
        return None
    return next(group for group in match.groups() if group)


def extract_document_library(sp_list, title, output_dir, logger):
    """Extract a library's schema and one .meta.json per file (no binaries)."""
    lib_dir = os.path.join(output_dir, "documents", title.replace("/", "_"))
    os.makedirs(lib_dir, exist_ok=True)

    schema = extract_list_schema(sp_list, logger)
    custom_columns = discover_custom_columns(schema)
    with open(
        os.path.join(lib_dir, "_schema_columns.json"), "w", encoding="utf-8"
    ) as f:
        json.dump(
            {"all_fields": schema, "likely_custom_columns": custom_columns},
            f,
            ensure_ascii=False,
            indent=2,
            default=str,
        )

    select_fields = ITEM_SELECT_FIELDS_BASE + [
        c["internal_name"] for c in custom_columns
    ]
    select_fields = list(dict.fromkeys(select_fields))

    if len(select_fields) > MANY_CUSTOM_COLUMNS_WARNING_THRESHOLD:
        logger.warning(
            f"    '{title}' has {len(select_fields)} columns to select "
            f"(threshold {MANY_CUSTOM_COLUMNS_WARNING_THRESHOLD}): the resulting "
            f"$select URL may be too long for the server."
        )

    items, dropped_columns = None, []
    remaining_fields = list(select_fields)
    max_attempts = len(select_fields) + 1
    for attempt in range(max_attempts):
        try:
            items = (
                sp_list.items.select(remaining_fields)
                .get_all(page_size=100)
                .execute_query()
            )
            break
        except Exception as e:
            bad_field = _extract_bad_field_name(getattr(e, "message", None) or str(e))
            if bad_field and bad_field in remaining_fields:
                logger.warning(
                    f"    Field '{bad_field}' rejected by server, dropping and retrying."
                )
                remaining_fields.remove(bad_field)
                dropped_columns.append(bad_field)
                continue

            too_long = len(remaining_fields) > MANY_CUSTOM_COLUMNS_WARNING_THRESHOLD
            if not too_long:
                # No field name to parse (client-side deserializer crash):
                # bisect the custom columns to find the culprit.
                base_present = [
                    f for f in remaining_fields if f in ITEM_SELECT_FIELDS_BASE
                ]
                custom_present = [
                    f for f in remaining_fields if f not in ITEM_SELECT_FIELDS_BASE
                ]
                logger.warning(
                    f"    Could not isolate offending field from the error message, "
                    f"bisecting {len(custom_present)} custom column(s) to find it..."
                )
                good_custom, bad_custom = _bisect_bad_fields(
                    sp_list, base_present, custom_present, logger
                )
                if bad_custom:
                    remaining_fields = base_present + good_custom
                    dropped_columns.extend(bad_custom)
                    continue

            error_category = "likely_url_too_long" if too_long else "unrecoverable"
            logger.error(
                f"    Failed to read items (could not isolate offending field): {e}"
            )
            return {
                "files": 0,
                "folders": 0,
                "error": str(e),
                "error_category": error_category,
            }

    if items is None:
        logger.error(f"    Exhausted retry attempts without success for '{title}'.")
        return {"files": 0, "folders": 0, "error": "exhausted retries"}

    library_root = None
    try:
        library_root = (
            sp_list.root_folder.get()
            .execute_query()
            .properties.get("ServerRelativeUrl")
        )
    except Exception as e:
        logger.warning(
            f"    Could not read the library root folder, writing .meta.json flat: {e}"
        )

    file_count, folder_count, flat_fallback = 0, 0, 0
    for item in items:
        props = item.properties
        is_folder = props.get("FileSystemObjectType") == 1
        name = (
            props.get("FileLeafRef") or props.get("Title") or f"item_{props.get('Id')}"
        )

        if is_folder:
            folder_count += 1
            continue

        file_count += 1
        relative = relative_path_in_library(props.get("FileRef"), library_root)
        if relative is None:
            relative = name.replace("/", "_")
            flat_fallback += 1
        props["_relative_path"] = relative
        meta_path = os.path.join(lib_dir, *relative.split("/")) + ".meta.json"
        os.makedirs(os.path.dirname(meta_path), exist_ok=True)
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(dict(props), f, ensure_ascii=False, indent=2, default=str)

    result = {"files": file_count, "folders": folder_count}
    if flat_fallback:
        result["meta_flat_fallback"] = flat_fallback
    if dropped_columns:
        result["dropped_columns"] = dropped_columns
    return result


def _bisect_bad_fields(sp_list, base_fields, custom_fields, logger):
    """
    Find the custom fields that crash the client's deserializer
    ("'int' object has no attribute 'get'"), which names no field, by
    recursively bisecting custom_fields.

    Returns (good_custom_fields, bad_fields).
    """
    if not custom_fields:
        return [], []

    try:
        sp_list.items.select(base_fields + custom_fields).get_all(
            page_size=100
        ).execute_query()
        return custom_fields, []
    except Exception:
        if len(custom_fields) == 1:
            logger.warning(
                f"    Field '{custom_fields[0]}' triggers a deserialization crash, dropping."
            )
            return [], custom_fields
        mid = len(custom_fields) // 2
        left_good, left_bad = _bisect_bad_fields(
            sp_list, base_fields, custom_fields[:mid], logger
        )
        right_good, right_bad = _bisect_bad_fields(
            sp_list, base_fields, custom_fields[mid:], logger
        )
        return left_good + right_good, left_bad + right_bad


def _is_index_key(key):
    return isinstance(key, int) or (isinstance(key, str) and key.isdigit())


def _normalize_indexed_values(record):
    """Turn {0: a, 1: b} values (multi-value fields) into [a, b]."""
    for key, value in record.items():
        if isinstance(value, dict) and value and all(_is_index_key(k) for k in value):
            record[key] = [value[k] for k in sorted(value, key=int)]
    return record


def person_columns(schema):
    """Author, Editor and the visible person columns of a list."""
    names = [
        f["internal_name"]
        for f in schema or []
        if not f.get("hidden")
        and f.get("type") in PERSON_TYPES
        and f.get("internal_name") not in PERSON_EXCLUDED
    ]
    return list(dict.fromkeys(["Author", "Editor"] + names))


def _person_titles(value):
    """Display name, or list of names, from an expanded person field."""
    if isinstance(value, list):
        return [_person_titles(v) for v in value]
    if not isinstance(value, dict):
        return None
    if "Title" in value:
        return value.get("Title")
    if value and all(_is_index_key(k) for k in value):
        return [_person_titles(value[k]) for k in sorted(value, key=int)]
    return None


def _read_person_titles(sp_list, columns):
    items = (
        sp_list.items.select(["Id"] + [f"{c}/Title" for c in columns])
        .expand(columns)
        .get_all(page_size=100)
        .execute_query()
    )
    return {item.properties.get("Id"): item.properties for item in items}


def add_person_names(sp_list, records, schema, logger):
    """
    Add <Column> = display name(s) next to <Column>Id, read through the items
    so that former users keep their name. Returns the columns that failed.
    """
    columns = person_columns(schema)
    if not records or not columns:
        return []
    try:
        by_id = _read_person_titles(sp_list, columns)
        readable, failed = columns, []
    except Exception as e:
        logger.warning(f"    Person names read failed ({e}), retrying per column.")
        by_id, readable, failed = {}, [], []
        for column in columns:
            try:
                for item_id, props in _read_person_titles(sp_list, [column]).items():
                    by_id.setdefault(item_id, {})[column] = props.get(column)
                readable.append(column)
            except Exception:
                failed.append(column)
        if failed:
            logger.warning(f"    Person names not read for {failed}, ids kept.")
    for record in records:
        props = by_id.get(record.get("Id"), {})
        for column in readable:
            if column in props:
                record[column] = _person_titles(props[column])
    return failed


def extract_generic_list(sp_list, title, output_dir, logger):
    """Extract a generic list: schema + all items."""
    list_dir = os.path.join(output_dir, "lists", title.replace("/", "_"))
    os.makedirs(list_dir, exist_ok=True)

    schema = extract_list_schema(sp_list, logger)
    custom_columns = discover_custom_columns(schema)
    with open(
        os.path.join(list_dir, "_schema_columns.json"), "w", encoding="utf-8"
    ) as f:
        json.dump(schema, f, ensure_ascii=False, indent=2, default=str)

    base_fields = [
        "Id",
        "Title",
        "Created",
        "Modified",
        "AuthorId",
        "EditorId",
    ]
    custom_field_names = list(
        dict.fromkeys(
            [c["internal_name"] for c in custom_columns] + extra_list_fields(schema)
        )
    )
    try:
        items = (
            sp_list.items.select(base_fields + custom_field_names)
            .get_all(page_size=100)
            .execute_query()
        )
        records = [dict(item.properties) for item in items]
        dropped_fields = []
    except Exception as e:
        logger.warning(
            f"    Read failed ({e}), bisecting fields to isolate the cause..."
        )
        good_fields, dropped_fields = _bisect_bad_fields(
            sp_list, base_fields, custom_field_names, logger
        )
        try:
            items = (
                sp_list.items.select(base_fields + good_fields)
                .get_all(page_size=100)
                .execute_query()
            )
            records = [dict(item.properties) for item in items]
        except Exception as e2:
            logger.error(f"    Still failing after dropping {dropped_fields}: {e2}")
            return {"items": 0, "error": str(e2)}

    records = [_normalize_indexed_values(r) for r in records]
    failed_person_names = add_person_names(sp_list, records, schema, logger)
    with open(os.path.join(list_dir, "items.json"), "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2, default=str)

    result = {"items": len(records)}
    if failed_person_names:
        result["person_names_failed"] = failed_person_names
    if dropped_fields:
        result["dropped_columns"] = dropped_fields
    return result


def extract_site(
    site_url, slug, username, password, ca_bundle_path, base_output_dir, logger
):
    logger.info(f"--- Extracting site: {slug} ({site_url}) ---")
    output_dir = os.path.join(base_output_dir, slug)
    os.makedirs(output_dir, exist_ok=True)

    try:
        ctx = build_context(site_url, username, password, ca_bundle_path)
        web = ctx.web.get().execute_query()
    except Exception as e:
        logger.error(f"Failed to connect to {site_url}: {e}")
        return {"error": str(e)}

    with open(
        os.path.join(output_dir, "site_metadata.json"), "w", encoding="utf-8"
    ) as f:
        json.dump(dict(web.properties), f, ensure_ascii=False, indent=2, default=str)

    try:
        lists = (
            ctx.web.lists.select(
                [
                    "Id",
                    "BaseType",
                    "ItemCount",
                    "LastItemModifiedDate",
                ]
                + LIST_FILTER_FIELDS
            )
            .get()
            .execute_query()
        )
    except Exception as e:
        logger.error(f"Failed to list lists/libraries: {e}")
        return {"error": str(e)}

    summary = {"document_libraries": {}, "generic_lists": {}, "skipped": []}

    for sp_list in lists:
        props = sp_list.properties
        title = props.get("Title")

        skip_reason = list_skip_reason(props)
        if skip_reason:
            summary["skipped"].append(
                {
                    "title": title,
                    "item_count": props.get("ItemCount"),
                    "hidden": props.get("Hidden"),
                    "base_template": props.get("BaseTemplate"),
                    "entity_type_name": props.get("EntityTypeName"),
                    "reason": skip_reason,
                }
            )
            continue

        base_type = props.get("BaseType")
        item_count = props.get("ItemCount")
        server_last_modified = props.get("LastItemModifiedDate")
        kind = "library" if base_type == 1 else "list"

        # Resume: skip lists unchanged since the last success.
        last_extracted = _read_last_extracted(output_dir, kind, title)
        if (
            last_extracted
            and server_last_modified
            and last_extracted >= str(server_last_modified)
        ):
            logger.info(
                f"  Skipping '{title}': unchanged since last extraction ({last_extracted})."
            )
            summary["skipped"].append(
                {
                    "title": title,
                    "item_count": item_count,
                    "reason": "unchanged_since_last_extraction",
                }
            )
            continue

        if item_count and item_count > 5000:
            logger.warning(
                f"  Processing '{title}' (BaseType={base_type}, {item_count} items) — exceeds the 5000-item threshold, this may be very slow or fail."
            )
        else:
            logger.info(
                f"  Processing '{title}' (BaseType={base_type}, {item_count} items)"
            )

        keyed_list = ctx.web.lists.get_by_id(props.get("Id"))

        if base_type == 1:  # DocumentLibrary
            result = extract_document_library(keyed_list, title, output_dir, logger)
            summary["document_libraries"][title] = result
        else:
            result = extract_generic_list(keyed_list, title, output_dir, logger)
            summary["generic_lists"][title] = result

        if "error" not in result and server_last_modified:
            _write_last_extracted(output_dir, kind, title, str(server_last_modified))

    return summary


def _state_path(output_dir, kind, title):
    """Path of a list's .state.json."""
    subdir = "documents" if kind == "library" else "lists"
    return os.path.join(output_dir, subdir, title.replace("/", "_"), ".state.json")


def _read_last_extracted(output_dir, kind, title):
    """LastItemModifiedDate saved at the last success, or None (re-extract)."""
    path = _state_path(output_dir, kind, title)
    try:
        with open(path, "r", encoding="utf-8") as f:
            state = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None
    if kind == "list" and state.get("format") != GENERIC_LIST_FORMAT:
        return None
    return state.get("last_item_modified_date")


# Folder of the root site (SPX_SITE_URL) under extraction/<root>/. Not its
# last segment: a host root (/) has none, and a path root (/sites/hr) would
# share it with a subsite of the same name.
ROOT_SITE_SLUG = "_root"


def strip_page_suffix(url_suffix):
    """Strip a trailing /SitePages/*.aspx: some entries point at a page, not the site root."""
    return re.sub(r"(?i)/sitepages/[^/]+\.aspx$", "", url_suffix.rstrip("/")) or "/"


def slug_from_path(url_suffix):
    """/sites/hr/team -> team."""
    clean = strip_page_suffix(url_suffix)
    return clean.strip("/").split("/")[-1].lower()


def relative_path_in_library(file_ref, library_root):
    """
    ('/equipe/X/Docs/Sub/a.pdf', '/equipe/X/Docs') -> 'Sub/a.pdf'.
    None if a value is missing or FileRef is outside the root.
    """
    if not file_ref or not library_root:
        return None
    prefix = library_root.rstrip("/") + "/"
    if not file_ref.lower().startswith(prefix.lower()):
        return None
    parts = [p for p in file_ref[len(prefix) :].split("/") if p not in ("", ".", "..")]
    return "/".join(parts) or None


def is_root_site(url_suffix, site_base_url):
    """Whether url_suffix is the root (SPX_SITE_URL) itself. Case-insensitive."""
    return strip_page_suffix(url_suffix).lower() == root_path(site_base_url).lower()


def site_slug(url_suffix, site_base_url):
    """Folder name of a site: ROOT_SITE_SLUG for the root, else its last segment."""
    if is_root_site(url_suffix, site_base_url):
        return ROOT_SITE_SLUG
    return slug_from_path(url_suffix)


def assign_unique_slugs(sites, site_base_url):
    """
    Slug = ROOT_SITE_SLUG for the root, else the last path segment, or the
    full path when several sites share it (/finances/Actualites ->
    finances-actualites) or when it is ROOT_SITE_SLUG. Case-insensitive.
    """
    paths_by_slug = {}
    for site in sites:
        paths_by_slug.setdefault(
            site_slug(site["url_suffix"], site_base_url), set()
        ).add(site["url_suffix"].lower())
    for site in sites:
        slug = site_slug(site["url_suffix"], site_base_url)
        if is_root_site(site["url_suffix"], site_base_url):
            site["slug"] = slug
            continue
        if len(paths_by_slug[slug]) > 1 or slug == ROOT_SITE_SLUG:
            slug = (
                strip_page_suffix(site["url_suffix"])
                .strip("/")
                .replace("/", "-")
                .lower()
            )
        site["slug"] = slug
    return sites


def _known_sites(root_name, site_base_url):
    """
    All known sites of the root (from disk). Empty if none.
    """
    discovery_dir = os.environ.get("SPX_OUTPUT_DIR", "./discovery_output")
    index_path = os.path.join(discovery_dir, f"sites_index_{root_name}.json")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return build_sites_from_discovery(json.load(f), site_base_url)
    return []


def _write_last_extracted(output_dir, kind, title, last_item_modified_date):
    """Save LastItemModifiedDate so the next run can skip an unchanged list."""
    path = _state_path(output_dir, kind, title)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    state = {"last_item_modified_date": last_item_modified_date}
    if kind == "list":
        state["format"] = GENERIC_LIST_FORMAT
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f)


def _describe_principal(principal):
    """Principal as a dict. SharePoint groups are expanded; AD groups and
    distribution lists are kept by name."""
    ptype = principal.principal_type_name
    resolvable = ptype in ("User", "SharePointGroup")
    return {
        "type": ptype,
        "title": principal.title,
        "login_name": principal.login_name or principal.user_principal_name,
        "members_resolvable": resolvable,
    }


class GroupMembers:
    """
    Members of a site's SharePoint groups, read once per site: the same
    groups come back on the site, its lists and its items. A failed read is
    not kept, so the group is asked again where it next appears.
    """

    def __init__(self, ctx, logger):
        self.ctx = ctx
        self.logger = logger
        self.by_id = {}
        self.requests = 0
        self.reused = 0

    def get(self, principal):
        """[{title, login_name}] of the group, or None if it cannot be read."""
        if principal.id in self.by_id:
            self.reused += 1
            return list(self.by_id[principal.id])
        self.requests += 1
        try:
            group = self.ctx.web.site_groups.get_by_id(principal.id)
            members = group.users.get().execute_query()
        except Exception as e:
            self.logger.warning(f"    Failed to expand group '{principal.title}': {e}")
            return None
        self.by_id[principal.id] = [
            {
                "title": m.title,
                "login_name": m.login_name or m.user_principal_name,
            }
            for m in members
        ]
        return list(self.by_id[principal.id])


def _role_assignments_to_dict(securable, ctx, logger, group_members=None):
    """
    Role assignments of a site, list or item, with SharePoint group members
    expanded. None if they cannot be read.
    """
    group_members = group_members or GroupMembers(ctx, logger)
    try:
        assignments = (
            securable.role_assignments.expand(["Member", "RoleDefinitionBindings"])
            .get()
            .execute_query()
        )
    except Exception as e:
        logger.warning(f"    Failed to read role assignments: {e}")
        return None

    result = []
    for assignment in assignments:
        principal = assignment.member
        entry = _describe_principal(principal)
        bindings = list(assignment.role_definition_bindings)
        entry["roles"] = [r.name for r in bindings]
        # RoleTypeKind does not depend on the site language (5 = full control).
        entry["role_types"] = [r.properties.get("RoleTypeKind") for r in bindings]

        if principal.principal_type_name == "SharePointGroup":
            entry["members"] = group_members.get(principal)

        result.append(entry)
    return result


def _item_kind(item_props, list_props):
    """folder (document sets included), file (in a library) or item (in a list)."""
    if item_props.get("FileSystemObjectType") == 1:
        return "folder"
    return "file" if list_props.get("BaseType") == 1 else "item"


def items_with_unique_permissions(sp_list, list_props):
    """
    [(id, kind, FileRef)] of the items of sp_list with their own permissions,
    from one paged read. Raises if the server does not return
    HasUniqueRoleAssignments: an unanswered question is not a "no".
    """
    items = (
        sp_list.items.select(ITEM_PERMISSION_FIELDS)
        .get_all(page_size=ITEM_PERMISSIONS_PAGE_SIZE)
        .execute_query()
    )
    result = []
    for item in items:
        props = item.properties
        if "HasUniqueRoleAssignments" not in props:
            raise ValueError(
                f"HasUniqueRoleAssignments missing for item {props.get('Id')}"
            )
        if props["HasUniqueRoleAssignments"]:
            result.append(
                (props.get("Id"), _item_kind(props, list_props), props.get("FileRef"))
            )
    return result


def _extract_item_permissions(ctx, lists, group_members, logger):
    """
    ({list title: [item]}, [error]) for the folders, files and items with
    their own permissions, in the lists reads_permissions() keeps.
    """
    by_list, errors = {}, []
    for sp_list in lists:
        props = sp_list.properties
        if not reads_permissions(props) or not props.get("ItemCount"):
            continue
        title = props.get("Title")
        keyed_list = ctx.web.lists.get_by_id(props.get("Id"))
        try:
            flagged = items_with_unique_permissions(keyed_list, props)
        except Exception as e:
            logger.warning(f"    Failed to scan '{title}' for item permissions: {e}")
            errors.append({"list": title, "item_id": None, "error": str(e)})
            continue
        for item_id, kind, url in flagged:
            permissions = _role_assignments_to_dict(
                keyed_list.items.get_by_id(item_id), ctx, logger, group_members
            )
            if permissions is None:
                errors.append(
                    {
                        "list": title,
                        "item_id": item_id,
                        "error": "role assignments could not be read",
                    }
                )
            by_list.setdefault(title, []).append(
                {"id": item_id, "kind": kind, "url": url, "permissions": permissions}
            )
    return by_list, errors


def extract_permissions(
    site_url, slug, username, password, ca_bundle_path, base_output_dir, logger
):
    """
    Write permissions.json: site role assignments, those of lists with
    broken inheritance, and those of folders, files and items with their own
    permissions (items_with_broken_inheritance). What could not be read is
    listed in items_scan_errors.
    """
    output_dir = os.path.join(base_output_dir, slug)
    os.makedirs(output_dir, exist_ok=True)
    logger.info(f"--- Extracting permissions: {slug} ---")

    try:
        ctx = build_context(site_url, username, password, ca_bundle_path)
        web = (
            ctx.web.get().select(["Title", "HasUniqueRoleAssignments"]).execute_query()
        )
    except Exception as e:
        logger.error(f"    Failed to connect for permissions: {e}")
        return {"error": str(e)}

    group_members = GroupMembers(ctx, logger)
    site_permissions = _role_assignments_to_dict(ctx.web, ctx, logger, group_members)
    result = {
        "format": PERMISSIONS_FORMAT,
        "site_has_unique_role_assignments": web.has_unique_role_assignments,
        "site_permissions": site_permissions,
        "lists_with_broken_inheritance": {},
        "items_with_broken_inheritance": {},
        "items_scan_errors": [],
    }

    try:
        lists = ctx.web.lists.select(PERMISSION_LIST_FIELDS).get().execute_query()
    except Exception as e:
        logger.warning(f"    Failed to list lists/libraries for permissions: {e}")
        result["items_scan_errors"].append(
            {"list": None, "item_id": None, "error": str(e)}
        )
        lists = []

    for sp_list in lists:
        if not reads_permissions(sp_list.properties):
            continue
        if not sp_list.properties.get("HasUniqueRoleAssignments"):
            continue
        title = sp_list.properties.get("Id")
        try:
            keyed_list = (
                ctx.web.lists.get_by_id(title).get().select(["Title"]).execute_query()
            )
        except Exception as e:
            logger.warning(
                f"    Failed to re-fetch a list with broken inheritance: {e}"
            )
            continue
        list_permissions = _role_assignments_to_dict(
            keyed_list, ctx, logger, group_members
        )
        result["lists_with_broken_inheritance"][keyed_list.title] = list_permissions

    items, errors = _extract_item_permissions(ctx, lists, group_members, logger)
    result["items_with_broken_inheritance"] = items
    result["items_scan_errors"] += errors
    logger.info(
        f"    Items with their own permissions: {sum(len(v) for v in items.values())}, "
        f"scan errors: {len(result['items_scan_errors'])}, group member reads: "
        f"{group_members.requests} ({group_members.reused} reused)"
    )

    with open(os.path.join(output_dir, "permissions.json"), "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2, default=str)

    return {
        "lists_with_broken_inheritance_count": len(
            result["lists_with_broken_inheritance"]
        ),
        "items_with_broken_inheritance_count": sum(len(v) for v in items.values()),
        "items_scan_errors_count": len(result["items_scan_errors"]),
        "group_members_requests": group_members.requests,
        "group_members_reused": group_members.reused,
    }


def permissions_up_to_date(permissions_path):
    """Whether permissions.json exists in the current format (else it is redone)."""
    try:
        with open(permissions_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return False
    return isinstance(data, dict) and data.get("format") == PERMISSIONS_FORMAT


def run_permissions_pass(
    sites, username, password, ca_bundle_path, site_base_url, base_output_dir, logger
):
    """
    Run extract_permissions() on each site without a permissions.json in the
    current format: files written by an older version are redone.
    """
    logger.info(f"--- Starting permissions pass on {len(sites)} site(s) ---")
    permissions_summaries = {}
    skipped_already_done = 0
    for site in sites:
        permissions_path = os.path.join(
            base_output_dir, site["slug"], "permissions.json"
        )
        if permissions_up_to_date(permissions_path):
            skipped_already_done += 1
            continue
        site_url = absolute_url(site_base_url, site["url_suffix"])
        result = extract_permissions(
            site_url,
            site["slug"],
            username,
            password,
            ca_bundle_path,
            base_output_dir,
            logger,
        )
        permissions_summaries[site["slug"]] = result

    if skipped_already_done:
        logger.info(
            f"Skipped {skipped_already_done} site(s) with an up-to-date permissions.json."
        )

    permissions_summary_path = os.path.join(
        base_output_dir, "permissions_run_summary.json"
    )
    merged = {}
    if os.path.exists(permissions_summary_path):
        try:
            with open(permissions_summary_path, "r", encoding="utf-8") as f:
                merged = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            logger.error(
                f"Could not read {permissions_summary_path} ({e}): not overwriting it."
            )
            return
    merged.update(permissions_summaries)
    with open(permissions_summary_path, "w", encoding="utf-8") as f:
        json.dump(merged, f, ensure_ascii=False, indent=2, default=str)
    logger.info(f"Permissions pass done. Summary written to {permissions_summary_path}")


def build_sites_from_discovery(discovered_sites, site_base_url):
    """
    discover_sites.py output -> [{"slug", "url_suffix", "label"}].
    Skips failed entries. The root itself is kept, in ROOT_SITE_SLUG.
    """
    sites = []
    for entry in discovered_sites:
        if entry.get("status") != "ok":
            continue
        path = entry.get("server_relative_url")
        if not path:
            continue
        path = strip_page_suffix(path)
        sites.append(
            {
                "slug": slug_from_path(path),
                "url_suffix": path,
                "label": entry.get("title") or path,
            }
        )
    return assign_unique_slugs(sites, site_base_url)


def main():
    enable_persistent_session()

    site_base_url, username, password, ca_bundle_path = get_config()

    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s"
    )
    logger = logging.getLogger("extract_metadata")

    root_name = get_root_name()
    base_output_dir = os.path.join(
        os.environ.get("SPX_OUTPUT_DIR", "./discovery_output"),
        "extraction",
        root_name,
    )
    os.makedirs(base_output_dir, exist_ok=True)

    args = sys.argv[1:]

    if not args:
        print(__doc__)
        sys.exit(1)

    if args[0] == "--permissions-all":
        summary_path = os.path.join(base_output_dir, "run_summary.json")
        known_sites = _known_sites(root_name, site_base_url)
        if not known_sites or not os.path.exists(summary_path):
            print(
                f"[ERROR] Missing sites_index_{root_name}.json or {summary_path}. "
                "Run the discovery and --all-from-discovery first."
            )
            sys.exit(1)
        with open(summary_path, "r", encoding="utf-8") as f:
            all_summaries = json.load(f)

        successful_sites = [
            s
            for s in known_sites
            if s["slug"] in all_summaries and "error" not in all_summaries[s["slug"]]
        ]
        run_permissions_pass(
            successful_sites,
            username,
            password,
            ca_bundle_path,
            site_base_url,
            base_output_dir,
            logger,
        )
        return

    if args[0] == "--all-from-discovery":
        if len(args) < 2:
            print(
                "Usage: python extract_metadata.py --all-from-discovery <discovery_output.json>"
            )
            sys.exit(1)
        discovery_path = args[1]
        if not os.path.exists(discovery_path):
            print(f"[ERROR] '{discovery_path}' not found.")
            sys.exit(1)
        with open(discovery_path, "r", encoding="utf-8") as f:
            discovered = json.load(f)

        sites_to_process = build_sites_from_discovery(discovered, site_base_url)
        print(
            f"[INFO] {len(sites_to_process)} site(s) to process from '{discovery_path}'."
        )

        all_summaries = {}
        for site in sites_to_process:
            site_url = absolute_url(site_base_url, site["url_suffix"])
            logger.info(f"=== {site['label']} ===")
            summary = extract_site(
                site_url,
                site["slug"],
                username,
                password,
                ca_bundle_path,
                base_output_dir,
                logger,
            )
            all_summaries[site["slug"]] = summary

        summary_path = os.path.join(base_output_dir, "run_summary.json")
        with open(summary_path, "w", encoding="utf-8") as f:
            json.dump(all_summaries, f, ensure_ascii=False, indent=2, default=str)
        logger.info(f"Done. Summary written to {summary_path}")

        successful_sites = [
            s
            for s in sites_to_process
            if "error" not in all_summaries.get(s["slug"], {})
        ]
        run_permissions_pass(
            successful_sites,
            username,
            password,
            ca_bundle_path,
            site_base_url,
            base_output_dir,
            logger,
        )

        consolidate_script = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "consolidate_errors.py"
        )
        extraction_root = os.path.dirname(base_output_dir)
        logger.info("--- Consolidating errors across all run_summary.json files ---")
        result = subprocess.run([sys.executable, consolidate_script, extraction_root])
        if result.returncode != 0:
            logger.warning(
                "consolidate_errors.py exited with an error; error consolidation "
                "was not completed, but the --all-from-discovery run itself is otherwise done."
            )
        return

    url_suffix = args[0]
    if not url_suffix.startswith("/"):
        url_suffix = "/" + url_suffix
    url_suffix = strip_page_suffix(url_suffix)
    site = {
        "slug": slug_from_path(url_suffix),
        "url_suffix": url_suffix,
        "label": url_suffix,
    }
    # Keep only the requested site, with its root-wide unique slug.
    site = assign_unique_slugs(
        _known_sites(root_name, site_base_url) + [site], site_base_url
    )[-1]
    site_url = absolute_url(site_base_url, site["url_suffix"])
    logger.info(f"=== {site['label']} ===")
    summary = extract_site(
        site_url,
        site["slug"],
        username,
        password,
        ca_bundle_path,
        base_output_dir,
        logger,
    )

    if "error" in summary:
        logger.warning(
            f"Skipping permissions for '{site['slug']}': metadata extraction failed."
        )
    else:
        perm_result = extract_permissions(
            site_url,
            site["slug"],
            username,
            password,
            ca_bundle_path,
            base_output_dir,
            logger,
        )
        logger.info(f"Permissions result for '{site['slug']}': {perm_result}")

    summary_path = os.path.join(base_output_dir, "run_summary.json")
    summaries_to_write = {}
    if os.path.exists(summary_path):
        try:
            with open(summary_path, "r", encoding="utf-8") as f:
                summaries_to_write = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            logger.error(
                f"Could not read existing {summary_path} ({e}): not overwriting it. "
                f"This run's summary: {json.dumps(summary, default=str)[:500]}"
            )
            return
    summaries_to_write[site["slug"]] = summary

    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summaries_to_write, f, ensure_ascii=False, indent=2, default=str)

    logger.info(f"Done. Summary written to {summary_path}")


if __name__ == "__main__":
    main()
