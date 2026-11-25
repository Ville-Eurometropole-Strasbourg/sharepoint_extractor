import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src" / "extraction"))

import aggregate_stats as stats  # noqa: E402


def _write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data), encoding="utf-8")


class ScanExtractionDirTest(unittest.TestCase):
    def test_folders_and_document_sets_are_not_counted_as_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            lib_dir = Path(tmp) / "intranet" / "team" / "documents" / "Docs"
            _write_json(
                lib_dir / "Lot 1.meta.json",
                {"FileSystemObjectType": 1, "OData__UIVersionString": "1.0"},
            )
            (lib_dir / "Lot 1").mkdir()
            _write_json(
                lib_dir / "Lot 1" / "offre.pdf.meta.json",
                {
                    "FileSystemObjectType": 0,
                    "FileSizeDisplay": "100",
                    "OData__UIVersionString": "1.0",
                },
            )

            scanned = stats.scan_extraction_dir(tmp)

        team = scanned["intranet"]["sites"]["team"]
        self.assertEqual(team["total_files"], 1)
        self.assertEqual(dict(team["by_extension_group"]), {"pdf": 1})
        self.assertEqual(team["total_size_bytes"], 100)
        self.assertEqual(team["files_with_unknown_size"], 0)


if __name__ == "__main__":
    unittest.main()
