#!/usr/bin/env python3

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

import sys

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "scripts"))

from kamus_source import (  # noqa: E402
    canonicalize_source_workbook,
    resolve_kamus_source_root,
    resolve_source_workbook,
)


class KamusSourceTests(unittest.TestCase):
    def test_default_source_root_is_repo_ho5(self) -> None:
        context = resolve_kamus_source_root()
        self.assertTrue(context.source_root.is_dir())
        self.assertIn("kamus-ho-config-20260729", str(context.source_root))
        self.assertIn("(HO) 5", str(context.source_root))
        self.assertTrue(context.inventory_config.is_file())

    def test_rejects_downloads_source_root(self) -> None:
        with self.assertRaises(ValueError):
            resolve_kamus_source_root(
                explicit_root=Path("/Users/example/Downloads/KAMUS KPI PELINDO GROUP 1 (HO) 4"),
            )

    def test_canonicalize_ho4_alias_to_inventory_workbook(self) -> None:
        inventory = json.loads((REPO / "configs/kamus_kpi_ho_visible_20260729.json").read_text())
        alias = (
            "Group Pengendalian Proyek/DIREKTORAT TEKNIK - Ibu Ika Oktania - "
            "Pengendalian Proyek (Selesai konfirmasi KPI) (3).xlsx"
        )
        canonical = canonicalize_source_workbook(alias, inventory)
        self.assertEqual(
            canonical,
            "Group Pengendalian Proyek/DIREKTORAT TEKNIK - Ibu Ika Oktania - "
            "Group Pengendalian Proyek (Selesai konfirmasi KPI).xlsx",
        )

    def test_resolve_workbook_from_inventory(self) -> None:
        context = resolve_kamus_source_root()
        inventory = json.loads(context.inventory_config.read_text(encoding="utf-8"))
        path = resolve_source_workbook(
            context.source_root,
            "Group Sekretariat Perusahaan/DIREKTORAT UTAMA - Group Sekretariat Perusahaan.xlsx",
            inventory,
        )
        self.assertTrue(path.is_file())

    def test_allow_external_only_when_explicit(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            external_root = Path(tmp) / "external-kamus"
            external_root.mkdir()
            context = resolve_kamus_source_root(explicit_root=external_root, allow_external=True)
            self.assertEqual(context.source_root, external_root.resolve())


if __name__ == "__main__":
    unittest.main()
