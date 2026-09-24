"""Selection and DS-field filter tests against the RF-RK catalogue extract."""

from __future__ import annotations

import unittest
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from mgm_selector.catalogue import MISSING, Catalogue  # noqa: E402
from mgm_selector.pdf_report import generate_pdf  # noqa: E402
from mgm_selector.selection import SelectionEngine  # noqa: E402

DS_KEYS = {
    "rated_motor_power",
    "rated_motor_speed",
    "overall_gear_ratio",
    "output_speed",
    "output_torque",
    "service_factor",
    "output_shaft_type",
    "output_shaft_dimensions",
    "permitted_output_overhung_load",
    "direction_of_rotation",
    "mounting_position",
    "flange_diameter",
    "lubricant_type",
    "lubricant_quantity",
    "position_of_terminal_box",
    "cable_entry_position",
    "finished_paint_colour",
    "motor_type",
    "number_of_poles",
    "duration_factor",
    "efficiency_class",
    "efficiency",
    "ce_marking",
    "motor_voltage",
    "frequency",
    "rated_current",
    "power_factor",
    "insulation_class",
    "protection_class",
    "design_requirement",
    "weight",
    "additional_features",
}


class SelectionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.engine = SelectionEngine(Catalogue())

    def _val(self, report, key):
        for f in report["fields"]:
            if f["key"] == key:
                return f["value"]
        self.fail(key)

    def test_ds_field_set(self):
        keys = self.engine.ds_field_keys()
        self.assertEqual(set(keys), DS_KEYS)

    def test_valid_power_ratio_rf(self):
        r = self.engine.lookup("RF", 1.5, 36.53)
        self.assertTrue(r["ok"], r)
        self.assertEqual(self._val(r["report"], "rated_motor_power"), "1.5")
        self.assertEqual(self._val(r["report"], "overall_gear_ratio"), "36.53")
        self.assertEqual(self._val(r["report"], "output_speed"), "39.4")
        self.assertEqual(self._val(r["report"], "output_torque"), "349")
        self.assertEqual(self._val(r["report"], "service_factor"), "1.1")

    def test_ratio_change_updates_dependents(self):
        a = self.engine.lookup("RF", 1.5, 20.0)
        b = self.engine.lookup("RF", 1.5, 36.53)
        self.assertTrue(a["ok"] and b["ok"])
        self.assertEqual(self._val(a["report"], "output_speed"), "72")
        self.assertEqual(self._val(a["report"], "output_torque"), "191")
        self.assertNotEqual(
            self._val(a["report"], "output_speed"),
            self._val(b["report"], "output_speed"),
        )
        self.assertNotEqual(
            self._val(a["report"], "output_torque"),
            self._val(b["report"], "output_torque"),
        )
        self.assertNotEqual(
            self._val(a["report"], "service_factor"),
            self._val(b["report"], "service_factor"),
        )

    def test_power_filters_ratios(self):
        r15 = set(self.engine.ratios("RF", 1.5))
        r22 = set(self.engine.ratios("RF", 2.2))
        self.assertTrue(r15)
        self.assertNotEqual(r15, r22)
        self.assertNotIn(36.53, r22)

    def test_invalid_combination_rejected(self):
        r = self.engine.lookup("RF", 1.5, 999.99)
        self.assertFalse(r["ok"])
        self.assertIn("No exact catalogue match", r["error"])

    def test_rf_and_rk_not_mixed(self):
        rf = self.engine.lookup("RF", 1.5, 25.00)
        # 25.00 exists on RK at 1.5 kW; RF 1.5 has 24.96 not 25.00
        self.assertFalse(rf["ok"])
        rk = self.engine.lookup("RK", 1.5, 25.00)
        self.assertTrue(rk["ok"], rk)
        self.assertTrue(rk["report"]["internal"]["model"].startswith("RK"))
        self.assertTrue(all(m["series"] == "RK" for m in rk["matches"]))

    def test_pdf_contains_only_ds_fields(self):
        r = self.engine.lookup("RF", 1.5, 20.0, model="RF 042 - IEC 90 B5")
        self.assertTrue(r["ok"])
        keys = [f["key"] for f in r["report"]["fields"]]
        self.assertEqual(keys, self.engine.ds_field_keys())
        extra_catalogue_keys = {"Fr1", "P1max", "Mn2", "Gd2"}
        self.assertTrue(extra_catalogue_keys.isdisjoint(set(keys)))
        out = ROOT / "scratch" / "test_ds0210046.pdf"
        generate_pdf(r["report"], out)
        self.assertTrue(out.exists() and out.stat().st_size > 1000)

    def test_unavailable_motor_electrical_not_invented(self):
        r = self.engine.lookup("RK", 1.5, 25.00)
        self.assertEqual(self._val(r["report"], "efficiency"), MISSING)
        self.assertEqual(self._val(r["report"], "rated_current"), MISSING)
        self.assertEqual(self._val(r["report"], "power_factor"), MISSING)
        self.assertEqual(self._val(r["report"], "finished_paint_colour"), MISSING)

    def test_lubricant_from_catalogue_not_ds_example(self):
        r = self.engine.lookup("RF", 1.5, 20.0, model="RF 042 - IEC 90 B5")
        self.assertEqual(self._val(r["report"], "lubricant_type"), "ISO VG 320")

    def test_ds_example_ratio_not_in_rf_rk(self):
        """DS-021-0046 sample is IR FPM 63 i=35.74 — not an RF/RK table row."""
        r = self.engine.lookup("RF", 1.5, 35.74)
        self.assertFalse(r["ok"])
        r2 = self.engine.lookup("RK", 1.5, 35.74)
        self.assertFalse(r2["ok"])


if __name__ == "__main__":
    unittest.main()
