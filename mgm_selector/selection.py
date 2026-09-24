"""Exact catalogue lookup. No nearest-ratio substitution."""

from __future__ import annotations

from typing import Any

from .catalogue import NOT_FOUND_MSG, Catalogue, build_report


class SelectionEngine:
    def __init__(self, catalogue: Catalogue | None = None) -> None:
        self.cat = catalogue or Catalogue()

    def powers(self, series: str) -> list[float]:
        return self.cat.powers_for_series(series)

    def ratios(self, series: str, power: float) -> list[float]:
        return self.cat.ratios_for(series, power)

    def lookup(
        self,
        series: str,
        power: float,
        ratio: float,
        *,
        model: str | None = None,
        mounting: str = "M1",
        shaft_code: str = "AC",
        motor_kind: str = "SMX",
        terminal_box: str = "KP1",
        cable_entry: str = "A",
    ) -> dict[str, Any]:
        hits = self.cat.matches(series, power, ratio)
        if model:
            hits = [h for h in hits if h["model"] == model]
        if not hits:
            return {"ok": False, "error": NOT_FOUND_MSG, "matches": []}
        if len(hits) > 1 and not model:
            return {
                "ok": False,
                "error": "Multiple exact catalogue records exist. Select the configuration.",
                "matches": hits,
            }
        report = build_report(
            self.cat,
            hits[0],
            mounting=mounting,
            shaft_code=shaft_code,
            motor_kind=motor_kind,
            terminal_box=terminal_box,
            cable_entry=cable_entry,
        )
        return {"ok": True, "report": report, "matches": hits}

    def ds_field_keys(self) -> list[str]:
        return [f["key"] for f in self.cat.ds_fields["fields"]]
