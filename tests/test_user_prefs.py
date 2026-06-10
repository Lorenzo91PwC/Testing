"""Tests for the tiny persistent UI prefs store."""
from __future__ import annotations

from pathlib import Path

import excel_pipeline.user_prefs as up


def test_save_and_load_roundtrip(tmp_path: Path, monkeypatch) -> None:
    """Values written via save_pref must be readable via load_pref."""
    monkeypatch.setattr(up, "_PREFS_PATH", tmp_path / "prefs.json")

    assert up.load_pref("sunrise_output_folder") is None
    up.save_pref("sunrise_output_folder", r"C:\Users\loren\Sunrise_outputs")
    assert up.load_pref("sunrise_output_folder") == r"C:\Users\loren\Sunrise_outputs"


def test_load_pref_default(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr(up, "_PREFS_PATH", tmp_path / "prefs.json")
    assert up.load_pref("missing_key", default="fallback") == "fallback"


def test_save_pref_does_not_drop_other_keys(
    tmp_path: Path, monkeypatch,
) -> None:
    """Writing one key must preserve previously stored keys."""
    monkeypatch.setattr(up, "_PREFS_PATH", tmp_path / "prefs.json")
    up.save_pref("sunrise_output_folder", "/sunrise/dir")
    up.save_pref("astra_output_folder", "/astra/dir")
    assert up.load_pref("sunrise_output_folder") == "/sunrise/dir"
    assert up.load_pref("astra_output_folder") == "/astra/dir"


def test_load_pref_on_corrupted_file(
    tmp_path: Path, monkeypatch,
) -> None:
    """A malformed JSON file must not crash the reader."""
    target = tmp_path / "prefs.json"
    target.write_text("not json {", encoding="utf-8")
    monkeypatch.setattr(up, "_PREFS_PATH", target)
    assert up.load_pref("sunrise_output_folder", default="fallback") == "fallback"
