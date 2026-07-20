"""Tests für die reinen Helfer in teams_export.py (keine Graph-/Netzwerk-Aufrufe)."""

import teams_export as te


def test_parse_indices():
    assert te.parse_indices("1, 3 5", 5) == [1, 3, 5]
    assert te.parse_indices("2,2,2", 5) == [2]          # Duplikate raus
    assert te.parse_indices("0, 6, 99", 5) == []        # außerhalb des Bereichs
    assert te.parse_indices("a, b", 5) == []
    assert te.parse_indices("", 5) == []


def test_safe_filenames():
    assert te.safe('a/b\\c:d*e?"f<g>h|i') == "a_b_c_d_e_f_g_h_i"
    assert te.safe("  viel   Leerraum  ") == "viel Leerraum"
    assert te.safe("x" * 200, maxlen=10) == "x" * 10
    assert te.safe("") == "unbenannt"
    assert te.safe(None) == "unbenannt"
    assert te.safe("...") == "unbenannt"                # nur Punkte -> leer


def test_short_id_is_stable_hex():
    assert te.short_id("abc") == te.short_id("abc")
    assert te.short_id("abc") != te.short_id("abd")
    assert len(te.short_id("abc")) == 8
    int(te.short_id("abc"), 16)  # hex


def test_parse_ts_handles_graph_timestamps():
    dt = te.parse_ts("2025-06-01T09:30:00Z")
    assert dt is not None and dt.tzinfo is not None
    # Graph liefert teils 7-stellige Sekundenbruchteile
    assert te.parse_ts("2025-06-01T09:30:00.1234567Z") is not None
    assert te.parse_ts("unsinn") is None
    assert te.parse_ts("") is None
    assert te.parse_ts(None) is None


def test_newest_iso_picks_latest_and_ignores_garbage():
    strings = ["2025-06-01T09:30:00Z", "kaputt", "2025-06-02T08:00:00Z", ""]
    assert te.newest_iso(strings) == "2025-06-02T08:00:00Z"
    assert te.newest_iso(["kaputt", ""]) is None
    assert te.newest_iso([]) is None


def test_strip_tags():
    assert te.strip_tags("<p>Hallo <b>Welt</b></p>") == "Hallo Welt"
    assert te.strip_tags(None) == ""


def test_clean_html_removes_scripts_and_event_handlers():
    s = '<div onclick="evil()" onmouseover=\'evil()\'>ok</div><script>alert(1)</script>'
    out = te.clean_html(s)
    assert "script" not in out
    assert "onclick" not in out
    assert "onmouseover" not in out
    assert ">ok</div>" in out


def test_default_categories():
    options = [("1on1", "a"), ("group", "b"), ("meeting", "c"), ("channels", "d")]
    assert te.default_categories(options) == {"1on1", "group", "meeting"}
