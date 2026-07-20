"""Tests für corpus.py – Parsen der Exporte und Chunking (nur Standardbibliothek)."""

import textwrap

import corpus


# --------------------------------------------------------------------------
# HTML / Text-Aufbereitung
# --------------------------------------------------------------------------
def test_strip_html_removes_tags_scripts_and_entities():
    s = "<p>Hallo <b>Welt</b></p><script>alert(1)</script><style>p{}</style>&amp; mehr"
    out = corpus.strip_html(s)
    assert "alert" not in out
    assert "p{}" not in out
    assert "<" not in out
    assert "Hallo" in out and "Welt" in out
    assert "& mehr" in out


def test_collapse_whitespace_and_cap():
    assert corpus.collapse("  a \n\t b   c ") == "a b c"
    assert corpus.collapse("x" * 100, cap=10) == "x" * 10
    assert corpus.collapse(None) == ""


def test_strip_quoted_cuts_outlook_history():
    text = textwrap.dedent("""\
        Danke, passt für mich!

        ________________________________
        Von: Alice Example <alice@example.com>
        Gesendet: Montag, 7. Juli 2025 10:00
        Betreff: AW: Termin
        Alter zitierter Text.
        """)
    out = corpus.strip_quoted(text)
    assert "Danke, passt für mich!" in out
    assert "Alter zitierter Text" not in out
    assert "Gesendet" not in out


def test_strip_quoted_cuts_on_wrote_marker_and_quote_lines():
    text = "Neue Antwort.\n\nAm 07.07.2025 um 10:00 schrieb Bob:\n> alte Zeile\n> noch eine\n"
    out = corpus.strip_quoted(text)
    assert "Neue Antwort." in out
    assert "alte Zeile" not in out


def test_strip_quoted_cuts_signature():
    text = "Kurze Antwort.\n-- \nAlice Example\nFirma GmbH\n"
    out = corpus.strip_quoted(text)
    assert "Kurze Antwort." in out
    assert "Firma GmbH" not in out


def test_parse_local():
    assert corpus.parse_local("2025-07-07 10:00") is not None
    assert corpus.parse_local("kein datum") is None
    assert corpus.parse_local(None) is None


# --------------------------------------------------------------------------
# Chunking
# --------------------------------------------------------------------------
def test_split_short_text_is_single_chunk():
    assert corpus._split("Hallo Welt", 100, 20) == ["Hallo Welt"]
    assert corpus._split("   ", 100, 20) == []
    assert corpus._split("", 100, 20) == []


def test_split_long_text_overlaps_and_covers():
    words = " ".join(f"wort{i}" for i in range(200))
    chunks = corpus._split(words, size=120, overlap=30)
    assert len(chunks) > 1
    # Jedes Stück ist Substring des Originals; Anfang und Ende sind abgedeckt
    for c in chunks:
        assert c in words
    assert words.startswith(chunks[0])
    assert words.endswith(chunks[-1])
    # Überlappung: der Anfang jedes Stücks liegt noch im Vorgänger
    for a, b in zip(chunks, chunks[1:], strict=False):
        assert b[:15] in a
    # Kein Stück (deutlich) über der Zielgröße
    assert all(len(c) <= 120 for c in chunks)


def test_chunk_records_assigns_chunk_ids():
    rec = {"uid": "outlook:a.eml:0", "title": "Betreff", "text": "kurzer text"}
    chunks = corpus.chunk_records([rec], size=1500, overlap=200)
    assert len(chunks) == 1
    assert chunks[0]["cid"] == "outlook:a.eml:0#0"

    long_rec = {"uid": "u", "title": "t", "text": "x" * 4000}
    chunks = corpus.chunk_records([long_rec], size=1500, overlap=200)
    assert len(chunks) > 1
    assert [c["cid"] for c in chunks] == [f"u#{j}" for j in range(len(chunks))]


def test_embed_text_and_hash_are_deterministic():
    c = {"title": "Betreff", "text": "Inhalt"}
    assert corpus.embed_text(c) == "Betreff\nInhalt"
    assert corpus.chunk_hash(c) == corpus.chunk_hash(dict(c))
    assert corpus.chunk_hash(c) != corpus.chunk_hash({"title": "Betreff", "text": "anders"})


# --------------------------------------------------------------------------
# Teams-HTML
# --------------------------------------------------------------------------
TEAMS_HTML = """<html><body>
<h1>Projekt Alpha</h1>
<div class="msg">
  <span class="name">Alice Example</span>
  <span class="time">2025-06-01 09:30</span>
  <div class="body"><p>Hallo <b>Bob</b>,</p><div>wie besprochen.</div></div>
</div>
<div class="msg">
  <span class="name">Bob</span>
  <span class="time">2025-06-01 09:35</span>
  <div class="body">Danke!</div>
</div>
</body></html>"""


def test_conv_parser_extracts_title_and_messages():
    pr = corpus.ConvParser()
    pr.feed(TEAMS_HTML)
    pr.finish()
    assert pr.title == "Projekt Alpha"
    assert len(pr.msgs) == 2
    assert pr.msgs[0]["n"] == "Alice Example"
    assert pr.msgs[0]["t"] == "2025-06-01 09:30"
    assert "Hallo Bob" in " ".join(pr.msgs[0]["x"].split())
    assert pr.msgs[1] == {"n": "Bob", "t": "2025-06-01 09:35", "x": "Danke!"}


def test_load_teams_builds_records(tmp_path):
    d = tmp_path / "1on1"
    d.mkdir()
    (d / "alice__abc123.html").write_text(TEAMS_HTML, encoding="utf-8")
    (tmp_path / "index.html").write_text("<html></html>", encoding="utf-8")  # wird ignoriert

    recs = corpus.load_teams(str(tmp_path))
    assert len(recs) == 2
    r = recs[0]
    assert r["uid"] == "teams:1on1/alice__abc123.html:0"
    assert r["src"] == "teams"
    assert r["ctx"] == "1:1-Chat"
    assert r["who"] == "Alice Example"
    assert "alice example" in r["ppl"]
    assert r["ts"] is not None


# --------------------------------------------------------------------------
# Outlook (.eml)
# --------------------------------------------------------------------------
EML = b"""\
From: Alice Example <alice@example.com>
To: Bob Builder <bob@example.com>
Subject: Testmail
Date: Mon, 07 Jul 2025 10:00:00 +0000
Content-Type: text/plain; charset=utf-8

Hallo Bob,

hier die neue Nachricht.

________________________________
Von: Bob Builder <bob@example.com>
Gesendet: Sonntag, 6. Juli 2025 09:00
Alter zitierter Verlauf.
"""


def test_load_outlook_parses_eml(tmp_path):
    d = tmp_path / "inbox"
    d.mkdir()
    (d / "mail.eml").write_bytes(EML)

    recs = corpus.load_outlook(str(tmp_path))
    assert len(recs) == 1
    r = recs[0]
    assert r["uid"] == "outlook:inbox/mail.eml:0"
    assert r["who"] == "Alice Example"
    assert r["title"] == "Testmail"
    assert r["ctx"] == "inbox"
    assert "bob@example.com" in r["ppl"]
    assert "hier die neue Nachricht" in r["text"]
    assert "Alter zitierter Verlauf" not in r["text"]
    assert r["ts"] is not None


# --------------------------------------------------------------------------
# Kalender (.ics) und Kontakte (.vcf)
# --------------------------------------------------------------------------
ICS = "\r\n".join([
    "BEGIN:VCALENDAR",
    "BEGIN:VEVENT",
    "SUMMARY:Planung\\, Quartal",
    "LOCATION:Raum 42",
    "DESCRIPTION:Agenda folgt",
    "DTSTART:20250601T120000Z",
    'ORGANIZER;CN="Alice Example":mailto:alice@example.com',
    'ATTENDEE;CN="Bob Builder":mailto:bob@example.com',
    "END:VEVENT",
    "END:VCALENDAR",
])


def test_load_calendar_parses_ics(tmp_path):
    d = tmp_path / "kalender" / "Arbeit"
    d.mkdir(parents=True)
    (d / "termin.ics").write_text(ICS, encoding="utf-8")

    recs = corpus.load_calendar(str(tmp_path))
    assert len(recs) == 1
    r = recs[0]
    assert r["title"] == "Planung, Quartal"
    assert r["ctx"] == "Kalender: Arbeit"
    assert r["who"] == "Alice Example"
    assert "bob@example.com" in r["ppl"]
    assert r["text"].startswith("Ort: Raum 42.")
    assert r["ts"] is not None


VCF = "\r\n".join([
    "BEGIN:VCARD",
    "FN:Alice Example",
    "N:Example;Alice;;;",
    "ORG:Firma GmbH;Entwicklung",
    "TITLE:Engineer",
    "EMAIL:alice@example.com",
    "TEL:+49 123 456",
    "NOTE:Erste Zeile",
    " weiter gefaltet",
    "END:VCARD",
])


def test_load_contacts_parses_vcf(tmp_path):
    d = tmp_path / "kontakte" / "Team"
    d.mkdir(parents=True)
    (d / "alice.vcf").write_text(VCF, encoding="utf-8")

    recs = corpus.load_contacts(str(tmp_path))
    assert len(recs) == 1
    r = recs[0]
    assert r["title"] == "Alice Example"
    assert r["ctx"] == "Kontakte: Team"
    assert "Firma GmbH · Entwicklung" in r["text"]
    assert "Erste Zeileweiter gefaltet" in r["text"]  # RFC-Zeilenfaltung aufgelöst
    assert "alice@example.com" in r["ppl"]


def test_ics_when_variants():
    ts, disp = corpus._ics_when("20250601", dateonly=True)
    assert disp == "2025-06-01" and ts is not None
    ts, disp = corpus._ics_when("20250601T120000", dateonly=False)
    assert disp == "2025-06-01 12:00" and ts is not None
    ts, disp = corpus._ics_when("", dateonly=False)
    assert ts is None and disp == ""
    ts, disp = corpus._ics_when("unsinn", dateonly=False)
    assert ts is None and disp == "unsinn"
