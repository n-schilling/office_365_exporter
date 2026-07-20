"""Tests für den statischen Suchseiten-Generator combined_search.py.

Geprüft werden die HTML-/Text-Helfer und Parser, der Index-Aufbau aus
synthetischen Export-Bäumen in tmp_path sowie die erzeugte Suchseite samt
eingebettetem JSON-Index (Template-Gerüst, Records, Escaping, relative Links).
Keine Netzwerkzugriffe.
"""

import json
import re
import sys
from email.header import Header

import pytest

import combined_search

# --------------------------------------------------------------------------
# Gemeinsame Helfer und Fixtures
# --------------------------------------------------------------------------
IDX_RE = re.compile(r'<script type="application/json" id="idx">(.*?)</script>', re.S)


def read_page(path):
    """Liest die erzeugte Suchseite und extrahiert den eingebetteten JSON-Index.

    Wäre das Escaping des Payloads kaputt (rohes "</script>" im Index), endete
    der Script-Block zu früh und json.loads schlüge fehl – die Extraktion
    prüft das Escaping also implizit mit.
    """
    html = path.read_text(encoding="utf-8")
    m = IDX_RE.search(html)
    assert m is not None, "Index-<script>-Block nicht gefunden"
    return html, json.loads(m.group(1))


TEAMS_HTML = """<html><body>
<h1>Projekt Alpha</h1>
<p class="sub">2 Teilnehmer</p>
<div class="msg">
  <span class="name">Alice Example</span>
  <span class="time">2025-06-01 09:30</span>
  <div class="body"><p>Hallo <b>Bob</b>,</p><div>Gr&uuml;&szlig;e &amp; bis morgen 🎉</div></div>
</div>
<div class="msg">
  <span class="name">Bob</span>
  <span class="time">2025-06-01 09:35</span>
  <div class="body">Danke!</div>
</div>
</body></html>"""

EVIL_SNIPPET = "</script><script>alert(1)</script>"


def make_eml(body="Hallo Bob, hier der Inhalt.", subject="Testmail",
             frm="Alice Example <alice@example.com>",
             to="Bob Builder <bob@example.com>", cc=None,
             date="Mon, 07 Jul 2025 10:00:00 +0000",
             ctype="text/plain; charset=utf-8"):
    """Baut eine minimale .eml als Bytes."""
    lines = [f"From: {frm}", f"To: {to}"]
    if cc:
        lines.append(f"Cc: {cc}")
    lines += [f"Subject: {subject}", f"Date: {date}", f"Content-Type: {ctype}", "", body]
    return "\r\n".join(lines).encode("utf-8")


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


# --------------------------------------------------------------------------
# combined_search: Helfer
# --------------------------------------------------------------------------
def test_combined_parse_local():
    assert combined_search.parse_local("2025-07-07 10:00") is not None
    assert combined_search.parse_local("kein datum") is None
    assert combined_search.parse_local(None) is None


def test_combined_link_kodiert_segmente(tmp_path):
    p = tmp_path / "export" / "Ordner mit Leerzeichen" / "datei ä.html"
    href = combined_search.link(p, tmp_path)
    assert href == "export/Ordner%20mit%20Leerzeichen/datei%20%C3%A4.html"


def test_combined_unfold_und_unescape():
    assert combined_search._unfold("A:1\r\n b\nB:2\n\tc") == ["A:1b", "B:2c"]
    assert combined_search._unescape(r"a\,b\;c\nd\\e") == "a,b;c\nd\\e"


def test_combined_prop_pval_demail():
    name, params, value = combined_search._prop(
        'ORGANIZER;CN="Alice; Ex":mailto:alice@example.com')
    assert name == "ORGANIZER"
    assert combined_search._pval(params, "CN") == "Alice; Ex"  # Anführungszeichen schützen ;
    assert value == "mailto:alice@example.com"
    assert combined_search._prop("zeile ohne doppelpunkt") == (None, None, None)
    assert combined_search._pval(";CN=Bob", "CN") == "Bob"
    assert combined_search._pval("", "CN") == ""
    assert combined_search._demail("MAILTO:Alice@Example.com") == "Alice@Example.com"
    assert combined_search._demail(None) == ""


def test_combined_ics_when_varianten():
    ts, disp = combined_search._ics_when("20250601", dateonly=True)
    assert disp == "2025-06-01" and ts is not None
    ts, disp = combined_search._ics_when("20250601T120000", dateonly=False)
    assert disp == "2025-06-01 12:00" and ts is not None
    ts, disp = combined_search._ics_when("20250601T120000Z", dateonly=False)
    assert ts is not None and re.fullmatch(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}", disp)
    assert combined_search._ics_when("", dateonly=False) == (None, "")
    assert combined_search._ics_when("unsinn", dateonly=False) == (None, "unsinn")


# --------------------------------------------------------------------------
# combined_search: Einleser (Teams, Mail, Kalender, Kontakte)
# --------------------------------------------------------------------------
def test_combined_read_teams_kategorien_links_und_cap(tmp_path):
    root = tmp_path / "teams_export"
    (root / "1on1").mkdir(parents=True)
    (root / "1on1" / "alice__abc.html").write_text(
        TEAMS_HTML.replace("Danke!", "a" * 4100), encoding="utf-8")
    kanal = root / "channels" / "Team Rocket"
    kanal.mkdir(parents=True)
    (kanal / "general__1.html").write_text(
        TEAMS_HTML.replace("Projekt Alpha", "Allgemein"), encoding="utf-8")
    # werden übersprungen: Index-/Suchseite und versteckte Ordner
    (root / "index.html").write_text("<html></html>", encoding="utf-8")
    (root / "search.html").write_text("<html></html>", encoding="utf-8")
    (root / ".imgcache").mkdir()
    (root / ".imgcache" / "bild.html").write_text(TEAMS_HTML, encoding="utf-8")

    people = set()
    recs = combined_search.read_teams(root.resolve(), tmp_path.resolve(), people)

    assert len(recs) == 4                        # je 2 Nachrichten, Cache/Index ignoriert
    eins = [r for r in recs if r["ctx"] == "1:1-Chat"]
    kanaele = [r for r in recs if r["ctx"] == "Kanal: Allgemein"]
    assert len(eins) == 2 and len(kanaele) == 2
    assert {"Alice Example", "Bob"} <= people
    assert eins[0]["who"] == "Alice Example"
    assert eins[0]["d"] == "2025-06-01 09:30" and eins[0]["ts"] is not None
    assert "alice example" in eins[0]["ppl"]
    # BODY_CAP: lange Nachricht wird gekappt
    bob = [r for r in eins if r["who"] == "Bob"][0]
    assert len(bob["x"]) == combined_search.BODY_CAP
    # Link: relativ zum Ausgabeordner, Segmente URL-kodiert
    assert kanaele[0]["p"] == "teams_export/channels/Team%20Rocket/general__1.html"


def test_combined_read_outlook_ordner_und_personen(tmp_path):
    root = tmp_path / "outlook_export"
    post = root / "E-Mail" / "Posteingang"
    post.mkdir(parents=True)
    (post / "mail.eml").write_bytes(make_eml(cc="Carol <carol@example.com>"))
    (root / "wurzel.eml").write_bytes(make_eml(subject="Wurzelmail"))
    (root / "kaputt.eml").mkdir()                # unlesbar -> wird übersprungen

    people = set()
    recs = combined_search.read_outlook(root.resolve(), tmp_path.resolve(), people)

    assert len(recs) == 2
    by_title = {r["title"]: r for r in recs}
    assert by_title["Testmail"]["ctx"] == "Posteingang"     # "E-Mail/" nur Anzeige entfernt
    assert by_title["Wurzelmail"]["ctx"] == "(Stamm)"
    r = by_title["Testmail"]
    assert r["src"] == "outlook" and r["who"] == "Alice Example"
    assert "carol@example.com" in r["ppl"]
    assert "hier der Inhalt" in r["x"]
    assert r["ts"] is not None
    assert re.fullmatch(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}", r["d"])
    assert r["p"] == "outlook_export/E-Mail/Posteingang/mail.eml"
    assert {"Alice Example", "Bob Builder", "carol@example.com"} <= people


def test_combined_read_calendar(tmp_path):
    root = tmp_path / "outlook_export"
    d = root / "kalender" / "Arbeit"
    d.mkdir(parents=True)
    (d / "termin.ics").write_text(ICS, encoding="utf-8")

    people = set()
    recs = combined_search.read_calendar(root.resolve(), tmp_path.resolve(), people)

    assert len(recs) == 1
    r = recs[0]
    assert r["src"] == "kalender"
    assert r["title"] == "Planung, Quartal"      # \\, entschärft
    assert r["ctx"] == "Kalender: Arbeit"
    assert r["who"] == "Alice Example"
    assert r["x"].startswith("Ort: Raum 42.")
    assert "bob@example.com" in r["ppl"]
    assert r["ts"] is not None
    assert {"Alice Example", "Bob Builder",
            "alice@example.com", "bob@example.com"} <= people


def test_combined_read_contacts(tmp_path):
    root = tmp_path / "outlook_export"
    d = root / "kontakte" / "Team"
    d.mkdir(parents=True)
    (d / "alice.vcf").write_text(VCF, encoding="utf-8")
    (root / "erika.vcf").write_text(
        "\r\n".join(["BEGIN:VCARD", "N:Muster;Erika;;;", "END:VCARD"]), encoding="utf-8")

    people = set()
    recs = combined_search.read_contacts(root.resolve(), tmp_path.resolve(), people)

    assert len(recs) == 2
    by_title = {r["title"]: r for r in recs}
    alice = by_title["Alice Example"]
    assert alice["ctx"] == "Kontakte: Team"
    assert alice["who"] == "Firma GmbH · Entwicklung"
    assert "Firma GmbH · Entwicklung" in alice["x"]
    assert "Erste Zeileweiter gefaltet" in alice["x"]   # RFC-Zeilenfaltung aufgelöst
    assert alice["ts"] is None and alice["d"] == ""
    assert "alice@example.com" in alice["ppl"]
    # ohne FN: Name aus N-Property (Vorname Nachname), generischer Rest
    erika = by_title["Erika Muster"]
    assert erika["ctx"] == "Kontakte" and erika["who"] == "Kontakt"
    assert {"Alice Example", "alice@example.com", "Erika Muster"} <= people


# --------------------------------------------------------------------------
# combined_search: build() und main()
# --------------------------------------------------------------------------
def test_combined_build_gesamtseite(tmp_path):
    teams = tmp_path / "teams_export"
    (teams / "1on1").mkdir(parents=True)
    (teams / "1on1" / "alice__abc.html").write_text(TEAMS_HTML, encoding="utf-8")
    outlook = tmp_path / "outlook_export"
    post = outlook / "E-Mail" / "Posteingang"
    post.mkdir(parents=True)
    (post / "mail.eml").write_bytes(make_eml(
        subject=Header("Grüße 🎉 Bericht", "utf-8").encode(),
        body="Zusammenfassung folgt. " + EVIL_SNIPPET))
    kal = outlook / "kalender" / "Arbeit"
    kal.mkdir(parents=True)
    (kal / "termin.ics").write_text(ICS, encoding="utf-8")
    kon = outlook / "kontakte" / "Team"
    kon.mkdir(parents=True)
    (kon / "alice.vcf").write_text(VCF, encoding="utf-8")

    ziel = tmp_path / "combined_search.html"
    out, counts = combined_search.build(str(teams), str(outlook), ziel)

    assert out == ziel and ziel.is_file()
    assert counts == {"teams": 2, "outlook": 1, "kalender": 1, "kontakte": 1}
    html, idx = read_page(ziel)
    assert html.startswith("<!DOCTYPE html>")
    assert "Teams + Outlook · Suche" in html     # Template-Gerüst
    # Escaping: kein rohes </script> aus Nachrichteninhalten im HTML
    assert EVIL_SNIPPET not in html
    assert "\\u003c/script" in html
    # Umlaute/Emoji unkodiert im HTML (ensure_ascii=False)
    assert "Grüße 🎉 Bericht" in html and "bis morgen 🎉" in html

    recs = idx["recs"]
    assert len(recs) == 5
    # Sortierung: Zeitstempel absteigend, undatierte (Kontakte) zuletzt
    ts = [r["ts"] for r in recs]
    datiert = [t for t in ts if t is not None]
    assert datiert == sorted(datiert, reverse=True)
    assert ts[-1] is None and recs[-1]["src"] == "kontakte"
    mail = [r for r in recs if r["src"] == "outlook"][0]
    assert mail["title"] == "Grüße 🎉 Bericht"
    assert EVIL_SNIPPET in mail["x"]             # Inhalt unversehrt im Index
    assert mail["p"] == "outlook_export/E-Mail/Posteingang/mail.eml"
    teams_recs = [r for r in recs if r["src"] == "teams"]
    assert len(teams_recs) == 2
    assert teams_recs[0]["ctx"] == "1:1-Chat"
    assert [r for r in recs if r["src"] == "kalender"][0]["title"] == "Planung, Quartal"
    # Personenliste: case-insensitiv sortiert, Namen und Adressen enthalten
    people = idx["people"]
    assert people == sorted(people, key=str.lower)
    assert "Alice Example" in people and "bob@example.com" in people


def test_combined_build_ohne_teams_ordner(tmp_path, capsys):
    outlook = tmp_path / "outlook_export"
    outlook.mkdir()
    (outlook / "mail.eml").write_bytes(make_eml())
    ziel = tmp_path / "nur_outlook.html"
    _, counts = combined_search.build(str(tmp_path / "fehlt"), str(outlook), ziel)
    assert counts["teams"] == 0 and counts["outlook"] == 1
    assert "übersprungen" in capsys.readouterr().out
    _, idx = read_page(ziel)
    assert [r["src"] for r in idx["recs"]] == ["outlook"]


def test_combined_build_leere_ordner(tmp_path):
    (tmp_path / "teams_export").mkdir()
    (tmp_path / "outlook_export").mkdir()
    ziel = tmp_path / "leer.html"
    _, counts = combined_search.build(str(tmp_path / "teams_export"),
                                      str(tmp_path / "outlook_export"), ziel)
    assert sum(counts.values()) == 0
    _, idx = read_page(ziel)
    assert idx["recs"] == [] and idx["people"] == []


def test_combined_main_default_ausgabe(tmp_path, monkeypatch):
    t = tmp_path / "teams_export" / "1on1"
    t.mkdir(parents=True)
    (t / "a__1.html").write_text(TEAMS_HTML, encoding="utf-8")
    o = tmp_path / "outlook_export" / "Posteingang"
    o.mkdir(parents=True)
    (o / "m.eml").write_bytes(make_eml())
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(sys, "argv", ["combined_search.py"])
    combined_search.main()
    # Default-Ausgabe im gemeinsamen übergeordneten Ordner beider Exporte
    out = tmp_path / "combined_search.html"
    assert out.is_file()
    _, idx = read_page(out)
    assert len(idx["recs"]) == 3                 # 2 Teams-Nachrichten + 1 Mail
    assert {r["src"] for r in idx["recs"]} == {"teams", "outlook"}


def test_combined_main_mit_output_flag(tmp_path, monkeypatch):
    outlook = tmp_path / "outlook_export"
    outlook.mkdir()
    (outlook / "m.eml").write_bytes(make_eml())
    ziel = tmp_path / "ergebnis.html"
    monkeypatch.setattr(sys, "argv", ["combined_search.py", str(tmp_path / "fehlt"),
                                      str(outlook), "-o", str(ziel)])
    combined_search.main()
    assert ziel.is_file()
    _, idx = read_page(ziel)
    assert [r["src"] for r in idx["recs"]] == ["outlook"]


def test_combined_main_ohne_ordner_bricht_ab(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(sys, "argv", ["combined_search.py"])
    with pytest.raises(SystemExit, match="nichts zu tun"):
        combined_search.main()


# --------------------------------------------------------------------------
# Regression: "<" wird komplett als < eingebettet (nicht nur "</")
# --------------------------------------------------------------------------
KOMMENTAR_ANGRIFF = "Beispiel <!--<script>alert(2)</script> Ende"

KOMMENTAR_TEAMS_HTML = """<html><body>
<h1>Chat</h1>
<div class="msg">
  <span class="name">Mallory</span>
  <span class="time">2025-06-02 08:00</span>
  <div class="body">Beispiel &lt;!--&lt;script&gt;alert(2)&lt;/script&gt; Ende</div>
</div>
</body></html>"""


def _payload(path):
    m = IDX_RE.search(path.read_text(encoding="utf-8"))
    assert m is not None
    return m.group(1)


def test_generator_escapet_jedes_kleiner_zeichen(tmp_path):
    """'<!--' + '<script' im Inhalt bricht sonst den Script-Block auf.

    Der Browser wechselt bei '<!--' gefolgt von '<script' in den
    "double-escaped"-Zustand und übersieht das echte '</script>' – deshalb
    wird jedes '<' als \\u003c eingebettet; im Payload darf keines übrig sein.
    """
    teams = tmp_path / "teams" / "1on1"
    teams.mkdir(parents=True)
    (teams / "mallory__1.html").write_text(KOMMENTAR_TEAMS_HTML, encoding="utf-8")
    outlook = tmp_path / "outlook"
    outlook.mkdir()
    (outlook / "mail.eml").write_bytes(make_eml(body=KOMMENTAR_ANGRIFF))

    ziel = tmp_path / "combined.html"
    combined_search.build(str(tmp_path / "teams"), str(outlook), ziel)
    assert "<" not in _payload(ziel)
    _, idx = read_page(ziel)
    treffer = [r for r in idx["recs"] if KOMMENTAR_ANGRIFF in r["x"]]
    assert len(treffer) == 2                     # Teams-Nachricht und Mail
