# -*- coding: utf-8 -*-
import base64
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import sigcore  # noqa: E402

PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==")

WORD_HTML = """<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 15">
<link rel=File-List href="Societe_fichiers/filelist.xml">
<!--[if gte mso 9]><xml>
 <o:OfficeDocumentSettings><o:AllowPNG/></o:OfficeDocumentSettings>
</xml><![endif]-->
<style>p.MsoNormal {margin:0cm;}</style>
</head>
<body lang=FR>
<table><tr><td><p class=MsoNormal>Henri Jausions<o:p></o:p></p></td><td>Ing\xe9nieur</td></tr></table>
<p class=MsoNormal><!--[if gte vml 1]><v:shape id="Image_x0020_1" style='width:50pt'>
 <v:imagedata src="Societe_fichiers/image001.png" o:title=""/>
</v:shape><![endif]--><![if !vml]><img width=67 height=67 src="Societe_fichiers/image002.png" v:shapes="Image_x0020_1"><![endif]><o:p></o:p></p>
<p><a href="https://exemple.fr">exemple.fr</a> <a href="Plan.pdf">plan</a></p>
</body></html>"""


def make_folder(tmp_path, html=WORD_HTML, name="Societe"):
    d = tmp_path / "sig"
    d.mkdir()
    (d / f"{name}.htm").write_bytes(html.encode("cp1252"))
    files = d / f"{name}_fichiers"
    files.mkdir()
    (files / "image001.png").write_bytes(PNG)
    (files / "image002.png").write_bytes(PNG)
    return d


def test_scan_finds_files_dir(tmp_path):
    sigs = sigcore.scan_folder(make_folder(tmp_path))
    assert len(sigs) == 1
    assert sigs[0].files_dir.name == "Societe_fichiers"
    assert sigs[0].image_count == 2


def test_duplicate_htm_html(tmp_path):
    d = make_folder(tmp_path)
    (d / "Societe.html").write_text("<p>x</p>")
    sigs = sigcore.scan_folder(d)
    assert len(sigs) == 1 and sigs[0].htm.suffix == ".htm"
    assert any("double" in w for w in sigs[0].warnings)


def test_clean_removes_vml_and_office_tags(tmp_path):
    sig = sigcore.scan_folder(make_folder(tmp_path))[0]
    warns = []
    html = sigcore.prepare_html(sig, warns)
    assert "v:shape" not in html and "v:imagedata" not in html
    assert "<o:p>" not in html and "<![if" not in html and "<![endif]>" not in html
    assert "filelist.xml" not in html and "ProgId" not in html and "OfficeDocumentSettings" not in html
    # une seule image (la version VML en double a disparu), integree en base64
    assert html.count("data:image/png;base64,") == 1
    assert "charset=utf-8" in html
    assert "Ingénieur" in html


def test_warnings_missing_image_and_local_link(tmp_path):
    d = make_folder(tmp_path)
    (d / "Societe_fichiers" / "image002.png").unlink()
    sig = sigcore.scan_folder(d)[0]
    size, warns = sigcore.analyse(sig)
    assert size > 0
    assert any("introuvable" in w for w in warns)
    assert any("Plan.pdf" in w for w in warns)


def test_unsupported_format(tmp_path):
    html = '<html><body><img src="S_files/logo.emf"></body></html>'
    d = tmp_path / "sig"
    d.mkdir()
    (d / "S.htm").write_text(html)
    (d / "S_files").mkdir()
    (d / "S_files" / "logo.emf").write_bytes(b"x")
    _, warns = sigcore.analyse(sigcore.scan_folder(d)[0])
    assert any("non affiche" in w for w in warns)


def test_text_separates_table_cells(tmp_path):
    sig = sigcore.scan_folder(make_folder(tmp_path))[0]
    text = sigcore.prepare_text(sig)
    assert "Henri Jausions" in text and "Ingénieur" in text
    d2 = tmp_path / "t"
    d2.mkdir()
    (d2 / "T.htm").write_text("<table><tr><td>Tel</td><td>05 49</td></tr></table>")
    assert sigcore.prepare_text(sigcore.scan_folder(d2)[0]).startswith("Tel 05 49")
    assert text.endswith("\r\n")


def test_validate_name():
    assert sigcore.validate_name("Societe") == []
    assert sigcore.validate_name("a" * 150)
    assert sigcore.validate_name("a:b")


def test_build_items_learned_format():
    fmt = {"count": 1, "secondary_keys": ["htm", "txt"], "sample": {"source": "Outlook"}}
    items = sigcore.build_items("S", "a@b.fr", "<p>h</p>", "t", fmt)
    assert [i["secondaryKey"] for i in items] == ["htm", "txt"]
    assert items[0]["source"] == "Outlook" and items[1]["value"] == "t"


CLOUD = {
    "account": {"status": 200, "body": {"settings": [
        {"name": "Societe", "parentSetting": "roaming_signature_list", "secondaryKey": "htm",
         "value": "<p>x</p>", "timestamp": 1, "type": "Blob", "source": "UserOverride"},
        {"name": "Societe", "parentSetting": "roaming_signature_list", "secondaryKey": "txt",
         "value": "x", "timestamp": 1, "type": "Blob", "source": "UserOverride"},
        {"name": "new_signature", "parentSetting": "", "value": "Societe", "timestamp": 1,
         "type": "String", "scope": "a@b.fr"},
        {"name": "reply_signature", "parentSetting": "", "value": "", "timestamp": 1,
         "type": "String", "scope": "a@b.fr"},
    ]}},
}


def test_learn_and_names():
    assert sigcore.extract_signature_names(CLOUD) == ["Societe"]
    fmt = sigcore.learn_format(CLOUD)
    assert fmt["count"] == 2 and fmt["secondary_keys"] == ["htm", "txt"]
    assert set(fmt["defaults"]) == {"new", "reply"}


def test_build_default_items():
    fmt = sigcore.learn_format(CLOUD)
    items = sigcore.build_default_items(fmt["defaults"], "a@b.fr", "Court", "Court")
    assert [i["name"] for i in items] == ["new_signature", "reply_signature"]
    assert all(i["value"] == "Court" for i in items)
    assert sigcore.build_default_items({}, "a@b.fr", "Court", None) == []


def test_restore_items():
    items = sigcore.restore_items(CLOUD)
    assert len(items) == 2 and all(i["timestamp"] > 1 for i in items)
