# -*- coding: utf-8 -*-
"""Tests de la logique d'import et de la ligne de commande, sans reseau."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import sig_o365  # noqa: E402
from test_sigcore import CLOUD, make_folder  # noqa: E402


class FakeResponse:
    def __init__(self, status=200):
        self.status_code, self.ok, self.text = status, status < 400, "{}"


class FakeApi:
    def __init__(self):
        self.uploads = []
        self.state = json.loads(json.dumps(CLOUD))

    def read(self, mailbox):
        return self.state

    def upload(self, mailbox, items):
        self.uploads.append(items)
        lst = self.state["account"]["body"]["settings"]
        lst.extend(i for i in items if i.get("parentSetting") == "roaming_signature_list")
        return FakeResponse()


def importer(tmp_path, monkeypatch, api=None):
    monkeypatch.setattr(sig_o365, "EXPORT_DIR", tmp_path / "exports")
    imp = sig_o365.Importer(None, "a@b.fr", lambda m: None)
    imp.api = api
    return imp


def test_simulation_without_login(tmp_path, monkeypatch):
    imp = importer(tmp_path, monkeypatch)
    sigs = sig_o365.scan_folder(make_folder(tmp_path))
    ok, failed = imp.send(sigs, dry=True)
    assert ok == 1 and not failed
    data = json.loads((tmp_path / "exports" / "simulation_Societe.json").read_text(encoding="utf-8"))
    assert {i["secondaryKey"] for i in data} == {"htm", "txt"}


def test_real_send_backs_up_sets_defaults_and_verifies(tmp_path, monkeypatch):
    monkeypatch.setattr(sig_o365.time, "sleep", lambda s: None)
    api = FakeApi()
    imp = importer(tmp_path, monkeypatch, api)
    sigs = sig_o365.scan_folder(make_folder(tmp_path))
    ok, failed = imp.send(sigs, dry=False, new_default="Societe", reply_default="")
    assert ok == 1 and not failed
    assert list((tmp_path / "exports").glob("sauvegarde_a@b.fr_*.json"))
    assert len(api.uploads) == 2                      # signature puis defauts
    assert [i["name"] for i in api.uploads[1]] == ["new_signature", "reply_signature"]
    good, bad = imp.verify(sigs)
    assert good == ["Societe"] and not bad


def test_restore(tmp_path, monkeypatch):
    api = FakeApi()
    imp = importer(tmp_path, monkeypatch, api)
    f = tmp_path / "backup.json"
    f.write_text(json.dumps(CLOUD), encoding="utf-8")
    assert imp.restore(f, dry=True) == 1 and not api.uploads
    assert imp.restore(f, dry=False) == 1 and len(api.uploads[0]) == 2


def test_cli_list_and_simulation(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(sig_o365, "EXPORT_DIR", tmp_path / "exports")
    monkeypatch.setattr(sig_o365, "CONFIG_FILE", tmp_path / "config.json")
    folder = make_folder(tmp_path)
    assert sig_o365.run_cli(["--folder", str(folder), "--list"]) == 0
    assert "Societe" in capsys.readouterr().out
    assert sig_o365.run_cli(["--folder", str(folder), "--all"]) == 0
    assert (tmp_path / "exports" / "simulation_Societe.json").exists()
    assert sig_o365.run_cli(["--folder", str(folder), "--name", "Inconnue"]) == 2


class FakeAuth:
    claims, upn, token = {}, "a@b.fr", "jeton"


def test_retry_on_429(monkeypatch):
    import requests
    calls = []

    class R:
        def __init__(self, code):
            self.status_code, self.headers = code, {"Retry-After": "1"}

    def fake_request(method, url, **kw):
        calls.append(kw["headers"]["x-anchormailbox"])
        return R(429 if len(calls) < 3 else 200)
    monkeypatch.setattr(requests, "request", fake_request)
    monkeypatch.setattr(sig_o365.time, "sleep", lambda s: None)
    api = sig_o365.CloudSettings(FakeAuth(), lambda m: None)
    assert api._request("GET", "https://x", "c@d.fr").status_code == 200
    assert calls == ["SMTP:c@d.fr"] * 3
