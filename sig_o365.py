# -*- coding: utf-8 -*-
"""
Signatures O365 - envoie un dossier de signatures Outlook (.htm/.txt + dossier d'images)
directement dans la boite Exchange Online (signatures "cloud" / roaming), sans passer
par le dossier local %APPDATA%\\Microsoft\\Signatures.

ATTENTION : l'API utilisee (substrate.office.com / outlookcloudsettings) est celle
qu'utilisent Outlook et OWA en interne. Elle n'est PAS documentee par Microsoft et peut
changer sans preavis. Tester d'abord sur une boite de test, en mode simulation.

Lancement sans argument : interface graphique. Avec arguments : ligne de commande
(voir `SignaturesO365.exe --help`).
"""

import argparse
import base64
import json
import os
import queue
import re
import sys
import threading
import time
import webbrowser
from pathlib import Path
from typing import Optional

import sigcore
from sigcore import (analyse, build_default_items, build_items, extract_signature_names,
                     learn_format, local_default_signatures, local_signature_dir,
                     prepare_html, prepare_text, restore_items, roaming_disabled_locally,
                     scan_folder)

APP_TITLE = "Signatures O365 - import dans la boite mail"
APP_VERSION = "1.9.1"
BOX_ON, BOX_OFF = "☑", "☐"

# Client public "Microsoft Office" (celui qu'utilise Outlook pour ce service).
DEFAULT_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
SCOPES = ["https://outlook.office.com/.default"]
SUBSTRATE = "https://substrate.office.com/ows/beta/outlookcloudsettings/settings"
USER_AGENT = "Microsoft Office/16.0 (Windows NT 10.0; Microsoft Outlook 16.0; Pro)"
RETRY_STATUS = {429, 500, 502, 503, 504}


# --------------------------------------------------------------------------- chemins / journal
def base_dir() -> Path:
    """Dossier de l'exe (PyInstaller) ou du script : l'outil reste portable."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


DATA_DIR = base_dir() / "data"
EXPORT_DIR = base_dir() / "exports"
CONFIG_FILE = DATA_DIR / "config.json"


def safe_name(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|\s]+', "_", s).strip("_") or "x"


def write_journal(line: str):
    try:
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        with open(EXPORT_DIR / f"journal_{time.strftime('%Y%m%d')}.txt", "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


# --------------------------------------------------------------------------- erreurs
AADSTS_HELP = {
    "53003": "Connexion bloquee par une strategie d'acces conditionnel du tenant. Essayez l'autre "
             "mode de connexion, ou demandez a l'administrateur d'autoriser ce poste.",
    "65001": "Consentement requis : un administrateur doit autoriser l'application dans le tenant.",
    "50126": "Identifiant ou mot de passe incorrect.",
    "50020": "Ce compte n'appartient pas au tenant indique (compte personnel ou invite ?).",
    "50076": "Authentification multifacteur exigee : refaites la connexion et validez la MFA.",
    "50079": "Inscription a la MFA exigee : terminez-la dans le navigateur puis recommencez.",
    "50105": "Le compte n'est pas autorise a utiliser cette application dans ce tenant.",
    "70016": "Le code n'a pas encore ete saisi.",
    "70019": "Le code a expire : relancez la connexion et saisissez-le dans les 15 minutes.",
}


def explain_error(msg: str) -> str:
    m = re.search(r"AADSTS(\d+)", msg)
    if m and m.group(1) in AADSTS_HELP:
        return f"{AADSTS_HELP[m.group(1)]}\n\nDetail technique :\n{msg[:600]}"
    low = msg.lower()
    if "expired_token" in low or "code_expired" in low:
        return "Le code a expire : relancez la connexion.\n\n" + msg[:600]
    if "proxy" in low or "connection" in low or "ssl" in low:
        return ("Impossible de joindre les services Microsoft (proxy, pare-feu, filtrage SSL ?).\n\n"
                + msg[:600])
    return msg[:800]


# --------------------------------------------------------------------------- cache chiffre
def _dpapi(data: bytes, protect: bool) -> bytes:
    """Chiffrement Windows DPAPI lie a la session utilisateur (illisible sur un autre compte/PC)."""
    import ctypes
    from ctypes import wintypes

    class BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    buf = ctypes.create_string_buffer(data, len(data))
    blob_in = BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    blob_out = BLOB()
    crypt32, kernel32 = ctypes.windll.crypt32, ctypes.windll.kernel32
    fn = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
    ok = fn(ctypes.byref(blob_in), None, None, None, None, 0x1, ctypes.byref(blob_out))
    if not ok:
        raise OSError("DPAPI : echec du (de)chiffrement")
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        kernel32.LocalFree(blob_out.pbData)


class TokenStore:
    """Cache MSAL chiffre par DPAPI (Windows). Ailleurs : fichier en clair, droits 600."""

    def __init__(self):
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        self.win = sys.platform == "win32"
        self.file = DATA_DIR / ("token_cache.dpapi" if self.win else "token_cache.bin")
        self.legacy = DATA_DIR / "token_cache.bin"

    def load(self) -> Optional[str]:
        try:
            if self.file.exists():
                raw = self.file.read_bytes()
                return (_dpapi(raw, False) if self.win else raw).decode("utf-8")
            if self.win and self.legacy.exists():   # cache en clair des versions <= 1.8
                return self.legacy.read_text()
        except Exception:
            return None
        return None

    def save(self, text: str):
        raw = text.encode("utf-8")
        if self.win:
            self.file.write_bytes(_dpapi(raw, True))
            if self.legacy.exists():
                try:
                    self.legacy.unlink()
                except OSError:
                    pass
        else:
            self.file.write_bytes(raw)
            try:
                os.chmod(self.file, 0o600)
            except OSError:
                pass

    def clear(self):
        for f in (self.file, self.legacy):
            try:
                f.unlink()
            except OSError:
                pass


# --------------------------------------------------------------------------- auth
def decode_jwt(token: str) -> dict:
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload))
    except Exception:
        return {}


class Auth:
    def __init__(self, client_id: str, tenant: str, use_broker: bool = True):
        import msal
        self.store = TokenStore()
        self.cache = msal.SerializableTokenCache()
        state = self.store.load()
        if state:
            try:
                self.cache.deserialize(state)
            except Exception:
                pass
        authority = f"https://login.microsoftonline.com/{tenant or 'organizations'}"
        self.broker = False
        self.app = None
        if use_broker and sys.platform == "win32":
            # Gestionnaire de comptes Windows (WAM) : meme mecanisme qu'Outlook/Office
            try:
                import pymsalruntime  # noqa: F401  (verifie la presence de WAM)
                self.app = msal.PublicClientApplication(
                    client_id, authority=authority, token_cache=self.cache,
                    enable_broker_on_windows=True)
                self.broker = True
            except Exception:
                self.app = None
        if self.app is None:
            self.app = msal.PublicClientApplication(client_id, authority=authority, token_cache=self.cache)
        self.result = None

    def _save(self):
        if self.cache.has_state_changed:
            self.store.save(self.cache.serialize())

    def _check(self, r):
        if not r or "access_token" not in r:
            r = r or {}
            raise RuntimeError(r.get("error_description") or r.get("error") or "Echec d'authentification")
        self.result = r
        self._save()
        return r

    def silent(self, username: Optional[str] = None):
        for acc in self.app.get_accounts(username=username or None):
            r = self.app.acquire_token_silent(SCOPES, account=acc)
            if r and "access_token" in r:
                return self._check(r)
        return None

    def device(self, show_code):
        flow = self.app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(flow.get("error_description", "Code d'appareil indisponible"))
        show_code(flow)
        return self._check(self.app.acquire_token_by_device_flow(flow))

    def interactive(self, parent_hwnd=None, login_hint: Optional[str] = None):
        kw = {"prompt": "select_account"}
        if login_hint:
            kw["login_hint"] = login_hint
        if self.broker:
            kw["parent_window_handle"] = parent_hwnd or self.app.CONSOLE_WINDOW_HANDLE
        return self._check(self.app.acquire_token_interactive(SCOPES, **kw))

    def logout(self):
        for acc in self.app.get_accounts():
            self.app.remove_account(acc)
        self.result = None
        self.store.clear()

    @property
    def token(self) -> str:
        r = self.silent() or self.result
        if not r:
            raise RuntimeError("Non connecte")
        return r["access_token"]

    @property
    def claims(self) -> dict:
        return decode_jwt(self.result["access_token"]) if self.result else {}

    @property
    def upn(self) -> str:
        c = self.claims
        return c.get("upn") or c.get("unique_name") or ""


# --------------------------------------------------------------------------- API cloud
class CloudSettings:
    def __init__(self, auth: Auth, log=print):
        self.auth = auth
        self.log = log

    def _headers(self, mailbox: str) -> dict:
        c = self.auth.claims
        me = self.auth.upn.lower()
        if mailbox.lower() == me and c.get("puid") and c.get("tid"):
            anchor = f"PUID:{c['puid']}@{c['tid']}"
        else:
            anchor = f"SMTP:{mailbox}"
        return {
            "Authorization": f"Bearer {self.auth.token}",
            "Content-Type": "application/json", "Accept": "application/json",
            "User-Agent": USER_AGENT, "x-anchormailbox": anchor, "x-islargesetting": "true",
        }

    def _request(self, method: str, url: str, mailbox: str, tries: int = 4, **kw):
        import requests
        delay = 2
        for attempt in range(1, tries + 1):
            try:
                r = requests.request(method, url, headers=self._headers(mailbox), **kw)
            except requests.RequestException as e:
                if attempt == tries:
                    raise
                self.log(f"  reseau : {e.__class__.__name__}, nouvel essai dans {delay} s...")
            else:
                if r.status_code not in RETRY_STATUS or attempt == tries:
                    return r
                wait = r.headers.get("Retry-After")
                delay = int(wait) if wait and wait.isdigit() else delay
                self.log(f"  HTTP {r.status_code}, nouvel essai dans {delay} s...")
            time.sleep(delay)
            delay = min(delay * 2, 60)

    def read(self, mailbox: str) -> dict:
        out = {}
        for path in ("account", "global?force=true"):
            r = self._request("GET", f"{SUBSTRATE}/{path}", mailbox, timeout=60)
            try:
                body = r.json()
            except ValueError:
                body = r.text
            out[path] = {"status": r.status_code, "body": body}
        return out

    def upload(self, mailbox: str, items: list):
        data = json.dumps(items, ensure_ascii=False).encode("utf-8")
        return self._request("PATCH", f"{SUBSTRATE}/account?fromows=true", mailbox,
                             data=data, timeout=120)


def save_json(prefix: str, mailbox: str, data) -> Path:
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    out = EXPORT_DIR / f"{prefix}_{safe_name(mailbox)}_{time.strftime('%Y%m%d_%H%M%S')}.json"
    out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return out


def access_denied(res: dict) -> bool:
    return any(v.get("status") in (401, 403) for v in res.values())


# --------------------------------------------------------------------------- traitement commun
class Importer:
    """Logique d'import partagee entre l'interface et la ligne de commande."""

    def __init__(self, auth: Optional[Auth], mailbox: str, log):
        self.auth, self.mailbox, self.log = auth, mailbox, log
        self.api = CloudSettings(auth, log) if auth else None
        self.fmt = None
        self.cloud = None

    def read_cloud(self, prefix: str = "cloud") -> dict:
        self.cloud = self.api.read(self.mailbox)
        out = save_json(prefix, self.mailbox, self.cloud)
        for k, v in self.cloud.items():
            self.log(f"  GET {k} -> HTTP {v['status']}")
        if access_denied(self.cloud):
            raise RuntimeError("Acces refuse (HTTP 401/403) par le service des signatures cloud. "
                               "Deconnectez-vous puis reconnectez-vous avec le compte de la boite.")
        self.fmt = learn_format(self.cloud)
        names = extract_signature_names(self.cloud)
        self.log(f"Signatures cloud : {', '.join(names) if names else '(aucune)'}")
        if self.fmt["count"]:
            self.log(f"Format Outlook detecte : cles {self.fmt['secondary_keys']}, champs {self.fmt['sample']}")
        else:
            self.log("Aucune signature existante dans la boite : format par defaut utilise.")
        d = self.fmt["defaults"]
        self.log("Parametres 'signature par defaut' reperes : "
                 + (", ".join(f"{k}={v.get('name')}" for k, v in d.items()) if d else "aucun"))
        self.log(f"Copie enregistree : {out}")
        return self.cloud

    def send(self, sigs: list, dry: bool, new_default=None, reply_default=None) -> tuple:
        if self.api and self.cloud is None:
            self.log("Lecture prealable de la boite (format et sauvegarde)...")
            self.read_cloud(prefix="sauvegarde")
        elif self.api and not dry:
            self.log("Sauvegarde de la boite avant envoi...")
            self.read_cloud(prefix="sauvegarde")
        ok, failed = 0, []
        for s in sigs:
            try:
                warns = list(s.warnings)
                items = build_items(s.name, self.mailbox, prepare_html(s, warns), prepare_text(s), self.fmt)
                for w in warns:
                    self.log(f"  /!\\ {s.name} : {w}")
                if dry:
                    out = EXPORT_DIR / f"simulation_{safe_name(s.name)}.json"
                    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
                    out.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
                    self.log(f"[SIMULATION] {s.name} -> {out.name}")
                    ok += 1
                    continue
                r = self.api.upload(self.mailbox, items)
                if r.ok:
                    self.log(f"Envoye  {s.name}  (HTTP {r.status_code})")
                    ok += 1
                else:
                    failed.append(s.name)
                    self.log(f"ECHEC {s.name}  HTTP {r.status_code} : {r.text[:400]}")
            except Exception as e:
                failed.append(s.name)
                self.log(f"ECHEC {s.name} : {e}")
        self.log(f"Termine : {ok}/{len(sigs)} signature(s) traitee(s).")
        if new_default is not None or reply_default is not None:
            self.set_defaults(new_default, reply_default, dry)
        return ok, failed

    def set_defaults(self, new, reply, dry: bool) -> bool:
        defaults = (self.fmt or {}).get("defaults") or {}
        items = build_default_items(defaults, self.mailbox, new, reply)
        wanted = [k for k, v in (("new", new), ("reply", reply)) if v is not None]
        missing = [k for k in wanted if k not in defaults]
        if missing:
            self.log("Signature par defaut non definissable automatiquement "
                     f"({', '.join(missing)}) : le parametre n'a pas ete trouve dans la boite. "
                     "Choisissez-la une fois dans Outlook Web, puis relancez 'Lire le cloud' : "
                     "l'outil saura la modifier ensuite.")
        if not items:
            return False
        if dry:
            EXPORT_DIR.mkdir(parents=True, exist_ok=True)
            out = EXPORT_DIR / "simulation_defauts.json"
            out.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
            self.log(f"[SIMULATION] signatures par defaut -> {out.name}")
            return True
        r = self.api.upload(self.mailbox, items)
        self.log(("Signatures par defaut definies" if r.ok else "ECHEC signatures par defaut")
                 + f" (HTTP {r.status_code}) {'' if r.ok else r.text[:300]}")
        return r.ok

    def verify(self, sigs: list) -> tuple:
        time.sleep(3)
        present = {n.lower() for n in extract_signature_names(self.api.read(self.mailbox))}
        good = [x.name for x in sigs if x.name.lower() in present]
        bad = [x.name for x in sigs if x.name.lower() not in present]
        if good:
            self.log(f"  Visibles dans le cloud : {', '.join(good)}")
        if bad:
            self.log(f"  NON visibles dans le cloud : {', '.join(bad)}")
        return good, bad

    def restore(self, backup_file: Path, dry: bool) -> int:
        data = json.loads(Path(backup_file).read_text(encoding="utf-8"))
        items = restore_items(data)
        if not items:
            self.log("Aucune signature dans ce fichier de sauvegarde.")
            return 0
        names = sorted({sigcore._ci(i, "name") for i in items})
        self.log(f"Restauration de {len(names)} signature(s) : {', '.join(names)}")
        if dry:
            self.log("[SIMULATION] rien n'est envoye.")
            return len(names)
        r = self.api.upload(self.mailbox, items)
        self.log(f"Restauration : HTTP {r.status_code} {'' if r.ok else r.text[:300]}")
        return len(names) if r.ok else 0


def write_preview(sig) -> Path:
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    out = EXPORT_DIR / f"apercu_{safe_name(sig.name)}.html"
    out.write_text(prepare_html(sig), encoding="utf-8")
    return out


def load_config() -> dict:
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


# --------------------------------------------------------------------------- interface
def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinter.scrolledtext import ScrolledText

    class App:
        def __init__(self, root):
            self.root = root
            root.title(f"{APP_TITLE}  v{APP_VERSION}")
            root.geometry("1100x760")
            root.minsize(900, 600)
            self.q = queue.Queue()
            self.sigs = []
            self.checked = set()
            self.auth = None
            self.importer = None
            self.busy = False

            self.folder = tk.StringVar()
            self.mailbox = tk.StringVar()
            self.tenant = tk.StringVar(value="organizations")
            self.dry_run = tk.BooleanVar(value=True)
            self.def_new = tk.StringVar(value="(ne pas changer)")
            self.def_reply = tk.StringVar(value="(ne pas changer)")
            self.status = tk.StringVar(value="Non connecte")
            self._load_config()
            self._build()
            root.after(100, self._poll)
            self.log(f"Signatures O365 v{APP_VERSION} - exports et journaux : {EXPORT_DIR}")
            self.log("Mode simulation actif par defaut : rien n'est envoye tant qu'il est coche.")
            if roaming_disabled_locally():
                self.log("/!\\ Les signatures itinerantes sont desactivees dans Outlook classique sur ce "
                         "poste (DisableRoamingSignaturesTemporaryToggle=1) : il n'affichera pas les "
                         "signatures importees.")

        # ---- configuration
        def _load_config(self):
            cfg = load_config()
            for key, var in (("tenant", self.tenant), ("mailbox", self.mailbox), ("folder", self.folder)):
                if cfg.get(key):
                    var.set(cfg[key])

        def _save_config(self):
            try:
                DATA_DIR.mkdir(parents=True, exist_ok=True)
                CONFIG_FILE.write_text(json.dumps({
                    "tenant": self.tenant.get().strip(),
                    "mailbox": self.mailbox.get().strip(), "folder": self.folder.get().strip(),
                }, indent=2), encoding="utf-8")
            except Exception as e:
                self.log(f"Configuration non enregistree : {e}")

        # ---- construction
        def _build(self):
            pad = {"padx": 6, "pady": 4}
            f1 = ttk.LabelFrame(self.root, text="1. Dossier de signatures")
            f1.pack(fill="x", **pad)
            ttk.Entry(f1, textvariable=self.folder).pack(side="left", fill="x", expand=True, **pad)
            ttk.Button(f1, text="Parcourir...", command=self.browse).pack(side="left", **pad)
            ttk.Button(f1, text="Mes signatures Outlook", command=self.use_local).pack(side="left", **pad)
            ttk.Button(f1, text="Analyser", command=self.analyse).pack(side="left", **pad)

            f2 = ttk.Frame(self.root)
            f2.pack(fill="both", expand=True, **pad)
            cols = ("choix", "images", "txt", "taille", "alertes")
            self.tree = ttk.Treeview(f2, columns=cols, selectmode="browse", height=8)
            self.tree.heading("#0", text="Signature")
            for c, t, w, a in (("choix", "Importer", 70, "center"), ("images", "Images", 60, "center"),
                               ("txt", "TXT", 50, "center"), ("taille", "Taille finale", 90, "center"),
                               ("alertes", "Alertes", 420, "w")):
                self.tree.heading(c, text=t)
                self.tree.column(c, width=w, anchor=a)
            self.tree.column("#0", width=240)
            self.tree.tag_configure("warn", foreground="#b35c00")
            sb = ttk.Scrollbar(f2, orient="vertical", command=self.tree.yview)
            self.tree.configure(yscrollcommand=sb.set)
            self.tree.pack(side="left", fill="both", expand=True)
            sb.pack(side="left", fill="y")
            self.tree.bind("<Button-1>", self._on_click)
            self.tree.bind("<space>", lambda e: self._toggle(self.tree.focus()))
            self.tree.bind("<Double-1>", lambda e: self.preview())

            fsel = ttk.Frame(self.root)
            fsel.pack(fill="x", padx=6)
            ttk.Button(fsel, text="Tout cocher", command=lambda: self._set_all(True)).pack(side="left", padx=4)
            ttk.Button(fsel, text="Tout decocher", command=lambda: self._set_all(False)).pack(side="left", padx=4)
            ttk.Button(fsel, text="Inverser", command=self._invert).pack(side="left", padx=4)
            ttk.Button(fsel, text="Apercu", command=self.preview).pack(side="left", padx=4)
            self.count_lbl = ttk.Label(fsel, text="0 signature cochee", foreground="gray")
            self.count_lbl.pack(side="left", padx=12)

            f3 = ttk.LabelFrame(self.root, text="2. Compte Microsoft 365")
            f3.pack(fill="x", **pad)
            ttk.Label(f3, text="Boite cible :").grid(row=0, column=0, sticky="w", **pad)
            ttk.Entry(f3, textvariable=self.mailbox, width=40).grid(row=0, column=1, sticky="we", **pad)
            ttk.Button(f3, text="Se connecter (Office 365)",
                       command=lambda: self.login("windows")).grid(row=0, column=2, **pad)
            ttk.Button(f3, text="Deconnexion", command=self.logout).grid(row=0, column=3, **pad)
            adv = ttk.Frame(f3)
            adv.grid(row=1, column=0, columnspan=4, sticky="we")
            ttk.Label(adv, text="Tenant (domaine) :").pack(side="left", padx=6)
            ttk.Entry(adv, textvariable=self.tenant, width=30).pack(side="left")
            ttk.Label(adv, text="   Secours :", foreground="gray").pack(side="left", padx=6)
            ttk.Button(adv, text="Code d'appareil", command=lambda: self.login("device")).pack(side="left", padx=4)
            ttk.Label(f3, textvariable=self.status, foreground="#0a5").grid(row=2, column=0, columnspan=4,
                                                                           sticky="w", **pad)
            f3.columnconfigure(1, weight=1)

            f5 = ttk.LabelFrame(self.root, text="3. Signatures par defaut")
            f5.pack(fill="x", **pad)
            ttk.Label(f5, text="Nouveau message :").pack(side="left", padx=6)
            self.cb_new = ttk.Combobox(f5, textvariable=self.def_new, width=28, state="readonly")
            self.cb_new.pack(side="left", padx=4)
            ttk.Label(f5, text="Reponse / transfert :").pack(side="left", padx=6)
            self.cb_reply = ttk.Combobox(f5, textvariable=self.def_reply, width=28, state="readonly")
            self.cb_reply.pack(side="left", padx=4)
            ttk.Button(f5, text="Reprendre ceux d'Outlook", command=self.use_local_defaults).pack(side="left", padx=6)
            self._fill_default_lists()

            f4 = ttk.LabelFrame(self.root, text="4. Actions")
            f4.pack(fill="x", **pad)
            ttk.Button(f4, text="Lire les signatures du cloud", command=self.read_cloud).pack(side="left", **pad)
            ttk.Button(f4, text="Restaurer une sauvegarde...", command=self.restore).pack(side="left", **pad)
            ttk.Checkbutton(f4, text="Simulation (genere le JSON sans rien envoyer)",
                            variable=self.dry_run, command=self._update_send_label).pack(side="left", **pad)
            self.send_btn = ttk.Button(f4, command=self.send)
            self.send_btn.pack(side="right", **pad)
            self.progress = ttk.Progressbar(f4, mode="indeterminate", length=120)
            self.progress.pack(side="right", **pad)
            self._update_send_label()

            self.logbox = ScrolledText(self.root, height=12, font=("Consolas", 9))
            self.logbox.pack(fill="both", expand=True, **pad)

        def _update_send_label(self):
            self.send_btn.config(text="Simuler (rien n'est envoye)" if self.dry_run.get()
                                 else "ENVOYER dans la boite mail")

        # ---- utilitaires
        def log(self, msg):
            line = f"[{time.strftime('%H:%M:%S')}] {msg}"
            write_journal(line)
            self.q.put(("log", line))

        def _poll(self):
            try:
                while True:
                    kind, val = self.q.get_nowait()
                    if kind == "log":
                        self.logbox.insert("end", val + "\n")
                        self.logbox.see("end")
                    elif kind == "call":
                        val()
            except queue.Empty:
                pass
            self.root.after(100, self._poll)

        def ui(self, fn):
            self.q.put(("call", fn))

        def _set_busy(self, on):
            self.busy = on
            if on:
                self.progress.start(12)
                self.send_btn.state(["disabled"])
            else:
                self.progress.stop()
                self.send_btn.state(["!disabled"])

        def run_bg(self, fn):
            if self.busy:
                messagebox.showinfo(APP_TITLE, "Une operation est deja en cours.")
                return
            self._set_busy(True)

            def wrap():
                try:
                    fn()
                except Exception as e:
                    text = explain_error(str(e))
                    self.log(f"ERREUR : {e}")
                    self.ui(lambda: messagebox.showerror(APP_TITLE, text))
                finally:
                    self.ui(lambda: self._set_busy(False))
            threading.Thread(target=wrap, daemon=True).start()

        # ---- dossier
        def browse(self):
            d = filedialog.askdirectory(title="Dossier contenant les signatures")
            if d:
                self.folder.set(d)
                self.analyse()

        def use_local(self):
            d = local_signature_dir()
            if not d:
                messagebox.showwarning(APP_TITLE, "Dossier de signatures d'Outlook classique introuvable "
                                                  "sur ce poste (%APPDATA%\\Microsoft\\Signatures).")
                return
            self.folder.set(str(d))
            self.analyse()

        def analyse(self):
            folder = Path(self.folder.get().strip('" '))
            if not folder.is_dir():
                messagebox.showerror(APP_TITLE, "Dossier introuvable.")
                return
            self.sigs = scan_folder(folder)
            self.tree.delete(*self.tree.get_children())
            self.checked.clear()
            nwarn = 0
            for i, s in enumerate(self.sigs):
                try:
                    size, warns = analyse(s)
                    size_txt = f"{size / 1024:.0f} Ko"
                except Exception as e:
                    size_txt, warns = "erreur", [str(e)]
                nwarn += bool(warns)
                self.tree.insert("", "end", iid=str(i), text=s.name, tags=("warn",) if warns else (),
                                 values=(BOX_OFF, s.image_count, "oui" if s.txt else "auto",
                                         size_txt, " ; ".join(warns) or "-"))
            self._refresh_count()
            self._fill_default_lists()
            self.log(f"{len(self.sigs)} signature(s) trouvee(s) dans {folder} - cochez celles a importer.")
            if nwarn:
                self.log(f"{nwarn} signature(s) avec alertes (en orange) : verifiez-les avec 'Apercu'.")
            if not self.sigs:
                self.log("Aucun fichier .htm trouve : verifier le dossier.")
            self._save_config()

        def preview(self):
            iid = self.tree.focus()
            if not iid:
                messagebox.showinfo(APP_TITLE, "Selectionnez une signature dans la liste.")
                return
            out = write_preview(self.sigs[int(iid)])
            webbrowser.open(out.as_uri())
            self.log(f"Apercu ouvert : {out.name} (HTML tel qu'il sera envoye)")

        # ---- signatures par defaut
        NO_CHANGE = "(ne pas changer)"

        def _fill_default_lists(self):
            names = [self.NO_CHANGE, "(aucune)"] + [s.name for s in self.sigs]
            self.cb_new["values"] = names
            self.cb_reply["values"] = names
            for var in (self.def_new, self.def_reply):
                if var.get() not in names:
                    var.set(self.NO_CHANGE)

        def use_local_defaults(self):
            d = local_default_signatures()
            if not d:
                messagebox.showinfo(APP_TITLE, "Aucune signature par defaut trouvee dans Outlook classique.")
                return
            known = {s.name for s in self.sigs}
            for kind, var in (("new", self.def_new), ("reply", self.def_reply)):
                if d.get(kind):
                    if d[kind] in known:
                        var.set(d[kind])
                    else:
                        self.log(f"Signature par defaut Outlook '{d[kind]}' absente du dossier analyse.")
            self.log(f"Defauts Outlook classique : nouveau={d.get('new', '-')}, reponse={d.get('reply', '-')}")

        def _default_value(self, var):
            v = var.get()
            if v == self.NO_CHANGE:
                return None
            return "" if v == "(aucune)" else v

        # ---- cases a cocher
        def _on_click(self, event):
            if self.tree.identify_region(event.x, event.y) not in ("cell", "tree"):
                return
            iid = self.tree.identify_row(event.y)
            if iid:
                self._toggle(iid)

        def _toggle(self, iid, state=None):
            if not iid:
                return
            on = (iid not in self.checked) if state is None else state
            (self.checked.add if on else self.checked.discard)(iid)
            self.tree.set(iid, "choix", BOX_ON if on else BOX_OFF)
            self._refresh_count()

        def _set_all(self, on):
            for iid in self.tree.get_children():
                self._toggle(iid, on)

        def _invert(self):
            for iid in self.tree.get_children():
                self._toggle(iid)

        def _refresh_count(self):
            n = len(self.checked)
            self.count_lbl.config(text=f"{n} signature{'s' if n > 1 else ''} cochee{'s' if n > 1 else ''} "
                                       f"sur {len(self.sigs)}")

        def selected(self):
            return [self.sigs[int(i)] for i in sorted(self.checked, key=int)]

        # ---- auth
        def _show_code(self, flow):
            def show():
                self.root.clipboard_clear()
                self.root.clipboard_append(flow["user_code"])
                webbrowser.open(flow["verification_uri"])
                messagebox.showinfo(APP_TITLE, f"Le navigateur va s'ouvrir sur :\n{flow['verification_uri']}\n\n"
                                               f"Code (deja copie dans le presse-papiers) :\n\n{flow['user_code']}\n\n"
                                               "Connectez-vous avec le compte de la boite mail, puis revenez ici.")
            self.ui(show)
            self.log(f"Code de connexion : {flow['user_code']}  ({flow['verification_uri']})")

        def login(self, mode):
            hwnd = None
            if sys.platform == "win32":
                try:
                    import ctypes
                    hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id()) or None
                except Exception:
                    hwnd = None
            tenant = self.tenant.get().strip() or "organizations"
            hint = self.mailbox.get().strip() or None
            self._save_config()
            labels = {"windows": "Office 365 (fenetre Windows)", "device": "code d'appareil"}

            def job():
                self.auth = Auth(DEFAULT_CLIENT_ID, tenant, use_broker=(mode == "windows"))
                self.importer = None
                self.log(f"Connexion {labels[mode]} - tenant {tenant}...")
                if mode == "windows" and sys.platform == "win32" and not self.auth.broker:
                    raise RuntimeError("Composant de connexion Windows (pymsalruntime) absent de l'exe. "
                                       "Recompilez avec build.bat. En attendant, utilisez 'Code d'appareil'.")
                r = self.auth.silent(hint)
                if not r:
                    r = (self.auth.device(self._show_code) if mode == "device"
                         else self.auth.interactive(hwnd, hint))
                c, upn = self.auth.claims, self.auth.upn
                self.ui(lambda: (self.status.set(f"Connecte : {upn}  (tenant {c.get('tid', '?')})"),
                                 self.mailbox.get() or self.mailbox.set(upn)))
                self.log(f"Connecte en tant que {upn}")
                self.ui(self._save_config)
            self.run_bg(job)

        def logout(self):
            if self.auth:
                self.auth.logout()
            else:
                TokenStore().clear()
            self.auth = None
            self.importer = None
            self.status.set("Non connecte")
            self.log("Deconnecte (cache de jeton supprime).")

        def _need_login(self) -> bool:
            if not self.auth or not self.auth.result:
                messagebox.showwarning(APP_TITLE, "Connectez-vous d'abord.")
                return True
            if not self.mailbox.get().strip():
                messagebox.showwarning(APP_TITLE, "Indiquez la boite cible.")
                return True
            return False

        def _importer(self, connected=True) -> Importer:
            mbx = self.mailbox.get().strip() or "boite@exemple.fr"
            auth = self.auth if (connected and self.auth and self.auth.result) else None
            if not self.importer or self.importer.mailbox != mbx or self.importer.auth is not auth:
                self.importer = Importer(auth, mbx, self.log)
            return self.importer

        # ---- actions
        def read_cloud(self):
            if self._need_login():
                return
            imp = self._importer()
            self.log(f"Lecture des parametres cloud de {imp.mailbox}...")
            self.run_bg(lambda: imp.read_cloud())

        def restore(self):
            if self._need_login():
                return
            f = filedialog.askopenfilename(title="Fichier de sauvegarde (exports\\sauvegarde_*.json)",
                                           initialdir=str(EXPORT_DIR), filetypes=[("JSON", "*.json")])
            if not f:
                return
            dry = self.dry_run.get()
            imp = self._importer()
            if not dry and not messagebox.askyesno(
                    APP_TITLE, f"Renvoyer les signatures de la sauvegarde\n{Path(f).name}\n"
                               f"dans la boite {imp.mailbox} ?\n\n"
                               "Les signatures de meme nom seront remplacees."):
                return
            self.run_bg(lambda: imp.restore(Path(f), dry))

        def send(self):
            if not self.sigs:
                messagebox.showwarning(APP_TITLE, "Analysez d'abord un dossier.")
                return
            sigs = self.selected()
            new_def, reply_def = self._default_value(self.def_new), self._default_value(self.def_reply)
            if not sigs and new_def is None and reply_def is None:
                messagebox.showwarning(APP_TITLE, "Cochez au moins une signature a importer.")
                return
            dry = self.dry_run.get()
            if not dry and self._need_login():
                return
            imp = self._importer(connected=not dry or bool(self.auth and self.auth.result))
            if not dry and not messagebox.askyesno(
                    APP_TITLE, f"Envoyer {len(sigs)} signature(s) dans la boite\n{imp.mailbox} ?\n\n"
                               + "\n".join(f"  - {x.name}" for x in sigs[:15])
                               + ("\n  ..." if len(sigs) > 15 else "") + "\n\n"
                               "Une copie de la boite est enregistree avant envoi (exports\\sauvegarde_*.json).\n"
                               "Une signature du meme nom sera remplacee."):
                return

            def job():
                ok, failed = imp.send(sigs, dry, new_def, reply_def)
                if dry:
                    self.ui(lambda: messagebox.showinfo(
                        APP_TITLE, "MODE SIMULATION : rien n'a ete envoye.\n\n"
                                   "Decochez 'Simulation' pour envoyer reellement dans la boite."))
                    return
                if ok:
                    self.log("Verification : relecture du cloud...")
                    good, bad = imp.verify(sigs)
                    summary = (f"{len(good)}/{len(sigs)} signature(s) confirmee(s) dans la boite."
                               + ("\n\nNon visibles : " + ", ".join(bad) +
                                  "\nEnvoyez le journal et le dernier fichier cloud_*.json." if bad else
                                  "\n\nOutlook Web : rechargez la page. Outlook : redemarrez-le."))
                    self.ui(lambda: messagebox.showinfo(APP_TITLE, summary))
            self.run_bg(job)

    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    App(root)
    root.mainloop()


# --------------------------------------------------------------------------- ligne de commande
def attach_console():
    """L'exe est compile en mode fenetre : on se rattache a la console qui l'a lance."""
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w")
    if sys.stderr is None:
        sys.stderr = sys.stdout
    if sys.platform != "win32" or sys.stdout.isatty():
        return
    try:
        import ctypes
        if ctypes.windll.kernel32.AttachConsole(-1):
            sys.stdout = open("CONOUT$", "w", encoding="utf-8", buffering=1)
            sys.stderr = sys.stdout
    except Exception:
        pass


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="SignaturesO365",
        description="Importe des signatures Outlook locales dans la boite Exchange Online. "
                    "Sans --apply, rien n'est envoye (simulation).")
    p.add_argument("--folder", help="dossier de signatures (defaut : dossier d'Outlook classique)")
    p.add_argument("--mailbox", help="boite cible (defaut : compte connecte)")
    p.add_argument("--tenant", help="tenant ou domaine (defaut : config ou 'organizations')")
    sel = p.add_mutually_exclusive_group()
    sel.add_argument("--all", action="store_true", help="importer toutes les signatures du dossier")
    sel.add_argument("--name", action="append", default=[], metavar="NOM",
                     help="signature a importer (repetable)")
    p.add_argument("--default-new", metavar="NOM", help="signature par defaut des nouveaux messages")
    p.add_argument("--default-reply", metavar="NOM", help="signature par defaut des reponses/transferts")
    p.add_argument("--outlook-defaults", action="store_true",
                   help="reprendre les signatures par defaut d'Outlook classique")
    p.add_argument("--apply", action="store_true", help="envoyer reellement (sinon simulation)")
    p.add_argument("--list", action="store_true", help="analyser le dossier et lister les signatures")
    p.add_argument("--read-cloud", action="store_true", help="lire et exporter les signatures de la boite")
    p.add_argument("--restore", metavar="FICHIER", help="renvoyer les signatures d'une sauvegarde JSON")
    p.add_argument("--device-code", action="store_true", help="connexion par code d'appareil")
    p.add_argument("--no-interactive", action="store_true",
                   help="echouer plutot qu'ouvrir une fenetre de connexion (cache uniquement)")
    p.add_argument("--version", action="version", version=f"SignaturesO365 {APP_VERSION}")
    return p


def run_cli(argv) -> int:
    attach_console()
    args = build_parser().parse_args(argv)
    cfg = load_config()

    def log(msg):
        line = f"[{time.strftime('%H:%M:%S')}] {msg}"
        write_journal(line)
        try:
            print(line, flush=True)
        except Exception:
            pass

    folder = Path(args.folder) if args.folder else local_signature_dir()
    sigs = []
    if folder:
        if not folder.is_dir():
            log(f"ERREUR : dossier introuvable : {folder}")
            return 2
        sigs = scan_folder(folder)
    if args.list or not (args.read_cloud or args.restore):
        if not folder:
            log("ERREUR : aucun dossier (--folder) et dossier d'Outlook classique introuvable.")
            return 2
        log(f"{len(sigs)} signature(s) dans {folder}")
        for s in sigs:
            size, warns = analyse(s)
            log(f"  - {s.name}  ({size // 1024} Ko, {s.image_count} image(s))"
                + (f"  /!\\ {' ; '.join(warns)}" if warns else ""))
        if args.list:
            return 0

    if args.all:
        chosen = sigs
    else:
        by_name = {s.name.lower(): s for s in sigs}
        unknown = [n for n in args.name if n.lower() not in by_name]
        if unknown:
            log(f"ERREUR : signature(s) absente(s) du dossier : {', '.join(unknown)}")
            return 2
        chosen = [by_name[n.lower()] for n in args.name]

    new_def, reply_def = args.default_new, args.default_reply
    if args.outlook_defaults:
        d = local_default_signatures()
        new_def = new_def if new_def is not None else d.get("new")
        reply_def = reply_def if reply_def is not None else d.get("reply")
        log(f"Defauts Outlook classique : nouveau={d.get('new', '-')}, reponse={d.get('reply', '-')}")

    if not (chosen or args.read_cloud or args.restore or new_def or reply_def):
        log("Rien a faire : indiquez --all, --name, --default-new/--default-reply, --read-cloud ou --restore.")
        return 2

    if roaming_disabled_locally():
        log("/!\\ Signatures itinerantes desactivees dans Outlook classique sur ce poste.")

    need_auth = args.apply or args.read_cloud or bool(args.restore) or bool(new_def or reply_def)
    auth = None
    tenant = args.tenant or cfg.get("tenant") or "organizations"
    if need_auth:
        auth = Auth(DEFAULT_CLIENT_ID, tenant, use_broker=not args.device_code)
        hint = args.mailbox or cfg.get("mailbox")
        if not auth.silent(hint):
            if args.no_interactive:
                log("ERREUR : aucun jeton en cache et --no-interactive demande.")
                return 3
            if args.device_code:
                auth.device(lambda f: log(f"Ouvrez {f['verification_uri']} et saisissez le code {f['user_code']}"))
            else:
                auth.interactive(login_hint=hint)
        log(f"Connecte en tant que {auth.upn}")

    mailbox = args.mailbox or (auth.upn if auth else "") or cfg.get("mailbox") or "boite@exemple.fr"
    imp = Importer(auth, mailbox, log)
    dry = not args.apply
    if dry:
        log("MODE SIMULATION : ajoutez --apply pour envoyer reellement.")

    rc = 0
    if args.read_cloud:
        imp.read_cloud()
    if args.restore:
        if not imp.restore(Path(args.restore), dry):
            rc = 1
    if chosen or new_def is not None or reply_def is not None:
        ok, failed = imp.send(chosen, dry, new_def, reply_def)
        if failed:
            rc = 1
        if not dry and ok and chosen:
            _, bad = imp.verify(chosen)
            if bad:
                rc = 1
    return rc


def main():
    if len(sys.argv) > 1:
        try:
            code = run_cli(sys.argv[1:])
        except SystemExit:
            raise
        except Exception as e:
            print("ERREUR : " + explain_error(str(e)), file=sys.stderr)
            write_journal(f"ERREUR : {e}")
            code = 1
        sys.exit(code)
    run_gui()


if __name__ == "__main__":
    main()
