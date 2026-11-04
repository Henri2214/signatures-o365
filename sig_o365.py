# -*- coding: utf-8 -*-
"""
Signatures O365 - envoie un dossier de signatures Outlook (.htm/.txt + dossier d'images)
directement dans la boite Exchange Online (signatures "cloud" / roaming), sans passer
par le dossier local %APPDATA%\\Microsoft\\Signatures.

ATTENTION : l'API utilisee (substrate.office.com / outlookcloudsettings) est celle
qu'utilisent Outlook et OWA en interne. Elle n'est PAS documentee par Microsoft et peut
changer sans preavis. Tester d'abord sur une boite de test, en mode simulation.
"""

import base64
import html as htmlmod
import json
import mimetypes
import queue
import re
import sys
import threading
import time
import urllib.parse
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import msal
import requests
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

APP_TITLE = "Signatures O365 - import dans la boite mail"
APP_VERSION = "1.8"
BOX_ON, BOX_OFF = "\u2611", "\u2610"

# Client public "Microsoft Office" (celui qu'utilise Outlook pour ce service).
# Remplacable dans l'onglet avance si votre tenant le bloque.
DEFAULT_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
SCOPES = ["https://outlook.office.com/.default"]
SUBSTRATE = "https://substrate.office.com/ows/beta/outlookcloudsettings/settings"
USER_AGENT = "Microsoft Office/16.0 (Windows NT 10.0; Microsoft Outlook 16.0; Pro)"
IMG_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
FILES_SUFFIXES = ("_files", "_fichiers", "-fichiers", ".files")


# --------------------------------------------------------------------------- chemins
def base_dir() -> Path:
    """Dossier de l'exe (PyInstaller) ou du script : l'outil reste portable."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


DATA_DIR = base_dir() / "data"
EXPORT_DIR = base_dir() / "exports"


# --------------------------------------------------------------------------- signatures
@dataclass
class Signature:
    name: str
    htm: Path
    txt: Optional[Path]
    files_dir: Optional[Path]

    @property
    def image_count(self) -> int:
        if not self.files_dir:
            return 0
        return sum(1 for f in self.files_dir.iterdir() if f.suffix.lower() in IMG_EXT)


def scan_folder(folder: Path) -> list:
    entries = {p.name.lower(): p for p in folder.iterdir()}
    sigs = []
    for p in sorted(folder.iterdir(), key=lambda x: x.name.lower()):
        if not (p.is_file() and p.suffix.lower() in (".htm", ".html")):
            continue
        stem = p.stem
        txt = entries.get(f"{stem}.txt".lower())
        files_dir = None
        for suf in FILES_SUFFIXES:
            d = entries.get(f"{stem}{suf}".lower())
            if d and d.is_dir():
                files_dir = d
                break
        sigs.append(Signature(stem, p, txt if txt and txt.is_file() else None, files_dir))
    return sigs


def read_text_file(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8", errors="replace")
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16", errors="replace")
    m = re.search(rb'charset\s*=\s*["\']?([\w\-]+)', raw[:4000], re.I)
    candidates = [m.group(1).decode("ascii", "ignore")] if m else []
    candidates += ["utf-8", "cp1252"]
    for enc in candidates:
        try:
            return raw.decode(enc)
        except (LookupError, UnicodeDecodeError):
            continue
    return raw.decode("cp1252", errors="replace")


def _is_local_ref(value: str) -> bool:
    v = value.strip().lower()
    return not v.startswith(("http:", "https:", "data:", "cid:", "mailto:", "tel:", "#", "javascript:"))


def _resolve(htm: Path, ref: str) -> Optional[Path]:
    ref = urllib.parse.unquote(ref.strip()).replace("\\", "/")
    p = (htm.parent / ref)
    try:
        p = p.resolve()
    except OSError:
        return None
    return p if p.is_file() else None


def to_data_uri(path: Path) -> str:
    mime = mimetypes.guess_type(path.name)[0] or "image/png"
    return f"data:{mime};base64,{base64.b64encode(path.read_bytes()).decode('ascii')}"


LINK_RE = re.compile(r"<link\b[^>]*>", re.I)
HREF_RE = re.compile(r'\bhref\s*=\s*(["\'])([^"\']+)\1', re.I)
SRC_RE = re.compile(r'(\b(?:src|o:href)\s*=\s*)(["\'])([^"\']+)\2', re.I)


def prepare_html(sig: Signature, log=None) -> str:
    """Lit le .htm, integre les images en base64 et nettoie les references locales."""
    html = read_text_file(sig.htm)

    # <link rel=File-List href="Nom_files/filelist.xml"> etc. : inutiles dans le cloud
    def drop_local_link(m):
        h = HREF_RE.search(m.group(0))
        return "" if h and _is_local_ref(h.group(2)) else m.group(0)
    html = LINK_RE.sub(drop_local_link, html)

    missing = []

    def inline(m):
        prefix, quote, value = m.groups()
        if not _is_local_ref(value):
            return m.group(0)
        p = _resolve(sig.htm, value)
        if p and p.suffix.lower() in IMG_EXT:
            return f"{prefix}{quote}{to_data_uri(p)}{quote}"
        missing.append(value)
        return m.group(0)
    html = SRC_RE.sub(inline, html)

    # le contenu est envoye en Unicode : on aligne la declaration de charset
    html = re.sub(r'(charset\s*=\s*["\']?)[\w\-]+', r"\1utf-8", html, count=1, flags=re.I)

    if missing and log:
        log(f"  /!\\ {sig.name} : {len(missing)} reference(s) locale(s) introuvable(s) : "
            + ", ".join(sorted(set(missing))[:5]))
    return html


def html_to_text(src: str) -> str:
    """Conversion HTML -> texte par expressions regulieres (aucun module externe)."""
    src = re.sub(r"(?is)<!--.*?-->", "", src)
    src = re.sub(r"(?is)<(head|style|script)\b.*?</\1\s*>", "", src)
    src = re.sub(r"(?i)<br\s*/?>|</(p|div|tr|li|h\d)\s*>", "\n", src)
    src = re.sub(r"(?s)<[^>]+>", "", src)
    return htmlmod.unescape(src)


def prepare_text(sig: Signature) -> str:
    if sig.txt:
        return read_text_file(sig.txt).replace("\r\n", "\n").replace("\n", "\r\n")
    text = html_to_text(read_text_file(sig.htm))
    lines = [re.sub(r"[ \t\xa0]+", " ", l).strip() for l in text.splitlines()]
    out, blank = [], False
    for l in lines:
        if not l:
            if not blank and out:
                out.append("")
            blank = True
        else:
            out.append(l)
            blank = False
    return "\r\n".join(out).strip() + "\r\n"


def dotnet_ticks() -> int:
    return int((time.time() + 62135596800) * 10_000_000)


COPYABLE_FIELDS = ("type", "source")


def _ci(d: dict, key: str):
    """Lecture d'un champ sans tenir compte de la casse (Timestamp / timestamp...)."""
    for k, v in d.items():
        if k.lower() == key.lower():
            return v
    return None


def learn_format(read_result) -> dict:
    """Analyse les signatures creees par Outlook pour copier leur format exact."""
    items = []

    def walk(o):
        if isinstance(o, dict):
            if str(_ci(o, "parentSetting") or "").lower() == "roaming_signature_list":
                items.append(o)
            for v in o.values():
                walk(v)
        elif isinstance(o, list):
            for v in o:
                walk(v)
    walk(read_result)
    keys = []
    for it in items:
        k = str(_ci(it, "secondaryKey") or "").lower()
        if k and k not in keys:
            keys.append(k)
    sample = {}
    if items:
        for f in COPYABLE_FIELDS:
            v = _ci(items[0], f)
            if v is not None and not isinstance(v, (dict, list)):
                sample[f] = v
    return {"count": len(items), "secondary_keys": keys, "sample": sample,
            "fields": sorted(items[0].keys()) if items else []}


def build_items(name: str, scope: str, html: str, text: str, fmt: Optional[dict] = None) -> list:
    base = {
        "name": name, "itemClass": "", "id": "", "scope": scope,
        "parentSetting": "roaming_signature_list", "type": "Blob",
        "timestamp": dotnet_ticks(), "metadata": "", "isFirstSync": "false",
        "source": "UserOverride",
    }
    keys = ["htm", "txt"]
    if fmt and fmt.get("count"):
        base.update(fmt.get("sample", {}))
        learned = fmt.get("secondary_keys") or []
        if learned:
            keys = [k for k in learned if k in ("htm", "html", "txt", "text")] or keys
    out = []
    for k in keys:
        out.append(dict(base, secondaryKey=k, value=text if k in ("txt", "text") else html))
    return out


# --------------------------------------------------------------------------- auth
def decode_jwt(token: str) -> dict:
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload))
    except Exception:
        return {}


AADSTS_HELP = {
    "50011": "Adresse de retour refusee. Dans l'app Entra > Authentification, ajoutez la plateforme "
             "'Applications mobiles et de bureau' avec l'URI http://localhost (sans port).",
    "53003": "Connexion bloquee par une strategie d'acces conditionnel du tenant (souvent le blocage "
             "du flux 'code d'appareil'). Essayez l'autre bouton de connexion, ou excluez "
             "temporairement le compte de la strategie dans Entra ID.",
    "65001": "Consentement requis : un administrateur doit autoriser l'application dans le tenant.",
    "90094": "Consentement administrateur requis pour cette application dans le tenant.",
    "900941": "Consentement administrateur requis pour cette application dans le tenant.",
    "700016": "Application introuvable dans ce tenant : verifiez le Client ID et le Tenant.",
    "50126": "Identifiant ou mot de passe incorrect.",
    "50020": "Ce compte n'appartient pas au tenant indique (compte personnel ou invite ?).",
    "50076": "Authentification multifacteur exigee : refaites la connexion et validez la MFA.",
    "50079": "Inscription a la MFA exigee : terminez-la dans le navigateur puis recommencez.",
    "50105": "Le compte n'est pas autorise a utiliser cette application dans ce tenant.",
    "7000218": "Dans l'app Entra : Authentification > 'Autoriser les flux de clients publics' = Oui.",
    "500113": "Aucune adresse de retour : ajoutez la plateforme 'Applications mobiles et de bureau' "
              "avec l'URI http://localhost dans l'app Entra.",
    "650057": "Ressource non autorisee : ajoutez l'API 'Office 365 Exchange Online' dans les "
              "autorisations de l'app Entra puis accordez le consentement administrateur.",
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
        return ("Impossible de joindre login.microsoftonline.com (proxy, pare-feu, filtrage SSL ?).\n\n"
                + msg[:600])
    return msg[:800]


class Auth:
    def __init__(self, client_id: str, tenant: str, use_broker: bool = True):
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        self.cache_file = DATA_DIR / "token_cache.bin"
        self.cache = msal.SerializableTokenCache()
        if self.cache_file.exists():
            self.cache.deserialize(self.cache_file.read_text())
        authority = f"https://login.microsoftonline.com/{tenant or 'organizations'}"
        self.broker = False
        self.app = None
        if use_broker and sys.platform == "win32":
            # Gestionnaire de comptes Windows (WAM) : meme mecanisme qu'Outlook/Office
            try:
                import pymsalruntime  # noqa: F401
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
            self.cache_file.write_text(self.cache.serialize())

    def _check(self, r):
        if not r or "access_token" not in r:
            r = r or {}
            raise RuntimeError(r.get("error_description") or r.get("error") or "Echec d'authentification")
        self.result = r
        self._save()
        return r

    def silent(self):
        for acc in self.app.get_accounts():
            r = self.app.acquire_token_silent(SCOPES, account=acc)
            if r and "access_token" in r:
                return self._check(r)
        return None

    def device(self, show_code):
        flow = self.app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(flow.get("error_description", "Device code indisponible"))
        show_code(flow)
        return self._check(self.app.acquire_token_by_device_flow(flow))

    def interactive(self, parent_hwnd=None):
        kw = {"prompt": "select_account"}
        if self.broker:
            kw["parent_window_handle"] = parent_hwnd or self.app.CONSOLE_WINDOW_HANDLE
        return self._check(self.app.acquire_token_interactive(SCOPES, **kw))

    def oauth_browser(self):
        """OAuth2 standard : code d'autorisation + PKCE, retour sur http://localhost:<port>."""
        return self._check(self.app.acquire_token_interactive(
            SCOPES, prompt="select_account", timeout=300,
            success_template="<html><body style='font-family:sans-serif'><h2>Connexion reussie</h2>"
                             "<p>Vous pouvez fermer cet onglet et revenir a SignaturesO365.</p></body></html>",
            error_template="<html><body style='font-family:sans-serif'><h2>Echec de connexion</h2>"
                           "<p>$error : $error_description</p></body></html>"))

    def logout(self):
        for acc in self.app.get_accounts():
            self.app.remove_account(acc)
        self.result = None
        self._save()

    @property
    def token(self) -> str:
        r = self.silent() or self.result
        if not r:
            raise RuntimeError("Non connecte")
        return r["access_token"]

    @property
    def claims(self) -> dict:
        return decode_jwt(self.result["access_token"]) if self.result else {}


# --------------------------------------------------------------------------- API cloud
class CloudSettings:
    def __init__(self, auth: Auth):
        self.auth = auth

    def _headers(self, mailbox: str) -> dict:
        c = self.auth.claims
        me = (c.get("upn") or c.get("unique_name") or "").lower()
        if mailbox.lower() == me and c.get("puid") and c.get("tid"):
            anchor = f"PUID:{c['puid']}@{c['tid']}"
        else:
            anchor = f"SMTP:{mailbox}"
        return {
            "Authorization": f"Bearer {self.auth.token}",
            "Content-Type": "application/json", "Accept": "application/json",
            "User-Agent": USER_AGENT, "x-anchormailbox": anchor, "x-islargesetting": "true",
        }

    def read(self, mailbox: str) -> dict:
        out = {}
        for path in ("account", "global?force=true"):
            r = requests.get(f"{SUBSTRATE}/{path}", headers=self._headers(mailbox), timeout=60)
            try:
                body = r.json()
            except ValueError:
                body = r.text
            out[path] = {"status": r.status_code, "body": body}
        return out

    def upload(self, mailbox: str, items: list) -> requests.Response:
        data = json.dumps(items, ensure_ascii=False).encode("utf-8")
        return requests.patch(f"{SUBSTRATE}/account?fromows=true",
                              headers=self._headers(mailbox), data=data, timeout=120)


def extract_signature_names(read_result: dict) -> list:
    names = set()

    def walk(o):
        if isinstance(o, dict):
            if str(_ci(o, "parentSetting") or "").lower() == "roaming_signature_list":
                names.add(_ci(o, "name"))
            for v in o.values():
                walk(v)
        elif isinstance(o, list):
            for v in o:
                walk(v)
    walk(read_result)
    return sorted(n for n in names if n)


# --------------------------------------------------------------------------- interface
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title(f"{APP_TITLE}  v{APP_VERSION}")
        root.geometry("980x700")
        root.minsize(820, 560)
        self.q = queue.Queue()
        self.sigs = []
        self.checked = set()
        self.auth = None

        self.folder = tk.StringVar()
        self.mailbox = tk.StringVar()
        self.tenant = tk.StringVar(value="organizations")
        self.cloud_format = None
        self.dry_run = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value="Non connecte")
        self._load_config()
        self._build()
        root.after(100, self._poll)
        self.log(f"Exports et journaux : {EXPORT_DIR}")
        self.log("Mode simulation actif par defaut : rien n'est envoye tant qu'il est coche.")

    # ---- configuration (data\config.json, a cote de l'exe)
    CONFIG_FILE = DATA_DIR / "config.json"

    def _load_config(self):
        try:
            cfg = json.loads(self.CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return
        for key, var in (("tenant", self.tenant), ("mailbox", self.mailbox), ("folder", self.folder)):
            if cfg.get(key):
                var.set(cfg[key])

    def _save_config(self):
        try:
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            self.CONFIG_FILE.write_text(json.dumps({
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
        ttk.Button(f1, text="Analyser", command=self.analyse).pack(side="left", **pad)

        f2 = ttk.Frame(self.root)
        f2.pack(fill="both", expand=True, **pad)
        cols = ("choix", "images", "txt", "dossier", "taille")
        self.tree = ttk.Treeview(f2, columns=cols, selectmode="browse", height=8)
        self.tree.heading("#0", text="Signature")
        self.tree["displaycolumns"] = cols
        for c, t, w in (("choix", "Importer", 70), ("images", "Images", 70), ("txt", "TXT", 50),
                        ("dossier", "Dossier images", 220), ("taille", "Taille HTML finale", 130)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="center")
        self.tree.column("#0", width=260)
        sb = ttk.Scrollbar(f2, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tree.bind("<Button-1>", self._on_click)
        self.tree.bind("<space>", lambda e: self._toggle(self.tree.focus()))

        fsel = ttk.Frame(self.root)
        fsel.pack(fill="x", padx=6)
        ttk.Button(fsel, text="Tout cocher", command=lambda: self._set_all(True)).pack(side="left", padx=4)
        ttk.Button(fsel, text="Tout decocher", command=lambda: self._set_all(False)).pack(side="left", padx=4)
        ttk.Button(fsel, text="Inverser", command=self._invert).pack(side="left", padx=4)
        self.count_lbl = ttk.Label(fsel, text="0 signature cochee", foreground="gray")
        self.count_lbl.pack(side="left", padx=12)

        f3 = ttk.LabelFrame(self.root, text="2. Compte Microsoft 365")
        f3.pack(fill="x", **pad)
        ttk.Label(f3, text="Boite cible :").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(f3, textvariable=self.mailbox, width=40).grid(row=0, column=1, sticky="we", **pad)
        ttk.Button(f3, text="Se connecter (Office 365)", command=lambda: self.login("windows")).grid(row=0, column=2, **pad)
        ttk.Button(f3, text="Deconnexion", command=self.logout).grid(row=0, column=3, **pad)
        adv = ttk.Frame(f3)
        adv.grid(row=1, column=0, columnspan=4, sticky="we")
        ttk.Label(adv, text="Tenant (domaine) :").pack(side="left", padx=6)
        ttk.Entry(adv, textvariable=self.tenant, width=30).pack(side="left")
        ttk.Label(adv, text="   Secours :", foreground="gray").pack(side="left", padx=6)
        ttk.Button(adv, text="Code d'appareil", command=lambda: self.login("device")).pack(side="left", padx=4)
        ttk.Label(f3, textvariable=self.status, foreground="#0a5").grid(row=2, column=0, columnspan=4, sticky="w", **pad)
        f3.columnconfigure(1, weight=1)

        f4 = ttk.LabelFrame(self.root, text="3. Actions")
        f4.pack(fill="x", **pad)
        ttk.Button(f4, text="Lire les signatures du cloud", command=self.read_cloud).pack(side="left", **pad)
        ttk.Checkbutton(f4, text="Simulation (genere le JSON sans rien envoyer)",
                        variable=self.dry_run, command=self._update_send_label).pack(side="left", **pad)
        self.send_btn = ttk.Button(f4, command=self.send)
        self.send_btn.pack(side="right", **pad)
        self._update_send_label()

        self.logbox = ScrolledText(self.root, height=12, font=("Consolas", 9))
        self.logbox.pack(fill="both", expand=True, **pad)

    def _update_send_label(self):
        self.send_btn.config(text="Simuler (rien n'est envoye)" if self.dry_run.get()
                             else "ENVOYER dans la boite mail")

    # ---- utilitaires
    def log(self, msg):
        self.q.put(("log", msg))

    def _poll(self):
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == "log":
                    self.logbox.insert("end", f"[{time.strftime('%H:%M:%S')}] {val}\n")
                    self.logbox.see("end")
                elif kind == "call":
                    val()
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def ui(self, fn):
        self.q.put(("call", fn))

    def run_bg(self, fn):
        def wrap():
            try:
                fn()
            except Exception as e:
                text = explain_error(str(e))
                self.log(f"ERREUR : {e}")
                self.ui(lambda: messagebox.showerror(APP_TITLE, text))
        threading.Thread(target=wrap, daemon=True).start()

    # ---- dossier
    def browse(self):
        d = filedialog.askdirectory(title="Dossier contenant les signatures")
        if d:
            self.folder.set(d)
            self.analyse()

    def analyse(self):
        folder = Path(self.folder.get().strip('" '))
        if not folder.is_dir():
            messagebox.showerror(APP_TITLE, "Dossier introuvable.")
            return
        self.sigs = scan_folder(folder)
        self.tree.delete(*self.tree.get_children())
        self.checked.clear()
        for i, s in enumerate(self.sigs):
            try:
                size = len(prepare_html(s).encode("utf-8"))
                size_txt = f"{size / 1024:.0f} Ko"
            except Exception as e:
                size_txt = f"erreur : {e}"
            self.tree.insert("", "end", iid=str(i), text=s.name, values=(
                BOX_OFF, s.image_count, "oui" if s.txt else "auto",
                s.files_dir.name if s.files_dir else "-", size_txt))
        self._refresh_count()
        self.log(f"{len(self.sigs)} signature(s) trouvee(s) dans {folder} - cochez celles a importer.")
        if not self.sigs:
            self.log("Aucun fichier .htm trouve : verifier le dossier.")

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

        client_id = DEFAULT_CLIENT_ID
        tenant = self.tenant.get().strip() or "organizations"
        self._save_config()
        labels = {"windows": "Office 365 (fenetre Windows)", "device": "code d'appareil"}

        def job():
            self.auth = Auth(client_id, tenant, use_broker=(mode == "windows"))
            self.log(f"Connexion {labels[mode]} - tenant {tenant}...")
            if mode == "windows" and sys.platform == "win32" and not self.auth.broker:
                raise RuntimeError("Composant de connexion Windows (pymsalruntime) absent de l'exe. "
                                   "Recompilez avec le nouveau build.bat (il installe Python 3.12 et "
                                   "verifie ce composant). En attendant, utilisez 'Code d'appareil'.")
            r = self.auth.silent()
            if not r:
                r = self.auth.device(self._show_code) if mode == "device" else self.auth.interactive(hwnd)
            c = self.auth.claims
            upn = c.get("upn") or c.get("unique_name") or ""
            self.ui(lambda: (self.status.set(f"Connecte : {upn}  (tenant {c.get('tid', '?')})"),
                             self.mailbox.get() or self.mailbox.set(upn)))
            self.log(f"Connecte en tant que {upn}")
            scp = c.get("scp", "")
            self.log(f"Autorisations du jeton : {scp or '(aucune)'}")
            self.ui(self._save_config)
        self.run_bg(job)

    def logout(self):
        if self.auth:
            self.auth.logout()
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

    # ---- actions
    def read_cloud(self):
        if self._need_login():
            return
        mbx = self.mailbox.get().strip()

        def job():
            self.log(f"Lecture des parametres cloud de {mbx}...")
            res = CloudSettings(self.auth).read(mbx)
            EXPORT_DIR.mkdir(parents=True, exist_ok=True)
            out = EXPORT_DIR / f"cloud_{mbx}_{time.strftime('%Y%m%d_%H%M%S')}.json"
            out.write_text(json.dumps(res, ensure_ascii=False, indent=2), encoding="utf-8")
            for k, v in res.items():
                self.log(f"  GET {k} -> HTTP {v['status']}")
            if any(v["status"] in (401, 403) for v in res.values()):
                msg = ("Acces refuse (HTTP 401/403) par le service des signatures cloud.\n"
                       "Cliquez sur Deconnexion puis reconnectez-vous avec le bouton Office 365.\n"
                       "Si le probleme persiste, envoyez le fichier JSON du dossier exports.")
                self.log(msg.replace("\n", " "))
                self.ui(lambda: messagebox.showwarning(APP_TITLE, msg))
            names = extract_signature_names(res)
            self.log(f"Signatures cloud trouvees : {', '.join(names) if names else '(aucune visible)'}")
            fmt = learn_format(res)
            if fmt["count"]:
                self.cloud_format = fmt
                self.log(f"Format Outlook detecte : cles {fmt['secondary_keys']}, champs {fmt['sample']}")
                self.log(f"  (champs presents : {', '.join(fmt['fields'])})")
            self.log(f"Reponse brute enregistree : {out}")
        self.run_bg(job)

    def send(self):
        if not self.sigs:
            messagebox.showwarning(APP_TITLE, "Analysez d'abord un dossier.")
            return
        sigs = self.selected()
        if not sigs:
            messagebox.showwarning(APP_TITLE, "Cochez au moins une signature a importer.")
            return
        dry = self.dry_run.get()
        if not dry and self._need_login():
            return
        mbx = self.mailbox.get().strip() or "boite@exemple.fr"
        if not dry and not messagebox.askyesno(
                APP_TITLE, f"Envoyer {len(sigs)} signature(s) dans la boite\n{mbx} ?\n\n"
                           + "\n".join(f"  - {x.name}" for x in sigs[:15])
                           + ("\n  ..." if len(sigs) > 15 else "") + "\n\n"
                           "Une signature du meme nom sera remplacee."):
            return

        def job():
            EXPORT_DIR.mkdir(parents=True, exist_ok=True)
            api = None if dry else CloudSettings(self.auth)
            fmt = self.cloud_format
            if not dry and not fmt:
                self.log("Lecture prealable du cloud pour copier le format des signatures Outlook...")
                fmt = learn_format(api.read(mbx))
                if fmt["count"]:
                    self.cloud_format = fmt
                    self.log(f"Format Outlook detecte : cles {fmt['secondary_keys']}")
                else:
                    self.log("Aucune signature Outlook existante : format par defaut utilise.")
            ok = 0
            for s in sigs:
                try:
                    items = build_items(s.name, mbx, prepare_html(s, self.log), prepare_text(s), fmt)
                    if dry:
                        out = EXPORT_DIR / f"simulation_{s.name}.json"
                        out.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
                        self.log(f"[SIMULATION] {s.name} -> {out.name}")
                        ok += 1
                        continue
                    r = api.upload(mbx, items)
                    if r.ok:
                        self.log(f"Envoye  {s.name}  (HTTP {r.status_code}) {r.text[:200]}")
                        ok += 1
                    else:
                        self.log(f"ECHEC {s.name}  HTTP {r.status_code} : {r.text[:400]}")
                except Exception as e:
                    self.log(f"ECHEC {s.name} : {e}")
            self.log(f"Termine : {ok}/{len(sigs)} signature(s) traitee(s).")
            if dry:
                self.ui(lambda: messagebox.showinfo(APP_TITLE,
                    "MODE SIMULATION : rien n'a ete envoye.\n\n"
                    "Decochez 'Simulation' pour envoyer reellement dans la boite."))
                return
            if ok:
                self.log("Verification : relecture du cloud...")
                time.sleep(3)
                present = set(n.lower() for n in extract_signature_names(api.read(mbx)))
                good = [x.name for x in sigs if x.name.lower() in present]
                bad = [x.name for x in sigs if x.name.lower() not in present]
                if good:
                    self.log(f"  Visibles dans le cloud : {', '.join(good)}")
                if bad:
                    self.log(f"  NON visibles dans le cloud : {', '.join(bad)}")
                summary = (f"{len(good)}/{len(sigs)} signature(s) confirmee(s) dans la boite."
                           + ("\n\nNon visibles : " + ", ".join(bad) +
                              "\nEnvoyez le journal et le dernier fichier cloud_*.json." if bad else
                              "\n\nOutlook Web : rechargez la page. Outlook : redemarrez-le."))
                self.ui(lambda: messagebox.showinfo(APP_TITLE, summary))
        self.run_bg(job)


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
