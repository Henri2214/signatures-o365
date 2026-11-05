# -*- coding: utf-8 -*-
"""
Coeur de Signatures O365 : lecture d'un dossier de signatures Outlook, nettoyage du HTML
genere par Word, integration des images, conversion texte et construction des elements
envoyes au service des signatures cloud.

Ce module ne depend ni de Tkinter, ni de MSAL, ni de requests : il est testable seul.
"""

import base64
import html as htmlmod
import mimetypes
import re
import sys
import time
import urllib.parse
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

IMG_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
# Formats produits par Word mais non affiches par OWA / le nouvel Outlook
UNSUPPORTED_IMG_EXT = {".emf", ".wmf", ".svg", ".tif", ".tiff", ".ico"}
FILES_SUFFIXES = ("_files", "_fichiers", "-fichiers", ".files")
SIG_LIST = "roaming_signature_list"

# Seuils d'alerte (la limite reelle du service n'est pas documentee)
WARN_IMAGE_BYTES = 200 * 1024
WARN_HTML_BYTES = 1024 * 1024
MAX_NAME_LEN = 100
FORBIDDEN_NAME_CHARS = set('\\/:*?"<>|')


# --------------------------------------------------------------------------- dossier
@dataclass
class Signature:
    name: str
    htm: Path
    txt: Optional[Path]
    files_dir: Optional[Path]
    warnings: list = field(default_factory=list)

    @property
    def image_count(self) -> int:
        if not self.files_dir:
            return 0
        return sum(1 for f in self.files_dir.iterdir() if f.suffix.lower() in IMG_EXT)


def validate_name(name: str) -> list:
    out = []
    if not name.strip():
        out.append("nom vide")
    if len(name) > MAX_NAME_LEN:
        out.append(f"nom trop long ({len(name)} caracteres)")
    bad = sorted(set(name) & FORBIDDEN_NAME_CHARS)
    if bad:
        out.append("caracteres interdits dans le nom : " + " ".join(bad))
    if name != name.strip():
        out.append("espaces en debut ou fin de nom")
    return out


def scan_folder(folder: Path) -> list:
    entries = {p.name.lower(): p for p in folder.iterdir()}
    sigs, seen = [], {}
    for p in sorted(folder.iterdir(), key=lambda x: (x.name.lower(), x.suffix.lower() != ".htm")):
        if not (p.is_file() and p.suffix.lower() in (".htm", ".html")):
            continue
        stem = p.stem
        if stem.lower() in seen:
            seen[stem.lower()].warnings.append(f"fichier en double ignore : {p.name}")
            continue
        txt = entries.get(f"{stem}.txt".lower())
        files_dir = None
        for suf in FILES_SUFFIXES:
            d = entries.get(f"{stem}{suf}".lower())
            if d and d.is_dir():
                files_dir = d
                break
        sig = Signature(stem, p, txt if txt and txt.is_file() else None, files_dir)
        sig.warnings.extend(validate_name(stem))
        seen[stem.lower()] = sig
        sigs.append(sig)
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


# --------------------------------------------------------------------------- HTML
def _is_local_ref(value: str) -> bool:
    v = value.strip().lower()
    return not v.startswith(("http:", "https:", "data:", "cid:", "mailto:", "tel:", "#",
                             "javascript:", "sip:", "callto:", "skype:", "msteams:"))


def _resolve(htm: Path, ref: str) -> Optional[Path]:
    ref = urllib.parse.unquote(ref.strip()).replace("\\", "/")
    if ref.lower().startswith("file:"):
        ref = re.sub(r"(?i)^file:/*", "", ref)
        p = Path(ref)
    else:
        p = htm.parent / ref
    try:
        p = p.resolve()
    except OSError:
        return None
    return p if p.is_file() else None


def to_data_uri(path: Path) -> str:
    mime = mimetypes.guess_type(path.name)[0] or "image/png"
    return f"data:{mime};base64,{base64.b64encode(path.read_bytes()).decode('ascii')}"


VML_BLOCK_RE = re.compile(r"<!--\[if\s+gte\s+vml\s+\d*\]>.*?<!\[endif\]-->", re.I | re.S)
MSO_XML_RE = re.compile(r"<!--\[if\s+gte\s+mso\s+\d*\]>\s*<xml>.*?</xml>\s*<!\[endif\]-->", re.I | re.S)
XML_ISLAND_RE = re.compile(r"<xml\b[^>]*>.*?</xml>", re.I | re.S)
DOWNLEVEL_RE = re.compile(r"<!\[if\s+!vml\]>|<!\[if\s+!supportLists\]>|<!\[endif\]>", re.I)
OFFICE_TAG_RE = re.compile(r"</?(?:o|v|w|m):[a-z]+\b[^>]*>", re.I)
VSHAPES_ATTR_RE = re.compile(r"""\s+v:shapes\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)""", re.I)
META_OFFICE_RE = re.compile(r"<meta\b[^>]*\bname\s*=\s*[\"']?(?:ProgId|Generator|Originator)\b[^>]*>", re.I)
LINK_RE = re.compile(r"<link\b[^>]*>", re.I)
HREF_ATTR_RE = re.compile(r'\bhref\s*=\s*(["\'])([^"\']+)\1', re.I)
SRC_RE = re.compile(r'(\b(?:src|o:href)\s*=\s*)(["\'])([^"\']+)\2', re.I)
A_HREF_RE = re.compile(r'(<a\b[^>]*?\bhref\s*=\s*)(["\'])([^"\']+)\2', re.I)


def clean_word_html(html: str) -> str:
    """Retire les elements propres a Word / VML que ni OWA ni le nouvel Outlook n'affichent.

    - blocs <!--[if gte vml 1]> ... <![endif]--> (images VML en double),
    - ilots XML Office (<xml>...</xml>, <!--[if gte mso 9]><xml>...),
    - balises d'espaces de noms Office (<o:p>, <v:shape>, <w:...>),
    - <meta name=ProgId/Generator/Originator> et <link> vers des fichiers locaux.
    Le contenu "bas niveau" (<![if !vml]> ... <![endif]>) est conserve, sans ses marqueurs.
    """
    html = MSO_XML_RE.sub("", html)
    html = VML_BLOCK_RE.sub("", html)
    html = XML_ISLAND_RE.sub("", html)
    html = DOWNLEVEL_RE.sub("", html)
    html = OFFICE_TAG_RE.sub("", html)
    html = VSHAPES_ATTR_RE.sub("", html)
    html = META_OFFICE_RE.sub("", html)

    def drop_local_link(m):
        h = HREF_ATTR_RE.search(m.group(0))
        return "" if h and _is_local_ref(h.group(2)) else m.group(0)
    html = LINK_RE.sub(drop_local_link, html)
    return html


def prepare_html(sig: Signature, warnings: Optional[list] = None) -> str:
    """Lit le .htm, le nettoie, integre les images en base64. Les alertes vont dans `warnings`."""
    warn = warnings if warnings is not None else []
    html = clean_word_html(read_text_file(sig.htm))

    missing, unsupported, heavy = [], [], []

    def inline(m):
        prefix, quote, value = m.groups()
        if not _is_local_ref(value):
            return m.group(0)
        p = _resolve(sig.htm, value)
        if not p:
            missing.append(value)
            return m.group(0)
        ext = p.suffix.lower()
        if ext in UNSUPPORTED_IMG_EXT:
            unsupported.append(p.name)
            return m.group(0)
        if ext not in IMG_EXT:
            missing.append(value)
            return m.group(0)
        if p.stat().st_size > WARN_IMAGE_BYTES:
            heavy.append(f"{p.name} ({p.stat().st_size // 1024} Ko)")
        return f"{prefix}{quote}{to_data_uri(p)}{quote}"
    html = SRC_RE.sub(inline, html)

    local_links = []

    def check_link(m):
        value = m.group(3)
        if _is_local_ref(value):
            local_links.append(value)
        return m.group(0)
    A_HREF_RE.sub(check_link, html)

    # le contenu est envoye en Unicode : on aligne la declaration de charset
    html = re.sub(r'(charset\s*=\s*["\']?)[\w\-]+', r"\1utf-8", html, count=1, flags=re.I)

    def short(lst):
        return ", ".join(sorted(set(lst))[:5])
    if missing:
        warn.append(f"{len(set(missing))} image(s) introuvable(s) : {short(missing)}")
    if unsupported:
        warn.append(f"format d'image non affiche par OWA : {short(unsupported)}")
    if heavy:
        warn.append(f"image(s) lourde(s) : {short(heavy)}")
    if local_links:
        warn.append(f"lien(s) vers des fichiers locaux : {short(local_links)}")
    size = len(html.encode("utf-8"))
    if size > WARN_HTML_BYTES:
        warn.append(f"signature tres lourde ({size // 1024} Ko)")
    return html


def html_to_text(src: str) -> str:
    """Conversion HTML -> texte par expressions regulieres (aucun module externe)."""
    src = re.sub(r"(?is)<!--.*?-->", "", src)
    src = re.sub(r"(?is)<(head|style|script|xml)\b.*?</\1\s*>", "", src)
    src = re.sub(r"(?i)</t[dh]\s*>", " \t", src)
    src = re.sub(r"(?i)<br\s*/?>|</(p|div|tr|li|h\d|table)\s*>", "\n", src)
    src = re.sub(r"(?s)<[^>]+>", "", src)
    return htmlmod.unescape(src)


def prepare_text(sig: Signature) -> str:
    if sig.txt:
        return read_text_file(sig.txt).replace("\r\n", "\n").replace("\n", "\r\n")
    text = html_to_text(clean_word_html(read_text_file(sig.htm)))
    lines = [re.sub(r"[ \t\xa0]+", " ", ln).strip() for ln in text.splitlines()]
    out, blank = [], False
    for ln in lines:
        if not ln:
            if not blank and out:
                out.append("")
            blank = True
        else:
            out.append(ln)
            blank = False
    return "\r\n".join(out).strip() + "\r\n"


def analyse(sig: Signature) -> tuple:
    """Retourne (taille du HTML final en octets, liste des alertes)."""
    warns = list(sig.warnings)
    html = prepare_html(sig, warns)
    return len(html.encode("utf-8")), warns


# --------------------------------------------------------------------------- format cloud
def dotnet_ticks() -> int:
    return int((time.time() + 62135596800) * 10_000_000)


COPYABLE_FIELDS = ("type", "source")


def _ci(d: dict, key: str):
    """Lecture d'un champ sans tenir compte de la casse (Timestamp / timestamp...)."""
    for k, v in d.items():
        if k.lower() == key.lower():
            return v
    return None


def iter_settings(obj):
    """Parcourt une reponse du service et renvoie tous les objets portant un champ 'name'."""
    if isinstance(obj, dict):
        if _ci(obj, "name") is not None and (_ci(obj, "value") is not None or _ci(obj, "parentSetting")):
            yield obj
        for v in obj.values():
            yield from iter_settings(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from iter_settings(v)


def signature_items(read_result) -> list:
    return [o for o in iter_settings(read_result)
            if str(_ci(o, "parentSetting") or "").lower() == SIG_LIST]


def extract_signature_names(read_result) -> list:
    return sorted({_ci(o, "name") for o in signature_items(read_result) if _ci(o, "name")})


def learn_format(read_result) -> dict:
    """Analyse les signatures creees par Outlook pour copier leur format exact."""
    items = signature_items(read_result)
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
            "fields": sorted(items[0].keys()) if items else [],
            "defaults": find_default_settings(read_result)}


def build_items(name: str, scope: str, html: str, text: str, fmt: Optional[dict] = None) -> list:
    base = {
        "name": name, "itemClass": "", "id": "", "scope": scope,
        "parentSetting": SIG_LIST, "type": "Blob",
        "timestamp": dotnet_ticks(), "metadata": "", "isFirstSync": "false",
        "source": "UserOverride",
    }
    keys = ["htm", "txt"]
    if fmt and fmt.get("count"):
        base.update(fmt.get("sample", {}))
        learned = fmt.get("secondary_keys") or []
        if learned:
            keys = [k for k in learned if k in ("htm", "html", "txt", "text")] or keys
    return [dict(base, secondaryKey=k, value=text if k in ("txt", "text") else html) for k in keys]


# --------------------------------------------------------------------------- signatures par defaut
DEFAULT_PATTERNS = {
    "new": re.compile(r"(^|_)(new|compose)(_\w*)?_signature|signature_(new|compose)|^new_signature", re.I),
    "reply": re.compile(r"(^|_)(reply|replyforward|reply_forward)(_\w*)?_signature|signature_reply|^reply_signature", re.I),
}


def find_default_settings(read_result) -> dict:
    """Repere dans la reponse du service les parametres 'signature par defaut'.

    Le nom exact de ces parametres n'est pas documente : on cherche des parametres hors
    liste des signatures dont le nom evoque 'new ... signature' ou 'reply ... signature'.
    Retourne {"new": {...setting...}, "reply": {...}} pour ceux qui sont trouves.
    """
    found = {}
    for o in iter_settings(read_result):
        if str(_ci(o, "parentSetting") or "").lower() == SIG_LIST:
            continue
        name = str(_ci(o, "name") or "")
        for kind, rx in DEFAULT_PATTERNS.items():
            if kind not in found and rx.search(name):
                found[kind] = {k: v for k, v in o.items() if not isinstance(v, (dict, list))}
    return found


def build_default_items(defaults_fmt: dict, scope: str, new: Optional[str], reply: Optional[str]) -> list:
    """Construit les elements a envoyer pour definir les signatures par defaut."""
    out = []
    for kind, sig_name in (("new", new), ("reply", reply)):
        if sig_name is None or kind not in defaults_fmt:
            continue
        item = dict(defaults_fmt[kind])
        for k in list(item):
            kl = k.lower()
            if kl == "value":
                item[k] = sig_name
            elif kl == "timestamp":
                item[k] = dotnet_ticks()
            elif kl == "scope":
                item[k] = scope
        if not any(k.lower() == "value" for k in item):
            item["value"] = sig_name
        out.append(item)
    return out


def restore_items(backup) -> list:
    """Elements de signatures a renvoyer depuis une sauvegarde (reponse brute du service)."""
    items = []
    for o in signature_items(backup):
        item = {k: v for k, v in o.items() if not isinstance(v, (dict, list))}
        for k in item:
            if k.lower() == "timestamp":
                item[k] = dotnet_ticks()
        items.append(item)
    return items


# --------------------------------------------------------------------------- Outlook local (Windows)
OFFICE_KEY = r"Software\Microsoft\Office\16.0"
ACCOUNTS_GUID = "9375CFF0413111d3B88A00104B2A6676"


def _winreg():
    if sys.platform != "win32":
        return None
    import winreg
    return winreg


def local_signature_dir() -> Optional[Path]:
    """Dossier des signatures d'Outlook classique (nom localise lu dans le registre)."""
    import os
    appdata = os.environ.get("APPDATA")
    if not appdata:
        return None
    name = "Signatures"
    wr = _winreg()
    if wr:
        try:
            with wr.OpenKey(wr.HKEY_CURRENT_USER, OFFICE_KEY + r"\Common\General") as k:
                v, _ = wr.QueryValueEx(k, "Signatures")
                if v:
                    name = v
        except OSError:
            pass
    p = Path(appdata) / "Microsoft" / name
    return p if p.is_dir() else None


def _decode_reg(v) -> str:
    if isinstance(v, bytes):
        return v.decode("utf-16-le", errors="ignore").split("\x00")[0]
    return str(v or "")


def local_default_signatures() -> dict:
    """Signatures par defaut d'Outlook classique : {"new": nom, "reply": nom}.

    Lit d'abord la strategie (Common\\MailSettings), puis le premier compte de profil
    qui en definit une.
    """
    wr = _winreg()
    res = {}
    if not wr:
        return res
    try:
        with wr.OpenKey(wr.HKEY_CURRENT_USER, OFFICE_KEY + r"\Common\MailSettings") as k:
            for kind, val in (("new", "NewSignature"), ("reply", "ReplySignature")):
                try:
                    v, _ = wr.QueryValueEx(k, val)
                    if v:
                        res[kind] = _decode_reg(v)
                except OSError:
                    pass
    except OSError:
        pass
    try:
        profiles = wr.OpenKey(wr.HKEY_CURRENT_USER, OFFICE_KEY + r"\Outlook\Profiles")
    except OSError:
        return res
    with profiles:
        i = 0
        while True:
            try:
                prof = wr.EnumKey(profiles, i)
            except OSError:
                break
            i += 1
            try:
                acc_root = wr.OpenKey(profiles, prof + "\\" + ACCOUNTS_GUID)
            except OSError:
                continue
            with acc_root:
                j = 0
                while True:
                    try:
                        acc = wr.EnumKey(acc_root, j)
                    except OSError:
                        break
                    j += 1
                    try:
                        with wr.OpenKey(acc_root, acc) as k:
                            for kind, val in (("new", "New Signature"), ("reply", "Reply-Forward Signature")):
                                if kind in res:
                                    continue
                                try:
                                    v, _ = wr.QueryValueEx(k, val)
                                    if _decode_reg(v):
                                        res[kind] = _decode_reg(v)
                                except OSError:
                                    pass
                    except OSError:
                        continue
    return res


def roaming_disabled_locally() -> bool:
    """True si Outlook classique a les signatures itinerantes desactivees sur ce poste."""
    wr = _winreg()
    if not wr:
        return False
    for path in (OFFICE_KEY + r"\Outlook\Setup",
                 r"Software\Policies\Microsoft\Office\16.0\Outlook\Setup"):
        try:
            with wr.OpenKey(wr.HKEY_CURRENT_USER, path) as k:
                v, _ = wr.QueryValueEx(k, "DisableRoamingSignaturesTemporaryToggle")
                if int(v) == 1:
                    return True
        except (OSError, ValueError):
            continue
    return False
