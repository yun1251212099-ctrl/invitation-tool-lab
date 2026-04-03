#!/usr/bin/env python3
"""批量邀请函生成工具 — Streamlit Web 应用 v3 (Cloud)"""

import csv
import base64
import io
import os
import tempfile
import zipfile
from pathlib import Path

try:
    import cv2
    _HAS_CV2 = True
except ImportError:
    _HAS_CV2 = False
import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image, ImageDraw, ImageFont
try:
    from psd_tools import PSDImage
    _HAS_PSD = True
except ImportError:
    _HAS_PSD = False

import subprocess as _sp
import datetime as _dt
import re as _re

_EMOJI_RE = _re.compile(
    "["
    "\U0001F600-\U0001F64F"
    "\U0001F300-\U0001F5FF"
    "\U0001F680-\U0001F6FF"
    "\U0001F1E0-\U0001F1FF"
    "\U0001F900-\U0001F9FF"
    "\U0001FA00-\U0001FA6F"
    "\U0001FA70-\U0001FAFF"
    "\U0000FE0F"
    "\U0000200D"
    "\U00002702-\U000027B0"
    "\U000023E9-\U000023F3"
    "\U000023F8-\U000023FA"
    "\U0000231A-\U0000231B"
    "\U000025AA-\U000025AB"
    "\U000025B6"
    "\U000025C0"
    "\U000025FB-\U000025FE"
    "\U00002614-\U00002615"
    "\U00002648-\U00002653"
    "\U0000267F"
    "\U00002693"
    "\U000026A1"
    "\U000026AA-\U000026AB"
    "\U000026BD-\U000026BE"
    "\U000026C4-\U000026C5"
    "\U000026D4"
    "\U000026EA"
    "\U000026F2-\U000026F3"
    "\U000026F5"
    "\U000026FA"
    "\U000026FD"
    "\U00002934-\U00002935"
    "\U00002B05-\U00002B07"
    "\U00002B1B-\U00002B1C"
    "\U00002B50"
    "\U00002B55"
    "\U00003030"
    "\U0000303D"
    "\U00003297"
    "\U00003299"
    "]+",
    flags=_re.UNICODE,
)

def _strip_emoji(text: str) -> str:
    return _EMOJI_RE.sub("", text).strip()


def _is_emoji(ch: str) -> bool:
    return bool(_EMOJI_RE.fullmatch(ch))


_EMOJI_FONT_PATHS = [
    "/System/Library/Fonts/Apple Color Emoji.ttc",
    "/usr/share/fonts/truetype/noto/NotoColorEmoji.ttf",
    "/usr/share/fonts/noto-cjk/NotoColorEmoji.ttf",
    "C:\\Windows\\Fonts\\seguiemj.ttf",
]
_APPLE_EMOJI_VALID_SIZES = [20, 26, 32, 40, 48, 52, 64, 96, 160]
_emoji_font_cache: dict = {}


def _get_emoji_font(size: int):
    if size in _emoji_font_cache:
        return _emoji_font_cache[size]
    for p in _EMOJI_FONT_PATHS:
        if not os.path.exists(p):
            continue
        try_sizes = [size]
        if "Apple Color Emoji" in p:
            best = min(_APPLE_EMOJI_VALID_SIZES, key=lambda s: abs(s - size))
            try_sizes = [best]
        for sz in try_sizes:
            try:
                f = ImageFont.truetype(p, sz, index=0)
                _emoji_font_cache[size] = f
                return f
            except Exception:
                continue
    _emoji_font_cache[size] = None
    return None


APP_DIR = Path(__file__).parent
FONTS_DIR = APP_DIR / "fonts"

# ── self-heal: detect & reset known-bad session state, log actions ──

_KNOWN_BAD_FLAGS = [
    "preview_confirmed",
    "preview_gallery_confirmed",
    "_password_ok",
]

def _init_self_heal():
    if "_self_heal_log" not in st.session_state:
        st.session_state["_self_heal_log"] = []
    for flag in _KNOWN_BAD_FLAGS:
        if st.session_state.get(flag):
            st.session_state[flag] = False
            _log_self_heal(f"reset '{flag}' -> False (known blocker)")

def _log_self_heal(action: str):
    ts = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"[{ts}] {action}"
    log = st.session_state.get("_self_heal_log", [])
    log.append(entry)
    st.session_state["_self_heal_log"] = log[-200:]

def _build_tag():
    try:
        sha = _sp.check_output(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=str(APP_DIR), stderr=_sp.DEVNULL).decode().strip()
    except Exception:
        sha = "unknown"
    return f"build {sha} · {_dt.datetime.now(_dt.timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}"

_BUILD_TAG = _build_tag()

st.set_page_config(page_title="批量邀请函工作流-M2.0", page_icon="📨", layout="wide")
_init_self_heal()

for _k in ["all_img_data", "preview_imgs", "check_done", "check_issues",
           "list_issues", "fix_log_text",
           "_do_regen", "single_check_issues"]:
    if _k not in st.session_state:
        st.session_state[_k] = None
st.markdown(
    f"""
    <div class="apple-hero">
      <h1>批量邀请函工作流-M2.0</h1>
      <p>上传模板与名单，预览确认后一键批量生成并下载压缩包。</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Streamlit 内建上传组件默认是英文，这里统一覆盖成简体中文提示，并优化界面层级。
st.markdown(
    """
    <style>
    .block-container {
        max-width: 1200px;
        padding-top: 1.25rem;
        padding-bottom: 3rem;
    }
    .apple-hero {
        padding: 1.1rem 0 0.6rem 0;
        border-bottom: 1px solid rgba(120, 120, 128, 0.25);
        margin-bottom: 1rem;
    }
    .apple-hero h1 {
        font-size: clamp(2rem, 4vw, 3rem);
        line-height: 1.15;
        letter-spacing: -0.02em;
        margin: 0;
    }
    .apple-hero p {
        margin-top: 0.55rem;
        margin-bottom: 0.25rem;
        font-size: 1.05rem;
        color: rgba(128, 128, 132, 0.95);
    }
    .apple-section-title {
        margin: 1.1rem 0 0.35rem 0;
        font-size: 1.3rem;
        font-weight: 700;
        letter-spacing: -0.01em;
    }
    .apple-info-card {
        border: 1px solid rgba(120, 120, 128, 0.3);
        border-radius: 14px;
        padding: 0.8rem 0.95rem;
        margin: 0.2rem 0 0.8rem 0;
        background: rgba(120, 120, 128, 0.08);
    }
    .apple-info-card strong {
        display: block;
        margin-bottom: 0.25rem;
        font-size: 1rem;
    }
    .apple-info-card span {
        font-size: 0.92rem;
        color: rgba(128, 128, 132, 0.95);
    }
    .action-card {
        border: 1px solid rgba(120, 120, 128, 0.35);
        border-radius: 12px;
        padding: 0.65rem 0.8rem;
        margin-bottom: 0.55rem;
        background: rgba(120, 120, 128, 0.06);
    }
    .action-card strong {
        display: block;
        font-size: 0.98rem;
        margin-bottom: 0.2rem;
    }
    .action-card span {
        display: block;
        font-size: 0.86rem;
        color: rgba(128, 128, 132, 0.95);
    }
    /* file_uploader: safe minimal style — no layout/interaction overrides */
    [data-testid="stFileUploaderDropzoneInstructions"] > div > small {
        font-size: 0.82rem;
    }
    /* Apple buttons */
    [data-testid="stButton"] > button, [data-testid="stDownloadButton"] > button {
        border-radius: 980px; min-height: 2.75rem; padding: 0 1.5rem;
        font-size: 0.94rem; font-weight: 400;
    }
    .check-pass-banner {
        border: 2px solid #38a169;
        border-radius: 12px;
        padding: 0.7rem 1rem;
        margin-bottom: 0.55rem;
        background: #f0fff4;
        text-align: center;
        font-weight: 500;
        color: #38a169;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

_heal_log = st.session_state.get("_self_heal_log", [])

PSD_EXTENSIONS = (".psd", ".psb")
IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".tiff", ".tif", ".bmp", ".webp")
VECTOR_EXTENSIONS = (".pdf", ".eps", ".ai")
ALL_TEMPLATE_TYPES = [e.lstrip(".") for e in PSD_EXTENSIONS + IMAGE_EXTENSIONS + VECTOR_EXTENSIONS]

LIST_EXTENSIONS = ["csv", "xlsx", "xls"]

# ── helpers ──────────────────────────────────────────────


@st.cache_data(ttl=300)
def scan_fonts():
    """Scan bundled fonts dir + system dirs. Returns {display: path}.
    Also populates _PS_NAME_CACHE so build_local_font_index doesn't re-read files."""
    def _is_garbled_text(s: str) -> bool:
        s = str(s or "")
        if not s.strip():
            return True
        bad = sum(1 for ch in s if ch == "?" or ch == "\ufffd" or ord(ch) < 32)
        return bad / max(1, len(s)) > 0.3

    fonts = {}
    ps_cache = {}
    search_roots = [str(FONTS_DIR)]
    for sys_dir in ["/System/Library/Fonts", "/System/Library/Fonts/Supplemental",
                    "/Library/Fonts", os.path.expanduser("~/Library/Fonts"),
                    "/usr/share/fonts", "/usr/local/share/fonts"]:
        if os.path.isdir(sys_dir):
            search_roots.append(sys_dir)

    for root_dir in search_roots:
        for dirpath, _dirnames, filenames in os.walk(root_dir):
            for f in filenames:
                if not any(f.lower().endswith(ext) for ext in (".ttf", ".otf", ".ttc")):
                    continue
                path = os.path.join(dirpath, f)
                try:
                    font_obj = ImageFont.truetype(path, 20)
                    family, style = font_obj.getname()
                    stem = Path(f).stem
                    if _is_garbled_text(family):
                        display = f"{stem} ({style})" if style and not _is_garbled_text(style) else stem
                    else:
                        display = f"{family} ({style})"
                    fonts[display] = path
                    ps_cache[path] = (family, style)
                except Exception:
                    pass
    return dict(sorted(fonts.items())), ps_cache


def _normalize_font_token(name: str) -> str:
    if not name:
        return ""
    s = str(name).lower().replace("-", " ").replace("_", " ")
    s = _re.sub(r"[^a-z0-9\u4e00-\u9fff]+", " ", s)
    parts = [p for p in s.split() if p]
    stop = {
        "regular", "normal", "medium", "bold", "semibold", "demibold",
        "black", "heavy", "light", "thin", "italic", "oblique", "book",
        "roman", "std", "mt", "ps", "pro",
    }
    filtered = [p for p in parts if p not in stop]
    return "".join(filtered) if filtered else "".join(parts)


def extract_psd_font_candidates(psd):
    """Extract font names from PSD text layers resource dict."""
    out = []
    seen = set()
    for tl in psd.descendants():
        if tl.kind != "type":
            continue
        try:
            ss = tl.engine_dict["StyleRun"]["RunArray"][0]["StyleSheet"]["StyleSheetData"]
            fs = tl.engine_dict.get("ResourceDict", {}).get("FontSet", [])
            fi = int(ss.get("Font", 0))
            if fi >= len(fs):
                continue
            entry = fs[fi]
            raw_name = str(entry.get("Name", "")).strip()
            family = str(entry.get("FamilyName", "")).strip()
            style = str(entry.get("StyleName", "")).strip()
            cand = {
                "raw": raw_name, "family": family, "style": style,
                "token_raw": _normalize_font_token(raw_name),
                "token_family": _normalize_font_token(family),
            }
            key = (cand["raw"], cand["family"], cand["style"])
            if key not in seen:
                seen.add(key)
                out.append(cand)
        except Exception:
            continue
    return out


def build_local_font_index(fonts_dict, ps_cache=None):
    """Build strict token->font-display/path index.
    ps_cache: {path: (family, style)} from scan_fonts() to avoid re-reading files.
    """
    index = {}
    for display, path in (fonts_dict or {}).items():
        family = display
        style = ""
        if display.endswith(")") and " (" in display:
            family, style = display.rsplit(" (", 1)
            style = style[:-1]
        tokens = {
            _normalize_font_token(display),
            _normalize_font_token(family),
            _normalize_font_token(f"{family} {style}"),
            _normalize_font_token(Path(path).stem),
        }
        ps_info = (ps_cache or {}).get(path)
        if ps_info:
            ps_family, ps_style = ps_info
            tokens.update({
                _normalize_font_token(ps_family),
                _normalize_font_token(f"{ps_family} {ps_style}"),
            })
        else:
            try:
                ps_family, ps_style = ImageFont.truetype(str(path), 20).getname()
                tokens.update({
                    _normalize_font_token(ps_family),
                    _normalize_font_token(f"{ps_family} {ps_style}"),
                })
            except Exception:
                pass
        for tk in [t for t in tokens if t]:
            if tk not in index:
                index[tk] = {"display": display, "path": path}
    return index


def match_psd_fonts_to_local(psd_fonts, local_index):
    matched, unmatched = [], []
    for cand in psd_fonts or []:
        pick = None
        for tk in [cand.get("token_raw", ""), cand.get("token_family", "")]:
            if tk and tk in local_index:
                pick = local_index[tk]
                break
        if pick:
            matched.append({"psd": cand, "local": pick})
        else:
            unmatched.append(cand)
    recommended = matched[0]["local"]["display"] if matched else None
    return matched, unmatched, recommended


def extract_per_layer_font(psd):
    """Return {layer_name: psd_font_name} for PSD text layers."""
    out = {}
    for tl in psd.descendants():
        if tl.kind != "type":
            continue
        try:
            ss = tl.engine_dict["StyleRun"]["RunArray"][0]["StyleSheet"]["StyleSheetData"]
            fs = tl.engine_dict.get("ResourceDict", {}).get("FontSet", [])
            fi = int(ss.get("Font", 0))
            if fi < len(fs):
                out[tl.name] = str(fs[fi].get("Name", "")).strip()
        except Exception:
            continue
    return out


def resolve_per_layer_font_path(layer_font_map, local_index, fallback_path):
    """Resolve each layer font to a local font path with fallback.

    Matching order:
    1) strict token match (_normalize_font_token)
    2) fuzzy substring match (uploaded/local font stem appears in PSD font name)
    """
    resolved = {}
    # build fuzzy stems from local_index values (display/path)
    stems = []
    for v in (local_index or {}).values():
        p = str(v.get("path", ""))
        if not p:
            continue
        stems.append(_normalize_font_token(Path(p).stem))
    stems = [s for s in dict.fromkeys(stems) if s]

    for layer_name, ps_name in (layer_font_map or {}).items():
        token = _normalize_font_token(ps_name)
        pick = None
        if token and token in (local_index or {}):
            pick = local_index[token]["path"]
        else:
            # fuzzy: find any stem contained in PSD token
            for stoken in stems:
                if stoken and token and (stoken in token or token in stoken):
                    # find first matching path for this stem
                    for v in (local_index or {}).values():
                        p = str(v.get("path", ""))
                        if p and _normalize_font_token(Path(p).stem) == stoken:
                            pick = p
                            break
                    if pick:
                        break
        resolved[layer_name] = pick or fallback_path
    return resolved


def push_font_history(display: str, path: str, max_items: int = 20):
    if not display or not path:
        return
    hist = st.session_state.get("_font_history", [])
    keep = []
    for it in hist:
        if not isinstance(it, dict):
            continue
        p = str(it.get("path", ""))
        d = str(it.get("display", ""))
        if p and d and os.path.exists(p):
            keep.append({"display": d, "path": p})
    keep = [it for it in keep if it["path"] != path]
    keep.insert(0, {"display": display, "path": path})
    st.session_state["_font_history"] = keep[:max_items]


def get_default_font_path():
    medium = FONTS_DIR / "OPPOSans-Medium.ttf"
    if medium.exists():
        return str(medium)
    bundled = FONTS_DIR / "OPPOSans4.ttf"
    if bundled.exists():
        return str(bundled)
    fonts, _ = scan_fonts()
    if fonts:
        return next(iter(fonts.values()))
    return None


def file_suffix(uploaded):
    return Path(uploaded.name).suffix.lower()


def load_psd(uploaded):
    suffix = file_suffix(uploaded)
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(uploaded.getvalue())
        tmp.flush()
        return PSDImage.open(tmp.name)


def _template_cache_key(uploaded):
    return f"{uploaded.name}:{uploaded.size}"


def _load_pdf_as_image(path: str) -> "Image.Image | None":
    """Convert first page of PDF to RGBA image using PyMuPDF."""
    try:
        import fitz  # pymupdf
        doc = fitz.open(path)
        page = doc[0]
        zoom = 3.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        doc.close()
        return img.convert("RGBA")
    except ImportError:
        st.error("PDF 支持需要 pymupdf 库。请在 requirements.txt 中添加 pymupdf。")
        return None
    except Exception as e:
        st.error(f"PDF 解析失败: {e}")
        return None


def load_image(uploaded):
    suffix = file_suffix(uploaded)
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(uploaded.getvalue())
        tmp.flush()
        if suffix == ".pdf":
            return _load_pdf_as_image(tmp.name)
        try:
            img = Image.open(tmp.name).convert("RGBA")
            return img
        except Exception as e:
            st.error(f"无法打开此文件格式。错误: {e}")
            return None


def get_text_layers(psd):
    return [l for l in psd.descendants() if l.kind == "type"]


def get_qr_layer(psd):
    _QR_KEYWORDS = [
        "二维码", "qr", "qrcode", "扫码", "扫一扫",
        "code", "barcode", "链接", "link", "url",
        "直播码", "小程序码", "微信码",
    ]
    for l in psd.descendants():
        name_lower = l.name.lower()
        if any(kw in name_lower for kw in _QR_KEYWORDS):
            if l.kind in ("smartobject", "pixel", "group"):
                return l
    for l in psd.descendants():
        name_lower = l.name.lower()
        if any(kw in name_lower for kw in _QR_KEYWORDS):
            return l
    if _HAS_CV2:
        try:
            full = psd.composite().convert("RGB")
            arr = np.array(full)
            gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
            detector = cv2.QRCodeDetector()
            _, pts, _ = detector.detectAndDecode(gray)
            if pts is not None and len(pts) > 0:
                xs = pts[0][:, 0]
                ys = pts[0][:, 1]
                pad = 20
                x1 = max(0, int(xs.min()) - pad)
                y1 = max(0, int(ys.min()) - pad)
                x2 = min(full.width, int(xs.max()) + pad)
                y2 = min(full.height, int(ys.max()) + pad)
                class _FakeLayer:
                    def __init__(self, box):
                        self.left, self.top, self.right, self.bottom = box
                        self.name = "auto_detected_qr"
                        self.kind = "pixel"
                    def composite(self):
                        return full.crop((self.left, self.top, self.right, self.bottom)).convert("RGBA")
                return _FakeLayer((x1, y1, x2, y2))
        except Exception:
            pass
    return None


def detect_qr_region(psd):
    qr_layer = get_qr_layer(psd)
    if qr_layer is not None:
        return (qr_layer.left, qr_layer.top, qr_layer.right, qr_layer.bottom)
    return None


def extract_qr_mask(psd):
    """Extract the real alpha mask from the QR smart object layer."""
    qr_layer = get_qr_layer(psd)
    if qr_layer is None:
        return None
    try:
        layer_img = qr_layer.composite()
        if layer_img and layer_img.mode == "RGBA":
            return layer_img.split()[3]
    except Exception:
        pass
    return None


def composite_background(psd):
    for l in psd.descendants():
        if l.kind == "type":
            l.visible = False
    return psd.composite()


def rounded_corner_mask(size, radius):
    mask = Image.new("L", size, 255)
    d = ImageDraw.Draw(mask)
    r = radius
    d.rectangle([0, 0, r, r], fill=0)
    d.rectangle([size[0] - r, 0, size[0], r], fill=0)
    d.rectangle([0, size[1] - r, r, size[1]], fill=0)
    d.rectangle([size[0] - r, size[1] - r, size[0], size[1]], fill=0)
    d.pieslice([0, 0, r * 2, r * 2], 180, 270, fill=255)
    d.pieslice([size[0] - r * 2, 0, size[0], r * 2], 270, 360, fill=255)
    d.pieslice([0, size[1] - r * 2, r * 2, size[1]], 90, 180, fill=255)
    d.pieslice([size[0] - r * 2, size[1] - r * 2, size[0], size[1]], 0, 90, fill=255)
    return mask


def replace_qr(background, qr_image, qr_box, original_layer_img=None, corner_radius=0):
    try:
        tw = qr_box[2] - qr_box[0]
        th = qr_box[3] - qr_box[1]
        if tw <= 0 or th <= 0:
            return background
        qr_resized = qr_image.convert('RGBA').resize((tw, th), Image.LANCZOS)
        mask = None
        if original_layer_img is not None:
            if original_layer_img.mode == 'L':
                mask = original_layer_img.resize((tw, th), Image.LANCZOS)
            else:
                layer_resized = original_layer_img.convert('RGBA').resize((tw, th), Image.LANCZOS)
                mask = layer_resized.split()[3]
        elif corner_radius > 0:
            mask = rounded_corner_mask((tw, th), corner_radius)
        if mask:
            bg_region = background.crop(qr_box).convert('RGBA')
            composite = Image.composite(qr_resized, bg_region, mask)
            background.paste(composite, (qr_box[0], qr_box[1]))
        else:
            background.paste(qr_resized, (qr_box[0], qr_box[1]))
    except Exception:
        background.paste(
            qr_image.convert('RGBA').resize(
                (qr_box[2] - qr_box[0], qr_box[3] - qr_box[1]), Image.LANCZOS),
            (qr_box[0], qr_box[1]))
    return background


def calibrate_font_size(font_path, text, target_height, raw_size):
    """Find Pillow font size where text height matches PSD layer height."""
    best_size = int(raw_size)
    best_diff = 999
    for s in range(int(raw_size), int(raw_size) + 20):
        f = ImageFont.truetype(font_path, s)
        bb = f.getbbox(text)
        h = bb[3] - bb[1]
        diff = abs(h - target_height)
        if diff < best_diff:
            best_diff = diff
            best_size = s
        if h >= target_height:
            return s
    return best_size


def get_font_color(psd):
    """Extract text color from the first PSD text layer (global fallback)."""
    for l in psd.descendants():
        if l.kind == "type":
            try:
                ss = l.engine_dict["StyleRun"]["RunArray"][0]["StyleSheet"]["StyleSheetData"]
                vals = ss.get("FillColor", {}).get("Values", [1.0, 0.0, 0.0, 0.0])
                if len(vals) == 4:
                    r, g, b = int(vals[1] * 255), int(vals[2] * 255), int(vals[3] * 255)
                    return (r, g, b, 255)
                elif len(vals) == 2:
                    gray = int(vals[1] * 255)
                    return (gray, gray, gray, 255)
            except Exception:
                continue
    return (255, 255, 255, 255)


def extract_per_layer_color(psd):
    """Return {layer_name: (R, G, B, A)} for each text layer."""
    out = {}
    for l in psd.descendants():
        if l.kind != "type":
            continue
        try:
            ss = l.engine_dict["StyleRun"]["RunArray"][0]["StyleSheet"]["StyleSheetData"]
            vals = ss.get("FillColor", {}).get("Values", [1.0, 0.0, 0.0, 0.0])
            if len(vals) == 4:
                r, g, b = int(vals[1] * 255), int(vals[2] * 255), int(vals[3] * 255)
                out[l.name] = (r, g, b, 255)
            elif len(vals) == 2:
                gray = int(vals[1] * 255)
                out[l.name] = (gray, gray, gray, 255)
        except Exception:
            continue
    return out


def get_text_layer_positions(psd, font_path=None, per_layer_fonts=None):
    """Return {layer_name: (center_y, calibrated_font_size, psd_width, stroke_width, center_x, align)}.

    center_x: horizontal center of the PSD text layer.
    align: 'center' if layer center is near image center, else 'left'.
    """
    img_w = psd.width
    positions = {}
    for l in psd.descendants():
        if l.kind == "type":
            cy = (l.top + l.bottom) // 2
            cx = (l.left + l.right) // 2
            layer_cx_ratio = abs(cx - img_w / 2) / max(1, img_w)
            align = "center" if layer_cx_ratio < 0.05 else "left"
            ss = l.engine_dict["StyleRun"]["RunArray"][0]["StyleSheet"]["StyleSheetData"]
            raw_size = ss.get("FontSize", 51)
            cal_size = int(raw_size)
            layer_font_path = (per_layer_fonts or {}).get(l.name, font_path)
            if layer_font_path and l.height > 0 and len(l.text) >= 1:
                cal_size = calibrate_font_size(layer_font_path, l.text, l.height, raw_size)

            stroke_w = 0
            faux_bold = ss.get("FauxBold", False)
            if faux_bold:
                stroke_w = max(1, round(cal_size / 36))

            if stroke_w == 0:
                try:
                    fs = l.engine_dict.get("ResourceDict", {}).get("FontSet", [])
                    font_idx = int(ss.get("Font", 0))
                    if font_idx < len(fs):
                        ps_name = fs[font_idx].get("Name", "").lower()
                        if any(w in ps_name for w in
                               ["bold", "heavy", "black", "semibold", "demibold", "medium"]):
                            stroke_w = max(1, round(cal_size / 36))
                except Exception:
                    pass

            if stroke_w == 0:
                try:
                    layer_img = l.composite()
                    if layer_img is not None:
                        arr = np.array(layer_img.convert("L"), dtype=np.float64)
                        bright = arr > 80
                        if bright.any():
                            cov = float(bright.sum()) / float(arr.size)
                            if cov > 0.35:
                                stroke_w = max(1, round(cal_size / 40))
                except Exception:
                    pass

            positions[l.name] = (cy, cal_size, l.width, stroke_w, cx, align)
    return positions


def _calibrate_single_stroke(layer, original_img, bg, font_path, font_size, color, img_width, center_y):
    """Find stroke_width whose ink volume best matches the PSD render."""
    text = layer.text.strip()
    if not text or layer.height < 5:
        return 0

    pad = 20
    box = (max(0, layer.left - pad), max(0, layer.top - pad),
           min(original_img.width, layer.right + pad),
           min(original_img.height, layer.bottom + pad))

    ori_crop = np.array(original_img.crop(box).convert("L"), dtype=np.float64)
    bg_crop_np = np.array(bg.crop(box).convert("L"), dtype=np.float64)
    bg_crop_img = bg.crop(box).convert("RGBA")

    psd_ink = np.sum(np.abs(ori_crop - bg_crop_np))
    if psd_ink < 100:
        return 0

    f = ImageFont.truetype(font_path, font_size)
    bbox = f.getbbox(text)
    text_w = bbox[2] - bbox[0]
    x = (img_width - text_w) // 2 - box[0]
    y = center_y - (bbox[1] + bbox[3]) // 2 - box[1]

    best_sw = 0
    best_diff = float("inf")
    for sw in range(0, 5):
        test = bg_crop_img.copy()
        draw = ImageDraw.Draw(test)
        if sw > 0:
            draw.text((x, y), text, font=f, fill=color,
                      stroke_width=sw, stroke_fill=color)
        else:
            draw.text((x, y), text, font=f, fill=color)

        test_np = np.array(test.convert("L"), dtype=np.float64)
        pil_ink = np.sum(np.abs(test_np - bg_crop_np))

        diff = abs(psd_ink - pil_ink)
        if diff < best_diff:
            best_diff = diff
            best_sw = sw
    return best_sw


def calibrate_stroke_weights(psd, positions, original_img, bg, font_path, color, img_width):
    """Override stroke_width in positions using pixel-level ink comparison."""
    updated = {}
    for l in psd.descendants():
        if l.kind == "type" and l.name in positions:
            pos = positions[l.name]
            cy, cal_size, lw, old_sw = pos[0], pos[1], pos[2], pos[3]
            extra = pos[4:] if len(pos) > 4 else ()
            try:
                sw = _calibrate_single_stroke(
                    l, original_img, bg, font_path, cal_size, color, img_width, cy)
            except Exception:
                sw = old_sw
            updated[l.name] = (cy, cal_size, lw, sw) + extra

    for k in positions:
        if k not in updated:
            updated[k] = positions[k]
    return updated


def draw_centered_text(draw, font, text, center_y, img_width, color, stroke_width=0, target_img=None, align="center"):
    """Draw text (centered or left-aligned) with emoji fallback."""
    if not text or not text.strip():
        return

    text_clean = _strip_emoji(text)
    has_emoji = (text_clean != text)
    emoji_font = _get_emoji_font(font.size) if has_emoji else None

    if not has_emoji or not emoji_font:
        render_text = text_clean or text
        if not render_text.strip():
            return
        bbox = font.getbbox(render_text)
        text_w = bbox[2] - bbox[0]
        x = (img_width - text_w) // 2 if align == "center" else max(0, (img_width - text_w) // 2)
        y = center_y - (bbox[1] + bbox[3]) // 2
        if stroke_width > 0:
            draw.text((x, y), render_text, font=font, fill=color,
                      stroke_width=stroke_width, stroke_fill=color)
        else:
            draw.text((x, y), render_text, font=font, fill=color)
        return

    segments = []
    cur = ""
    cur_is_emoji = False
    for ch in text:
        ie = _is_emoji(ch)
        if ie != cur_is_emoji and cur:
            segments.append((cur, cur_is_emoji))
            cur = ""
        cur += ch
        cur_is_emoji = ie
    if cur:
        segments.append((cur, cur_is_emoji))

    total_w = 0
    seg_metrics = []
    for seg_text, is_em in segments:
        f = emoji_font if is_em else font
        bbox = f.getbbox(seg_text)
        w = bbox[2] - bbox[0]
        total_w += w
        seg_metrics.append((seg_text, is_em, w, bbox))

    main_bbox = font.getbbox(text_clean or "A")
    text_top = main_bbox[1]
    text_h = main_bbox[3] - main_bbox[1]
    draw_y = center_y - (main_bbox[1] + main_bbox[3]) // 2

    em_target_h = int(text_h * 0.92)
    em_render_size = max(8, em_target_h)
    ef_sized = _get_emoji_font(em_render_size) if has_emoji else None

    if ef_sized:
        recalc_w = 0
        new_metrics = []
        for seg_text, is_em, w, bbox in seg_metrics:
            if is_em:
                eb = ef_sized.getbbox(seg_text)
                ew = eb[2] - eb[0] if eb else em_render_size
                new_metrics.append((seg_text, True, ew, eb))
                recalc_w += ew
            else:
                new_metrics.append((seg_text, False, w, bbox))
                recalc_w += w
        seg_metrics = new_metrics
        total_w = recalc_w

    cursor_x = (img_width - total_w) // 2

    for seg_text, is_em, w, bbox in seg_metrics:
        if is_em and target_img and ef_sized:
            canvas_sz = em_render_size + 20
            em_layer = Image.new("RGBA", (canvas_sz, canvas_sz), (0, 0, 0, 0))
            em_draw = ImageDraw.Draw(em_layer)
            em_draw.text((0, 0), seg_text, font=ef_sized, fill=(255, 255, 255, 255),
                         embedded_color=True)
            em_bbox = em_layer.getbbox()
            if em_bbox:
                em_layer = em_layer.crop(em_bbox)
            em_w, em_h = em_layer.size
            text_visual_top = draw_y + text_top
            text_visual_bottom = text_visual_top + text_h
            text_visual_center = (text_visual_top + text_visual_bottom) // 2
            paste_y = text_visual_center - em_h // 2
            paste_x = int(cursor_x)
            target_img.paste(em_layer, (paste_x, int(paste_y)), em_layer)
        else:
            seg_y = draw_y
            if stroke_width > 0:
                draw.text((cursor_x, seg_y), seg_text, font=font, fill=color,
                          stroke_width=stroke_width, stroke_fill=color)
            else:
                draw.text((cursor_x, seg_y), seg_text, font=font, fill=color)
        cursor_x += w


def _compute_ssim(img_a, img_b):
    """Lightweight SSIM between two same-sized grayscale uint8 numpy arrays."""
    if not _HAS_CV2 or img_a.shape != img_b.shape or img_a.size == 0:
        return 1.0
    C1 = (0.01 * 255) ** 2
    C2 = (0.03 * 255) ** 2
    a = img_a.astype(np.float64)
    b = img_b.astype(np.float64)
    k = (11, 11)
    mu_a = cv2.GaussianBlur(a, k, 1.5)
    mu_b = cv2.GaussianBlur(b, k, 1.5)
    sig_a2 = cv2.GaussianBlur(a * a, k, 1.5) - mu_a * mu_a
    sig_b2 = cv2.GaussianBlur(b * b, k, 1.5) - mu_b * mu_b
    sig_ab = cv2.GaussianBlur(a * b, k, 1.5) - mu_a * mu_b
    num = (2 * mu_a * mu_b + C1) * (2 * sig_ab + C2)
    den = (mu_a ** 2 + mu_b ** 2 + C1) * (sig_a2 + sig_b2 + C2)
    ssim_map = num / den
    return float(ssim_map.mean())


def generate_one(background, text_items, img_width, color, font_path):
    """text_items: list of (text, cy, fsize[, stroke_w[, font_path[, color[, align]]]])."""
    img = background.copy()
    draw = ImageDraw.Draw(img)
    for item in text_items:
        text, cy, fsize = item[0], item[1], item[2]
        sw = item[3] if len(item) > 3 else 0
        item_font_path = item[4] if len(item) > 4 and item[4] else font_path
        item_color = item[5] if len(item) > 5 and item[5] else color
        item_align = item[6] if len(item) > 6 else "center"
        if text:
            f = ImageFont.truetype(item_font_path, fsize)
            draw_centered_text(draw, f, text, cy, img_width, item_color,
                               stroke_width=sw, target_img=img, align=item_align)
    return img


def check_image_quality(img, text_items, img_width, qr_box, font_path,
                        original_img=None, replaced_qr_img=None):
    """Quality checks: text bounds, spacing, triple QR validation.
    replaced_qr_img: if user uploaded a replacement QR, pass the PIL Image here
    to use replacement-specific checks instead of original-comparison checks.
    """
    issues = []

    for item in text_items:
        text, center_y, fsize = item[0], item[1], item[2]
        text = _strip_emoji(text) if text else text
        if not text:
            continue
        f = ImageFont.truetype(font_path, fsize)
        bbox = f.getbbox(text)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        x = (img_width - text_w) // 2
        draw_y = center_y - (bbox[1] + bbox[3]) // 2
        if x < 0:
            issues.append(("error", "文字「{}」超出图片宽度".format(text)))
        elif x < 20:
            issues.append(("warning", "文字「{}」距边缘太近 ({}px)".format(text, int(x))))
        if draw_y < 0 or draw_y + text_h > img.size[1]:
            issues.append(("error", "文字「{}」超出图片高度".format(text)))

    if qr_box:
        pad = 10
        crop_box = (max(0, qr_box[0] - pad), max(0, qr_box[1] - pad),
                    min(img.width, qr_box[2] + pad), min(img.height, qr_box[3] + pad))
        try:
            gen_qr_crop = img.crop(crop_box).convert("RGB")
            gen_qr_arr = np.array(gen_qr_crop)
            tw = qr_box[2] - qr_box[0]
            th = qr_box[3] - qr_box[1]

            if replaced_qr_img is not None:
                # ── 用户上传了替换二维码：与上传的新QR对比 ──
                ref_resized = replaced_qr_img.convert("RGB").resize((tw + pad * 2, th + pad * 2), Image.LANCZOS)
                ref_arr = np.array(ref_resized)
                if _HAS_CV2 and ref_arr.shape == gen_qr_arr.shape:
                    ref_gray = cv2.cvtColor(ref_arr, cv2.COLOR_RGB2GRAY)
                    gen_gray = cv2.cvtColor(gen_qr_arr, cv2.COLOR_RGB2GRAY)
                    ssim_val = _compute_ssim(ref_gray, gen_gray)
                    if ssim_val > 0.5:
                        issues.append(("success", "替换二维码结构一致 (SSIM {:.2f})".format(ssim_val)))
                    elif ssim_val > 0.3:
                        issues.append(("warning", "替换二维码与上传图有差异 (SSIM {:.2f})，可能是圆角裁切".format(ssim_val)))
                    else:
                        issues.append(("error", "替换二维码与上传图差异过大 (SSIM {:.2f})，替换可能未生效".format(ssim_val)))
                    ref_sharp = cv2.Laplacian(ref_gray, cv2.CV_64F).var()
                    gen_sharp = cv2.Laplacian(gen_gray, cv2.CV_64F).var()
                    if ref_sharp > 0:
                        ratio = gen_sharp / ref_sharp
                        if ratio > 0.6:
                            issues.append(("success", "替换二维码清晰度正常 (比值 {:.2f})".format(ratio)))
                        else:
                            issues.append(("warning", "替换二维码清晰度下降 (比值 {:.2f})".format(ratio)))
                else:
                    issues.append(("info", "跳过替换二维码 SSIM 对比（尺寸不匹配或无 OpenCV）"))
            elif original_img and _HAS_CV2:
                # ── 未替换二维码：与原始模板 1:1 对比 ──
                try:
                    ori_crop_box = (max(0, qr_box[0] - pad), max(0, qr_box[1] - pad),
                                    min(original_img.width, qr_box[2] + pad),
                                    min(original_img.height, qr_box[3] + pad))
                    ori_qr_crop = original_img.crop(ori_crop_box).convert("RGB")
                    ori_qr_arr = np.array(ori_qr_crop)
                    if ori_qr_arr.shape == gen_qr_arr.shape:
                        ori_gray = cv2.cvtColor(ori_qr_arr, cv2.COLOR_RGB2GRAY)
                        gen_gray = cv2.cvtColor(gen_qr_arr, cv2.COLOR_RGB2GRAY)
                        ssim_val = _compute_ssim(ori_gray, gen_gray)
                        if ssim_val > 0.85:
                            issues.append(("success", "二维码结构校验通过 (SSIM {:.2f})".format(ssim_val)))
                        elif ssim_val > 0.6:
                            issues.append(("warning", "二维码结构有差异 (SSIM {:.2f})，可能是圆角/边缘变化".format(ssim_val)))
                        else:
                            issues.append(("error", "二维码结构严重偏差 (SSIM {:.2f})，图形可能被破坏".format(ssim_val)))
                        ori_sharp = cv2.Laplacian(ori_gray, cv2.CV_64F).var()
                        gen_sharp = cv2.Laplacian(gen_gray, cv2.CV_64F).var()
                        if ori_sharp > 0:
                            ratio = gen_sharp / ori_sharp
                            if ratio > 0.85:
                                issues.append(("success", "二维码清晰度正常 (比值 {:.2f})".format(ratio)))
                            elif ratio > 0.5:
                                issues.append(("warning", "二维码清晰度下降 (比值 {:.2f})".format(ratio)))
                            else:
                                issues.append(("error", "二维码清晰度严重下降 (比值 {:.2f})".format(ratio)))
                except Exception:
                    issues.append(("warning", "二维码 SSIM/清晰度检查异常，建议人工扫码确认"))
            elif not _HAS_CV2:
                issues.append(("success", "OpenCV 未安装，跳过二维码结构与清晰度检查"))

            # 4c) decode check (always runs)
            qr_2x = Image.fromarray(gen_qr_arr).resize(
                (gen_qr_arr.shape[1] * 2, gen_qr_arr.shape[0] * 2), Image.LANCZOS)
            qr_2x_arr = np.array(qr_2x)
            if _HAS_CV2:
                detector = cv2.QRCodeDetector()
                data, det_bbox, _ = detector.detectAndDecode(qr_2x_arr)
                if data:
                    issues.append(("success", "二维码可解码：内容确认正常"))
                elif det_bbox is not None and len(det_bbox) > 0:
                    issues.append(("info", "检测到二维码定位方块但解码失败，建议手机扫码确认（不影响实际使用）"))
                else:
                    issues.append(("error", "未检测到二维码定位方块，可能严重乱码或裁切"))
            else:
                issues.append(("success", "二维码已替换，建议手机扫码确认"))
        except Exception:
            issues.append(("warning", "二维码检查异常，请人工扫码复核"))

    return issues


def compare_preview_quality(original_img, preview_img, text_items, img_width, qr_box, font_path, use_custom_font=False):
    """Compare original vs preview for diff-based quality checks."""
    issues = []

    # 2) 字体应用是否符合指定字体策略
    font_name = Path(font_path).name.lower() if font_path else ""
    if use_custom_font:
        issues.append(("success", f"已应用自定义字体: {Path(font_path).name}"))
    else:
        if "oppo" in font_name:
            issues.append(("success", "字体检查通过：已使用 OPPO 字体"))
        else:
            issues.append(("warning", f"字体可能不是指定 OPPO 字体（当前: {Path(font_path).name}）"))

    # 1) 字体大小、字距、边距、居中、对齐 + 5) 空格
    for item in text_items:
        text, center_y, fsize = item[0], item[1], item[2]
        text = _strip_emoji(text) if text else text
        if not text:
            issues.append(("warning", "检测到空文本，可能导致内容缺失"))
            continue

        if text != text.strip():
            issues.append(("warning", f"文本「{text}」存在首尾空格"))
        if "  " in text:
            issues.append(("warning", f"文本「{text}」存在连续空格"))

        f = ImageFont.truetype(font_path, fsize)
        bbox = f.getbbox(text)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        x = (img_width - text_w) // 2
        y = center_y - (bbox[1] + bbox[3]) // 2
        left_gap = x
        right_gap = img_width - (x + text_w)
        if abs(left_gap - right_gap) > 6:
            issues.append(("warning", f"文本「{text}」左右边距不平衡（差值 {abs(left_gap-right_gap)}px）"))
        draw_center = y + text_h / 2
        v_offset = abs(draw_center - center_y)
        if v_offset > 15:
            issues.append(("warning", f"文本「{text}」垂直对齐偏差 {v_offset:.1f}px"))
        elif v_offset > 3:
            issues.append(("info", f"文本「{text}」垂直对齐偏差 {v_offset:.1f}px（不影响实际使用）"))

    # 3) 背景清晰和细节是否变化
    try:
        ori_rgb = np.array(original_img.convert("RGB"), dtype=np.int16)
        pre_rgb = np.array(preview_img.convert("RGB"), dtype=np.int16)
        if ori_rgb.shape == pre_rgb.shape:
            h, w, _ = ori_rgb.shape
            mask = np.ones((h, w), dtype=np.uint8)

            for item in text_items:
                text, center_y, fsize = item[0], item[1], item[2]
                if not text:
                    continue
                f = ImageFont.truetype(font_path, fsize)
                bbox = f.getbbox(text)
                text_w = bbox[2] - bbox[0]
                text_h = bbox[3] - bbox[1]
                x = max(0, (img_width - text_w) // 2 - 18)
                y = max(0, int(center_y - (bbox[1] + bbox[3]) // 2) - 18)
                x2 = min(w, x + text_w + 36)
                y2 = min(h, y + text_h + 36)
                mask[y:y2, x:x2] = 0

            if qr_box:
                x1, y1, x2, y2 = qr_box
                x1 = max(0, x1 - 16)
                y1 = max(0, y1 - 16)
                x2 = min(w, x2 + 16)
                y2 = min(h, y2 + 16)
                mask[y1:y2, x1:x2] = 0

            diff = np.abs(pre_rgb - ori_rgb).mean(axis=2)
            bg_diff = float(diff[mask == 1].mean()) if np.any(mask == 1) else 0.0
            if bg_diff > 7.5:
                issues.append(("warning", f"背景差异偏大（均值 {bg_diff:.2f}），可能影响清晰度或细节"))
            else:
                issues.append(("success", "背景差异检查通过"))

            if _HAS_CV2:
                ori_gray = cv2.cvtColor(ori_rgb.astype(np.uint8), cv2.COLOR_RGB2GRAY)
                pre_gray = cv2.cvtColor(pre_rgb.astype(np.uint8), cv2.COLOR_RGB2GRAY)
                ori_sharp = cv2.Laplacian(ori_gray, cv2.CV_64F).var()
                pre_sharp = cv2.Laplacian(pre_gray, cv2.CV_64F).var()
                if ori_sharp > 0 and pre_sharp / ori_sharp < 0.82:
                    issues.append(("warning", f"替换图清晰度下降（{pre_sharp/ori_sharp:.2f}x）"))
    except Exception:
        issues.append(("warning", "背景清晰度检查失败，请人工放大复核细节"))

    # 4) 二维码清晰度（相对比较）
    if qr_box and _HAS_CV2:
        try:
            qx1, qy1, qx2, qy2 = qr_box
            ori_qr = original_img.crop((qx1, qy1, qx2, qy2)).convert("RGB")
            pre_qr = preview_img.crop((qx1, qy1, qx2, qy2)).convert("RGB")
            ori_g = cv2.cvtColor(np.array(ori_qr), cv2.COLOR_RGB2GRAY)
            pre_g = cv2.cvtColor(np.array(pre_qr), cv2.COLOR_RGB2GRAY)
            ori_s = cv2.Laplacian(ori_g, cv2.CV_64F).var()
            pre_s = cv2.Laplacian(pre_g, cv2.CV_64F).var()
            if ori_s > 0:
                r = pre_s / ori_s
                if r < 0.5:
                    issues.append(("error", "二维码清晰度严重下降 (比值 {:.2f})".format(r)))
                elif r < 0.85:
                    issues.append(("warning", "二维码清晰度下降 (比值 {:.2f})".format(r)))
        except Exception:
            issues.append(("warning", "二维码清晰度对比检查失败，请人工扫码复核"))

    # 6) 图片是否乱码（全图异常波动）
    try:
        ori = np.array(original_img.convert("RGB"), dtype=np.int16)
        pre = np.array(preview_img.convert("RGB"), dtype=np.int16)
        if ori.shape == pre.shape:
            global_diff = float(np.abs(pre - ori).mean())
            if global_diff > 20:
                issues.append(("warning", f"全图差异偏大（{global_diff:.2f}），可能存在异常渲染或乱码"))
    except Exception:
        pass

    return issues


def check_list_generation_quality(rows, enable_company, company_field, enable_name, name_field, filename_builder):
    """检查名单信息是否会导致生成错误（空值、异常字符、重名文件）。"""
    issues = []
    seen_names = {}
    for idx, row in enumerate(rows):
        row_no = idx + 1
        if enable_company and company_field:
            cv = str(row.get(company_field, "")).strip()
            if (not cv) or cv.lower() == "nan":
                issues.append(("error", f"第 {row_no} 条公司名为空"))
            if "\ufffd" in cv:
                issues.append(("error", f"第 {row_no} 条公司名包含异常字符"))
        if enable_name and name_field:
            nv = str(row.get(name_field, "")).strip()
            if (not nv) or nv.lower() == "nan":
                issues.append(("error", f"第 {row_no} 条人名为空"))
            if "\ufffd" in nv:
                issues.append(("error", f"第 {row_no} 条人名包含异常字符"))

        fname = filename_builder(row).strip()
        if not fname:
            issues.append(("error", f"第 {row_no} 条生成文件名为空"))
        elif fname in seen_names:
            issues.append(("error", f"文件名重复：第 {seen_names[fname]} 条 和 第 {row_no} 条 -> {fname}"))
        else:
            seen_names[fname] = row_no
    return issues


_FIX_SUGGESTIONS = {
    "超出图片宽度": "建议：减小字号 2-4px 或缩短文本内容",
    "距边缘太近": "建议：减小字号 1-2px 或缩短文本",
    "超出图片高度": "建议：调整垂直位置或减小字号",
    "字体可能不是指定": "建议：切换到 OPPO Sans 或上传指定字体文件",
    "首尾空格": "建议：已在生成时自动 trim，检查源数据",
    "连续空格": "建议：检查名单源数据是否有多余空格",
    "左右边距不平衡": "建议：属于字体 metrics 特性，通常可忽略",
    "垂直对齐偏差": "建议：属于字体 metrics 特性，通常可忽略",
    "背景差异偏大": "建议：检查模板是否为有损压缩格式，改用 PNG 无损模板",
    "清晰度下降": "建议：使用更高分辨率的模板或二维码源图",
    "二维码结构严重偏差": "建议：二维码图形被破坏，请重新上传二维码替换图",
    "二维码结构有差异": "建议：可能是圆角/边缘处理导致，建议手机扫码确认",
    "二维码清晰度严重下降": "建议：使用更高分辨率的二维码源图",
    "二维码清晰度下降": "建议：使用更高分辨率的二维码源图",
    "未检测到二维码定位方块": "建议：二维码严重损坏或被裁切，请重新上传",
    "解码失败": "建议：二维码可能受损，请手机扫码确认或重新上传",
    "全图差异偏大": "建议：模板可能已损坏，请重新上传原始模板",
    "公司名为空": "建议：检查名单中对应行的公司名字段",
    "人名为空": "建议：检查名单中对应行的人名字段",
    "异常字符": "建议：名单中包含乱码字符 (U+FFFD)，请修正源文件编码",
    "文件名为空": "建议：检查名单中公司名/人名字段是否为空",
    "文件名重复": "建议：名单中存在重名，请添加序号或其他区分信息",
}


def _suggest_fix(msg):
    for keyword, suggestion in _FIX_SUGGESTIONS.items():
        if keyword in msg:
            return suggestion
    return ""


def build_fix_log(all_check_issues, list_issues):
    lines = [
        "=" * 60,
        "  批量邀请函 — 质量检查 & 自动修复日志",
        "  时间: {}".format(_dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        "=" * 60,
        "",
    ]

    total_imgs = len(all_check_issues) if all_check_issues else 0
    err_count = 0
    warn_count = 0
    info_count = 0
    pass_count = 0
    auto_fixes = []

    if list_issues:
        for lvl, _ in list_issues:
            if lvl == "error":
                err_count += 1
            elif lvl == "warning":
                warn_count += 1
            elif lvl == "info":
                info_count += 1
    if all_check_issues:
        for _, issues in all_check_issues:
            img_has_problem = False
            for lvl, _ in issues:
                if lvl == "error":
                    err_count += 1
                    img_has_problem = True
                elif lvl == "warning":
                    warn_count += 1
                    img_has_problem = True
                elif lvl == "info":
                    info_count += 1
            if not img_has_problem:
                pass_count += 1

    lines.append("【摘要】")
    lines.append("  检查图片数: {}".format(total_imgs))
    lines.append("  通过: {}  |  提示: {}  |  警告: {}  |  错误: {}".format(pass_count, info_count, warn_count, err_count))
    lines.append("")

    lines.append("=" * 60)
    lines.append("【名单信息检查】")
    lines.append("-" * 40)
    if list_issues:
        for lvl, msg in list_issues:
            fix = _suggest_fix(msg)
            lines.append("  [{}] {}".format(lvl.upper(), msg))
            if fix:
                lines.append("        -> {}".format(fix))
                if lvl == "error":
                    auto_fixes.append((msg, fix))
    else:
        lines.append("  [PASS] 名单信息检查全部通过")
    lines.append("")

    lines.append("=" * 60)
    lines.append("【图片质量检查（14 项标准 + 三重二维码校验）】")
    lines.append("-" * 40)
    if all_check_issues:
        for fname, issues in all_check_issues:
            non_success = [(l, m) for l, m in issues if l != "success"]
            if not non_success:
                lines.append("  [PASS] {}: 全部检查通过".format(fname))
            else:
                lines.append("  [FILE] {}".format(fname))
                for lvl, msg in non_success:
                    fix = _suggest_fix(msg)
                    lines.append("    [{}] {}".format(lvl.upper(), msg))
                    if fix:
                        lines.append("          -> {}".format(fix))
                        if lvl == "error":
                            auto_fixes.append(("{}: {}".format(fname, msg), fix))
    else:
        lines.append("  [INFO] 未执行图片质量检查")
    lines.append("")

    lines.append("=" * 60)
    lines.append("【自动修复建议汇总】")
    lines.append("-" * 40)
    if auto_fixes:
        for i, (problem, fix) in enumerate(auto_fixes, 1):
            lines.append("  {}. 问题: {}".format(i, problem))
            lines.append("     修复: {}".format(fix))
    else:
        lines.append("  无需自动修复，所有错误级别问题为零。")
    lines.append("")
    lines.append("=" * 60)
    lines.append("  日志结束")
    lines.append("=" * 60)
    return "\n".join(lines) + "\n"


def parse_spreadsheet(uploaded):
    suffix = file_suffix(uploaded)
    try:
        if suffix == ".csv":
            text = uploaded.getvalue().decode("utf-8-sig")
            df = pd.read_csv(io.StringIO(text))
        elif suffix == ".xlsx":
            df = pd.read_excel(uploaded, engine="openpyxl")
        elif suffix == ".xls":
            df = pd.read_excel(uploaded, engine="xlrd")
        elif suffix == ".et":
            st.error("WPS .et 格式暂不支持，请另存为 .xlsx 后重新上传。")
            return [], []
        else:
            st.error(f"不支持的名单格式: {suffix}")
            return [], []
    except Exception as e:
        st.error(f"读取名单失败: {e}")
        return [], []

    df = df.dropna(how="all")
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
    df = df.dropna(axis=1, how="all")
    if df.empty:
        return [], []
    first_col = df.iloc[:, 0]
    df = df[first_col.notna() & (first_col.astype(str).str.strip() != "") & (first_col.astype(str) != "nan")]
    df.columns = [str(c).strip() for c in df.columns]
    fields = list(df.columns)
    df = df.fillna("")
    rows = df.astype(str).to_dict("records")
    for row in rows:
        for k, v in row.items():
            if v == "nan":
                row[k] = ""
    return rows, fields



def parse_manual_lines(raw_text, mode):
    rows = []
    for line in raw_text.splitlines():
        line = line.strip()
        if not line: continue
        if mode == "只有姓名":
            rows.append({"姓名": line})
            continue
        parts = [p.strip() for p in line.replace("，", ",").split(",", 1)]
        if len(parts) < 2 or not parts[0] or not parts[1]:
            return None, f'格式错误：{line}'
        rows.append({"姓名": parts[0], "公司名": parts[1]})
    return rows, None

@st.dialog("手动输入名单", width="large")
def manual_input_dialog():
    mode = st.radio("名单类型", ["只有姓名", "姓名和公司"], horizontal=True, key="manual_list_mode")
    placeholder = "张三\n李四\n王五"
    if mode == "姓名和公司": placeholder = "张三, ABC公司\n李四, XYZ集团"
    st.caption("每行一条名单")
    raw_text = st.text_area("名单内容", key="manual_list_raw_text", height=260, placeholder=placeholder)
    if st.button("确认", type="primary", use_container_width=True, key="manual_list_confirm_btn"):
        rows, err = parse_manual_lines(raw_text, mode)
        if err: st.error(err)
        elif not rows: st.warning('请至少输入一条名单。')
        else:
            st.session_state["manual_list_rows"] = rows

@st.dialog("输入问题", width="large")
def preview_issue_dialog():
    dcol1, dcol2 = st.columns(2)
    with dcol1:
        report = st.text_input("问题描述", key="dlg_issue_report", placeholder="请描述发现的问题...")
    with dcol2:
        if st.button("重新生成", type="primary", use_container_width=True, key="dlg_issue_regen"):
            if report.strip():
                st.session_state["_do_regen"] = True
            else:
                st.warning("请先输入问题描述")

    preview_in_dlg = st.session_state.get("_dlg_preview_img")
    if preview_in_dlg:
        buf = io.BytesIO()
        preview_in_dlg.save(buf, format="PNG")
        b64 = base64.b64encode(buf.getvalue()).decode("ascii")
        iw, ih = preview_in_dlg.size
        viewer_html = f"""
        <div id="viewer-wrap" style="height:500px;position:relative;overflow:hidden;border:1px solid rgba(120,120,128,.28);border-radius:12px;">
          <canvas id="viewer-canvas" style="width:100%;height:100%;display:block;cursor:grab;background:
            linear-gradient(45deg,#f5f5f7 25%,transparent 25%,transparent 75%,#f5f5f7 75%,#f5f5f7) 0 0/24px 24px,
            linear-gradient(45deg,#f5f5f7 25%,transparent 25%,transparent 75%,#f5f5f7 75%,#f5f5f7) 12px 12px/24px 24px,
            #fff;"></canvas>
          <div id="zoom-badge" style="position:absolute;left:10px;bottom:10px;background:rgba(29,29,31,.75);color:#fff;padding:4px 8px;border-radius:999px;font-size:12px;">100%</div>
        </div>
        <script>
          (() => {{
            const wrap = document.getElementById("viewer-wrap");
            const canvas = document.getElementById("viewer-canvas");
            const badge = document.getElementById("zoom-badge");
            const ctx = canvas.getContext("2d");
            const img = new Image();
            img.src = "data:image/png;base64,{b64}";

            let scale = 1.0;
            let minScale = 0.2;
            let maxScale = 8.0;
            let tx = 0, ty = 0;
            let dragging = false;
            let lx = 0, ly = 0;
            let pinchStartDist = 0;
            let pinchStartScale = 1;
            let pinchCenter = null;

            function resizeCanvas() {{
              const r = wrap.getBoundingClientRect();
              const dpr = window.devicePixelRatio || 1;
              canvas.width = Math.floor(r.width * dpr);
              canvas.height = Math.floor(r.height * dpr);
              canvas.style.width = r.width + "px";
              canvas.style.height = r.height + "px";
              ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
              draw();
            }}

            function draw() {{
              const r = wrap.getBoundingClientRect();
              const vw = r.width, vh = r.height;
              ctx.clearRect(0, 0, vw, vh);
              if (!img.complete) return;
              ctx.save();
              ctx.translate(tx, ty);
              ctx.scale(scale, scale);
              ctx.drawImage(img, 0, 0);
              ctx.restore();
              badge.textContent = Math.round(scale * 100) + "%";
            }}

            function clamp(v, a, b) {{ return Math.max(a, Math.min(b, v)); }}

            function zoomAt(factor, cx, cy) {{
              const oldScale = scale;
              const next = clamp(scale * factor, minScale, maxScale);
              if (next === oldScale) return;
              const px = (cx - tx) / oldScale;
              const py = (cy - ty) / oldScale;
              scale = next;
              tx = cx - px * scale;
              ty = cy - py * scale;
              draw();
            }}

            img.onload = () => {{
              // default: 1:1 actual size, centered in viewport
              const r = wrap.getBoundingClientRect();
              tx = (r.width - {iw}) / 2;
              ty = (r.height - {ih}) / 2;
              scale = 1.0;
              resizeCanvas();
            }};

            window.addEventListener("resize", resizeCanvas);
            canvas.addEventListener("wheel", (e) => {{
              e.preventDefault();
              const rect = canvas.getBoundingClientRect();
              const cx = e.clientX - rect.left;
              const cy = e.clientY - rect.top;
              zoomAt(e.deltaY < 0 ? 1.1 : 1 / 1.1, cx, cy);
            }}, {{ passive: false }});

            canvas.addEventListener("dblclick", (e) => {{
              const r = wrap.getBoundingClientRect();
              scale = 1.0;
              tx = (r.width - {iw}) / 2;
              ty = (r.height - {ih}) / 2;
              draw();
            }});

            canvas.addEventListener("mousedown", (e) => {{
              dragging = true;
              lx = e.clientX; ly = e.clientY;
              canvas.style.cursor = "grabbing";
            }});
            window.addEventListener("mouseup", () => {{
              dragging = false;
              canvas.style.cursor = "grab";
            }});
            window.addEventListener("mousemove", (e) => {{
              if (!dragging) return;
              tx += e.clientX - lx;
              ty += e.clientY - ly;
              lx = e.clientX; ly = e.clientY;
              draw();
            }});

            function touchDist(t0, t1) {{
              const dx = t0.clientX - t1.clientX;
              const dy = t0.clientY - t1.clientY;
              return Math.hypot(dx, dy);
            }}
            function touchCenter(t0, t1, rect) {{
              return {{
                x: ((t0.clientX + t1.clientX) / 2) - rect.left,
                y: ((t0.clientY + t1.clientY) / 2) - rect.top
              }};
            }}
            canvas.addEventListener("touchstart", (e) => {{
              if (e.touches.length === 1) {{
                dragging = true;
                lx = e.touches[0].clientX; ly = e.touches[0].clientY;
              }} else if (e.touches.length === 2) {{
                const rect = canvas.getBoundingClientRect();
                pinchStartDist = touchDist(e.touches[0], e.touches[1]);
                pinchStartScale = scale;
                pinchCenter = touchCenter(e.touches[0], e.touches[1], rect);
              }}
            }}, {{ passive: true }});
            canvas.addEventListener("touchmove", (e) => {{
              if (e.touches.length === 1 && dragging) {{
                tx += e.touches[0].clientX - lx;
                ty += e.touches[0].clientY - ly;
                lx = e.touches[0].clientX; ly = e.touches[0].clientY;
                draw();
              }} else if (e.touches.length === 2 && pinchStartDist > 0) {{
                const dist = touchDist(e.touches[0], e.touches[1]);
                const ratio = dist / pinchStartDist;
                const target = clamp(pinchStartScale * ratio, minScale, maxScale);
                const factor = target / scale;
                zoomAt(factor, pinchCenter.x, pinchCenter.y);
              }}
            }}, {{ passive: true }});
            canvas.addEventListener("touchend", () => {{
              dragging = false;
              pinchStartDist = 0;
            }});
          }})();
        </script>
        """
        components.html(viewer_html, height=510)
    else:
        st.info("点击「重新生成」后预览图将显示在此处")

    if st.button("确认", type="primary", use_container_width=True, key="dlg_confirm_btn"):
        st.session_state["_dlg_confirmed"] = True

# ── UI ───────────────────────────────────────────────────

st.markdown('<div class="apple-section-title">第一步：上传文件</div>', unsafe_allow_html=True)

upload_col1, upload_col2, upload_col3 = st.columns([1, 1.3, 1])

with upload_col1:
    template_file = st.file_uploader(
        "上传模板文件",
        type=ALL_TEMPLATE_TYPES,
        help="PSD / PSB / PNG / JPG / TIFF / BMP / WebP / PDF / EPS / AI",
    )
    st.markdown(
        '<div class="apple-info-card"><strong>模板文件规则</strong>'
        '<span>支持 PSD、PSB、PNG、JPG、AI、PDF 等格式。建议优先使用原始设计稿，图层识别更准确。</span>'
        '<br><span style="font-size:0.84rem;color:rgba(128,128,132,0.75);">如需修改模板，重新上传即可覆盖当前文件。</span>'
        '</div>',
        unsafe_allow_html=True,
    )

with upload_col2:
    list_file = st.file_uploader("上传名单文件", type=LIST_EXTENSIONS)
    if st.button("手动输入名单", use_container_width=True, key="open_manual_input_dialog_btn"):
        manual_input_dialog()
    manual_rows = st.session_state.get("manual_list_rows", [])
    if manual_rows:
        st.caption(f"已手动输入 {len(manual_rows)} 条名单")
    st.markdown(
        '<div class="apple-info-card"><strong>名单规则（二选一）</strong>'
        '<span>支持 CSV、XLSX、XLS。少量名单可点「手动输入名单」。</span>'
        '</div>',
        unsafe_allow_html=True,
    )

with upload_col3:
    qr_file = st.file_uploader(
        "上传替换二维码（可选）",
        type=["png", "jpg", "jpeg", "webp"],
    )
    st.markdown(
        '<div class="apple-info-card"><strong>二维码规则</strong>'
        '<span>可选项。不上传则保留模板原二维码；上传后按模板位置自动替换并保留圆角效果。</span>'
        '<br><span style="font-size:0.84rem;color:rgba(128,128,132,0.75);">如需更换二维码，重新上传图片即可覆盖。</span>'
        '</div>',
        unsafe_allow_html=True,
    )



has_list = list_file or manual_rows

def _file_fingerprint():
    parts = []
    if template_file:
        parts.append(f"t:{template_file.name}:{template_file.size}")
    if list_file:
        parts.append(f"l:{list_file.name}:{list_file.size}")
    if manual_rows:
        parts.append(f"m:{len(manual_rows)}")
    if qr_file:
        parts.append(f"q:{qr_file.name}:{qr_file.size}")
    return "|".join(parts)

_cur_fp = _file_fingerprint()
_prev_fp = st.session_state.get("_file_fingerprint")
if _cur_fp != _prev_fp:
    st.session_state["_file_fingerprint"] = _cur_fp
    for _k in ["all_img_data", "preview_imgs", "check_done",
               "check_issues", "list_issues", "fix_log_text"]:
        st.session_state[_k] = None

if template_file and has_list:
    suffix = file_suffix(template_file)
    is_psd = suffix in PSD_EXTENSIONS

    _tpl_key = _template_cache_key(template_file)
    _cached_tpl = st.session_state.get("_tpl_cache_key")
    _need_parse = (_cached_tpl != _tpl_key) or ("_tpl_data" not in st.session_state)

    if _need_parse:
        with st.spinner("正在解析模板..."):
            if is_psd:
                psd = load_psd(template_file)
                text_layers = get_text_layers(psd)
                layer_names = [l.name for l in text_layers]
                _default_font = get_default_font_path()
                positions = get_text_layer_positions(psd, _default_font)
                font_color = get_font_color(psd)
                font_size = 51
                qr_box = detect_qr_region(psd)
                qr_layer = get_qr_layer(psd)
                qr_layer_img = qr_layer.composite() if qr_layer else None
                _unused_mask = extract_qr_mask(psd)
                original_img = psd.composite()
                bg = composite_background(psd)
                img_width = psd.width
                img_height = psd.height
                positions = calibrate_stroke_weights(
                    psd, positions, original_img, bg,
                    _default_font, font_color, img_width)
                st.session_state["_tpl_data"] = {
                    "psd": psd, "text_layers": text_layers,
                    "layer_names": layer_names, "positions": positions,
                    "font_color": font_color, "font_size": font_size,
                    "qr_box": qr_box, "qr_layer_img": qr_layer_img,
                    "original_img": original_img, "bg": bg,
                    "img_width": img_width, "img_height": img_height,
                    "is_psd": True,
                }
            else:
                loaded = load_image(template_file)
                if loaded is None:
                    st.stop()
                original_img = loaded.copy()
                bg = loaded
                st.session_state["_tpl_data"] = {
                    "original_img": original_img, "bg": bg,
                    "img_width": bg.size[0], "img_height": bg.size[1],
                    "is_psd": False, "psd": None, "text_layers": [],
                    "layer_names": [], "positions": {},
                    "font_color": (255, 255, 255, 255), "font_size": 51,
                    "qr_box": None, "qr_layer_img": None,
                }
            st.session_state["_tpl_cache_key"] = _tpl_key
    else:
        _td = st.session_state["_tpl_data"]
        is_psd = _td["is_psd"]

    _td = st.session_state["_tpl_data"]
    if is_psd:
        psd = _td["psd"]
        text_layers = _td["text_layers"]
        layer_names = _td["layer_names"]
        positions = _td["positions"]
        font_color = _td["font_color"]
        font_size = _td["font_size"]
        qr_box = _td["qr_box"]
        qr_layer_img = _td["qr_layer_img"]
        original_img = _td["original_img"]
        bg = _td["bg"].copy()
        img_width = _td["img_width"]
        img_height = _td["img_height"]
    else:
        original_img = _td["original_img"]
        bg = _td["bg"].copy()
        img_width = _td["img_width"]
        img_height = _td["img_height"]
        font_size = _td["font_size"]
        font_color = _td["font_color"]
        qr_box = _td["qr_box"]
        qr_layer_img = _td["qr_layer_img"]
        layer_names = _td["layer_names"]
        positions = _td["positions"]

    if list_file:
        rows, fields = parse_spreadsheet(list_file)
    else:
        rows = manual_rows
        fields = list(rows[0].keys()) if rows else []
    if not rows:
        st.warning("名单为空或读取失败，请检查文件。")
        st.stop()

    info_parts = []
    if is_psd:
        info_parts.append(f"检测到 {len(text_layers)} 个文字图层")
    else:
        info_parts.append(f"模板尺寸 {img_width}x{img_height}")
    info_parts.append(f"名单共 {len(rows)} 条记录")
    if qr_box:
        info_parts.append("已识别二维码区域")
    st.success(f"解析完成：{'，'.join(info_parts)}")
    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric("名单记录数", len(rows))
    with m2:
        st.metric("模板宽度", img_width)
    with m3:
        if qr_box:
            st.metric("二维码状态", "已检测")
        elif qr_file:
            st.metric("二维码状态", "未检测到")
            st.warning("已上传替换二维码但模板中未检测到二维码区域。请确认 PSD 图层名称包含「二维码」「QR」「扫码」等关键词。")

    # ── auto-detect fields ──
    COMPANY_KEYWORDS = ["company", "\u516c\u53f8", "\u5355\u4f4d", "\u673a\u6784", "\u4f01\u4e1a", "\u7ec4\u7ec7"]
    NAME_KEYWORDS = ["name", "\u59d3\u540d", "\u4eba\u540d", "\u79f0\u547c", "\u540d\u5b57"]

    def auto_detect_field(fields, keywords):
        for i, f in enumerate(fields):
            for kw in keywords:
                if kw in f.lower():
                    return i
        return None

    def auto_detect_layer(layer_names, keywords):
        for i, name in enumerate(layer_names):
            for kw in keywords:
                if kw in name.lower():
                    return i
        return None

    company_idx = auto_detect_field(fields, COMPANY_KEYWORDS)
    name_idx = auto_detect_field(fields, NAME_KEYWORDS)
    has_company_field = company_idx is not None
    has_name_field = name_idx is not None

    psd_company_layer_idx = auto_detect_layer(layer_names, COMPANY_KEYWORDS) if is_psd else None
    psd_name_layer_idx = auto_detect_layer(layer_names, NAME_KEYWORDS) if is_psd else None

    # ── field mapping UI ──
    st.markdown('<div class="apple-section-title">第二步：字段映射与样式设置</div>', unsafe_allow_html=True)

    if has_company_field and has_name_field:
        st.info(f"\u5df2\u81ea\u52a8\u8bc6\u522b: \u516c\u53f8\u540d=\u300c{fields[company_idx]}\u300d, \u4eba\u540d=\u300c{fields[name_idx]}\u300d")
    elif has_name_field and not has_company_field:
        st.info(f"\u5df2\u81ea\u52a8\u8bc6\u522b: \u4eba\u540d=\u300c{fields[name_idx]}\u300d (\u672a\u68c0\u6d4b\u5230\u516c\u53f8\u540d\u5b57\u6bb5)")
    elif has_company_field and not has_name_field:
        st.info(f"\u5df2\u81ea\u52a8\u8bc6\u522b: \u516c\u53f8\u540d=\u300c{fields[company_idx]}\u300d (\u672a\u68c0\u6d4b\u5230\u4eba\u540d\u5b57\u6bb5)")

    st.caption("\u6839\u636e\u4f60\u7684\u540d\u5355\u5185\u5bb9\u52fe\u9009\u9700\u8981\u66ff\u6362\u7684\u5b57\u6bb5\u3002\u672a\u52fe\u9009\u7684\u5b57\u6bb5\u4e0d\u4f1a\u66ff\u6362\uff0c\u4fdd\u6301\u539f\u6a21\u677f\u5185\u5bb9\u3002")
    enable_company = st.checkbox("\u542f\u7528\u516c\u53f8\u540d", value=has_company_field,
                                 help="\u52fe\u9009\u540e\u4f1a\u7528\u540d\u5355\u4e2d\u7684\u516c\u53f8\u540d\u66ff\u6362\u6a21\u677f\u4e2d\u7684\u516c\u53f8\u540d\u4f4d\u7f6e")
    enable_name = st.checkbox("\u542f\u7528\u4eba\u540d", value=has_name_field,
                              help="\u52fe\u9009\u540e\u4f1a\u7528\u540d\u5355\u4e2d\u7684\u4eba\u540d\u66ff\u6362\u6a21\u677f\u4e2d\u7684\u4eba\u540d\u4f4d\u7f6e")

    if not enable_company and not enable_name:
        st.warning("\u81f3\u5c11\u9700\u8981\u542f\u7528\u4e00\u4e2a\u5b57\u6bb5")
        st.stop()

    company_field = None
    company_y = 0
    company_fsize = font_size
    company_stroke = 0
    company_align = "center"
    name_field = None
    name_y = 0
    name_fsize = font_size
    name_stroke = 0
    name_align = "center"
    company_font_path = None
    name_font_path = None
    company_color = None
    name_color = None
    company_layer = None
    name_layer = None
    company_font_hint = None
    name_font_hint = None

    if enable_company:
        mcol1, mcol2 = st.columns(2) if enable_name else [st.container(), None]
        with mcol1:
            company_field = st.selectbox(
                "\u516c\u53f8\u540d\u5bf9\u5e94\u5b57\u6bb5", fields,
                index=company_idx if company_idx is not None else 0)
            if is_psd and layer_names:
                cl_idx = psd_company_layer_idx if psd_company_layer_idx is not None else max(0, len(layer_names) - 1)
                company_layer = st.selectbox("\u516c\u53f8\u540d\u5bf9\u5e94 PSD \u56fe\u5c42", layer_names, index=cl_idx)
                company_y = positions[company_layer][0]
                company_fsize = positions[company_layer][1]
                company_stroke = positions[company_layer][3]
                company_align = positions[company_layer][5] if len(positions[company_layer]) > 5 else "center"
                company_font_hint = st.empty()
            else:
                company_y = st.number_input("\u516c\u53f8\u540d Y \u5750\u6807", 0, img_height, int(img_height * 0.45))

    if enable_name:
        col_ctx = mcol2 if (enable_company and mcol2 is not None) else st.container()
        with col_ctx:
            name_field = st.selectbox(
                "\u4eba\u540d\u5bf9\u5e94\u5b57\u6bb5", fields,
                index=name_idx if name_idx is not None else min(1, len(fields) - 1))
            if is_psd and layer_names:
                nl_idx = psd_name_layer_idx if psd_name_layer_idx is not None else 0
                name_layer = st.selectbox("\u4eba\u540d\u5bf9\u5e94 PSD \u56fe\u5c42", layer_names, index=nl_idx)
                name_y = positions[name_layer][0]
                name_fsize = positions[name_layer][1]
                name_stroke = positions[name_layer][3]
                name_align = positions[name_layer][5] if len(positions[name_layer]) > 5 else "center"
                name_font_hint = st.empty()
            else:
                name_y = st.number_input("\u4eba\u540d Y \u5750\u6807", 0, img_height, int(img_height * 0.48))

    # ── error detection ──
    mapping_ok = True
    if enable_company and enable_name:
        if company_field == name_field:
            st.error("\u516c\u53f8\u540d\u548c\u4eba\u540d\u9009\u4e86\u540c\u4e00\u4e2a\u5b57\u6bb5\uff0c\u8bf7\u4fee\u6539!")
            mapping_ok = False
        if is_psd and company_layer and name_layer and company_layer == name_layer:
            st.error("\u516c\u53f8\u540d\u548c\u4eba\u540d\u9009\u4e86\u540c\u4e00\u4e2a PSD \u56fe\u5c42\uff0c\u8bf7\u4fee\u6539!")
            mapping_ok = False

    if mapping_ok and enable_name and name_field:
        sample_names = [rows[i][name_field] for i in range(min(3, len(rows)))]
        long_names = [n for n in sample_names if len(n) > 4]
        if long_names:
            st.warning(f"\u4eba\u540d\u5b57\u6bb5\u4e2d\u53d1\u73b0\u8f83\u957f\u7684\u503c: \u300c{'、'.join(long_names)}\u300d\uff0c\u8bf7\u786e\u8ba4\u662f\u5426\u9009\u5bf9\u4e86\u5b57\u6bb5")

    # ── editable name preview ──
    if mapping_ok:
        preview_n = min(180, len(rows))
        st.markdown(f"**名单预览（前 {preview_n} 名）：**")
        preview_data = []
        col_map = {}
        for i in range(preview_n):
            row_preview = {}
            if enable_company and company_field:
                row_preview["公司名"] = rows[i].get(company_field, "")
                col_map["公司名"] = company_field
            if enable_name and name_field:
                row_preview["人名"] = rows[i].get(name_field, "")
                col_map["人名"] = name_field
            preview_data.append(row_preview)
        if preview_data:
            preview_df = pd.DataFrame(preview_data)
            edited_df = st.data_editor(
                preview_df, num_rows="fixed",
                use_container_width=True, key="name_editor",
            )
            for idx, row_edit in edited_df.iterrows():
                if idx < len(rows):
                    for col_label, field_key in col_map.items():
                        new_val = str(row_edit.get(col_label, ""))
                        if new_val != "nan":
                            rows[idx][field_key] = new_val


    if not mapping_ok:
        st.stop()

    # ── font selection (system overhaul) ──
    st.markdown("### \u5b57\u4f53\u9009\u62e9")

    all_fonts, _ps_name_cache = scan_fonts()
    uploaded_entries = []
    for it in st.session_state.get("_uploaded_font_entries", []):
        p = str(it.get("path", ""))
        d = str(it.get("display", ""))
        if p and d and os.path.exists(p):
            uploaded_entries.append({"display": d, "path": p})
    st.session_state["_uploaded_font_entries"] = uploaded_entries
    uploaded_fonts = {it["display"]: it["path"] for it in uploaded_entries}
    merged_fonts = dict(all_fonts)
    merged_fonts.update(uploaded_fonts)
    merged_font_names = list(merged_fonts.keys())
    _fidx = build_local_font_index(merged_fonts, _ps_name_cache)

    psd_candidates = []
    psd_matched = []
    psd_unmatched = []
    psd_recommended = None
    layer_font_map = {}
    per_layer_fonts = {}
    per_layer_colors = {}
    if is_psd:
        psd_candidates = extract_psd_font_candidates(psd)
        psd_matched, psd_unmatched, psd_recommended = match_psd_fonts_to_local(psd_candidates, _fidx)
        layer_font_map = extract_per_layer_font(psd)
        per_layer_colors = extract_per_layer_color(psd)
        if psd_candidates:
            with st.expander("\u6a21\u677f\u5185\u6807\u6ce8\u5b57\u4f53\uff08PSD \u7cbe\u786e\u8bc6\u522b\uff09", expanded=False):
                if psd_matched:
                    st.caption("\u5df2\u5339\u914d\u5230\u5b57\u4f53\uff1a")
                    for _m in psd_matched[:20]:
                        st.text(f"  \u2022 {_m['psd'].get('raw') or _m['psd'].get('family')} \u2192 {_m['local']['display']}")
                if psd_unmatched:
                    st.caption("\u672a\u5339\u914d\u5b57\u4f53\uff1a")
                    for _u in psd_unmatched[:20]:
                        st.text(f"  \u2022 {_u.get('raw') or _u.get('family')}")

    _pref_path = st.session_state.get("_preferred_font_path")
    _pref_label = st.session_state.get("_preferred_font_label", "")
    _has_pref = bool(_pref_path and os.path.exists(str(_pref_path)))
    _default_src = 2 if _has_pref else (1 if (is_psd and psd_recommended) else 0)
    font_source = st.radio(
        "\u5b57\u4f53\u6765\u6e90",
        ["\u9ed8\u8ba4\u5b57\u4f53", "\u672c\u673a\u5b57\u4f53", "\u4e0a\u4f20\u5b57\u4f53"],
        index=_default_src,
        horizontal=True,
        key="font_source_radio",
    )

    # fallback default font = OPPO 4.0 first
    _oppo_medium = FONTS_DIR / "OPPOSans-Medium.ttf"
    _oppo4 = FONTS_DIR / "OPPOSans4.ttf"
    _fallback_font_path = str(_oppo4) if _oppo4.exists() else (str(_oppo_medium) if _oppo_medium.exists() else get_default_font_path())
    _fallback_font_label = "OPPO Sans 4.0 (Regular)" if _oppo4.exists() else ("OPPO Sans (Medium)" if _oppo_medium.exists() else Path(_fallback_font_path).name)

    # upload area (for upload mode and reuse by other modes)
    uploaded_files = []
    if font_source == "\u4e0a\u4f20\u5b57\u4f53":
        uploaded_files = st.file_uploader(
            "\u4e0a\u4f20\u5b57\u4f53\u6587\u4ef6\uff08\u53ef\u591a\u9009 .ttf/.otf\uff09",
            type=["ttf", "otf"],
            accept_multiple_files=True,
            key="multi_font_uploader",
        ) or []
        if uploaded_files:
            _rt_dir = APP_DIR / ".runtime" / "uploaded-fonts"
            _rt_dir.mkdir(parents=True, exist_ok=True)
            for f in uploaded_files:
                _dst = _rt_dir / f"{Path(f.name).stem}{file_suffix(f)}"
                _dst.write_bytes(f.getvalue())
                try:
                    fam, sty = ImageFont.truetype(str(_dst), 20).getname()
                    _disp = f"{fam} ({sty})"
                except Exception:
                    _disp = _dst.stem
                if _disp not in merged_fonts:
                    merged_fonts[_disp] = str(_dst)
                push_font_history(_disp, str(_dst))
            # refresh merged after upload
            merged_font_names = list(merged_fonts.keys())
            _fidx = build_local_font_index(merged_fonts, _ps_name_cache)
            st.session_state["_uploaded_font_entries"] = [
                {"display": d, "path": p}
                for d, p in merged_fonts.items()
                if str(p).startswith(str(APP_DIR / ".runtime" / "uploaded-fonts")) and os.path.exists(p)
            ]

    # font history
    _hist = st.session_state.get("_font_history", [])
    _hist = [h for h in _hist if isinstance(h, dict) and os.path.exists(str(h.get("path", "")))]
    st.session_state["_font_history"] = _hist[:20]

    def _show_layer_font_detail(plf, lfm, fallback_label):
        """Display per-layer font assignment detail."""
        if not plf or not lfm:
            return
        st.markdown("\u6e90\u6587\u4ef6\u5b57\u4f53\u8bc6\u522b\u7ed3\u679c\uff1a")
        for lname, psname in lfm.items():
            fpath = plf.get(lname, "")
            fname = Path(fpath).stem if fpath else fallback_label
            is_fallback = (fpath == _fallback_font_path) or not fpath
            if is_fallback:
                st.warning(f"\u56fe\u5c42\u300c{lname}\u300d: {psname} -> {fname}\uff08\u672a\u5339\u914d\uff0c\u5df2\u56de\u9000\u9ed8\u8ba4\u5b57\u4f53\uff09")
            else:
                st.info(f"\u56fe\u5c42\u300c{lname}\u300d: {psname} -> {fname}\uff08\u5339\u914d\u6210\u529f\uff09")

    if font_source == "\u9ed8\u8ba4\u5b57\u4f53":
        if is_psd:
            per_layer_fonts = resolve_per_layer_font_path(layer_font_map, _fidx, _fallback_font_path)
            font_path = _fallback_font_path
            _missing = []
            for _lname, _psname in layer_font_map.items():
                if per_layer_fonts.get(_lname) == _fallback_font_path:
                    _missing.append((_lname, _psname))
            st.success("\u9ed8\u8ba4\u5b57\u4f53\u5df2\u81ea\u52a8\u6309 PSD \u56fe\u5c42\u5339\u914d")
            _show_layer_font_detail(per_layer_fonts, layer_font_map, _fallback_font_label)
            if _missing:
                st.warning("\u5b58\u5728\u6a21\u677f\u5b57\u4f53\u672a\u5728\u5f53\u524d\u73af\u5883\u627e\u5230\uff0c\u53ef\u5207\u6362\u5230\u300c\u4e0a\u4f20\u5b57\u4f53\u300d\u4e0a\u4f20\u5bf9\u5e94\u5b57\u4f53\u6587\u4ef6\u4ee5\u7cbe\u51c6\u5339\u914d\u3002")
        else:
            font_path = _fallback_font_path
            st.success(f"\u5df2\u4f7f\u7528\u9ed8\u8ba4\u5b57\u4f53 {_fallback_font_label}")
        push_font_history(_fallback_font_label, font_path)
        if _hist:
            hist_labels = [f"{i+1}. {h.get('display','')}" for i, h in enumerate(_hist)]
            _hidx = st.selectbox("\u8fc7\u5f80\u5b57\u4f53\u5217\u8868\uff08\u53ef\u8986\u76d6\u5f53\u524d\u9ed8\u8ba4\uff09", hist_labels, index=0)
            _hi = int(_hidx.split(".", 1)[0]) - 1
            if 0 <= _hi < len(_hist):
                font_path = _hist[_hi]["path"]
                st.caption(f"\u5df2\u4ece\u5386\u53f2\u5217\u8868\u9009\u62e9\uff1a{_hist[_hi]['display']}")
    elif font_source == "\u672c\u673a\u5b57\u4f53":
        local_names = list(all_fonts.keys())
        if local_names:
            _font_query = st.text_input("\u641c\u7d22\u5b57\u4f53\uff08\u4e2d/\u82f1\u6587\uff09", "", key="font_search_input",
                                        placeholder="\u8f93\u5165\u5173\u952e\u8bcd\u8fc7\u6ee4\uff0c\u5982 PingFang\u3001\u82f9\u65b9\u3001OPPO\u3001\u5b8b\u4f53...")
            if _font_query.strip():
                _q = _font_query.strip().lower()
                filtered_names = [n for n in local_names if _q in n.lower() or _q in Path(all_fonts[n]).stem.lower()]
            else:
                filtered_names = local_names
            if not filtered_names:
                st.caption(f"\u672a\u627e\u5230\u5305\u542b\u300c{_font_query}\u300d\u7684\u5b57\u4f53\uff0c\u663e\u793a\u5168\u90e8 {len(local_names)} \u4e2a")
                filtered_names = local_names
            default_idx = 0
            if is_psd and psd_recommended and psd_recommended in filtered_names:
                default_idx = filtered_names.index(psd_recommended)
            elif "OPPO Sans 4.0 (Regular)" in filtered_names:
                default_idx = filtered_names.index("OPPO Sans 4.0 (Regular)")
            selected_font = st.selectbox(
                "\u672c\u673a\u5b57\u4f53\u9009\u62e9",
                filtered_names,
                index=default_idx,
                help=f"\u5df2\u626b\u63cf\u5230 {len(local_names)} \u4e2a\u672c\u673a\u5b57\u4f53\uff08\u5f53\u524d\u663e\u793a {len(filtered_names)} \u4e2a\uff09",
            )
            font_path = all_fonts[selected_font]
            if is_psd:
                per_layer_fonts = {ln: font_path for ln in layer_font_map.keys()}
            push_font_history(selected_font, font_path)
        else:
            font_path = _fallback_font_path
            st.warning("\u672a\u626b\u63cf\u5230\u672c\u673a\u5b57\u4f53\uff0c\u5df2\u56de\u9000\u5230\u9ed8\u8ba4\u5b57\u4f53")
    else:
        uploaded_only = dict(uploaded_fonts)
        # Include freshly uploaded fonts in current turn
        for d, p in merged_fonts.items():
            if str(p).startswith(str(APP_DIR / ".runtime" / "uploaded-fonts")):
                uploaded_only[d] = p
        if not uploaded_only:
            st.info("\u8bf7\u4e0a\u4f20\u4e00\u4e2a\u6216\u591a\u4e2a .ttf/.otf \u5b57\u4f53\u6587\u4ef6")
            font_path = _fallback_font_path
        else:
            up_names = list(uploaded_only.keys())
            if is_psd:
                _up_idx = build_local_font_index(uploaded_only, _ps_name_cache)
                per_layer_fonts = resolve_per_layer_font_path(layer_font_map, _up_idx, _fallback_font_path)
                font_path = _fallback_font_path
                _all_fallback = True
                for _ln in layer_font_map.keys():
                    if per_layer_fonts.get(_ln) and per_layer_fonts.get(_ln) != _fallback_font_path:
                        _all_fallback = False
                        break
                # If user uploaded a single font but nothing matched, force apply globally.
                if _all_fallback and len(uploaded_only) == 1:
                    _single_disp = up_names[0]
                    _single_path = uploaded_only[_single_disp]
                    per_layer_fonts = {ln: _single_path for ln in layer_font_map.keys()}
                    font_path = _single_path
                    st.success("\u5df2\u4e0a\u4f20 1 \u4e2a\u5b57\u4f53\uff0c\u6a21\u677f\u6e90\u5b57\u4f53\u540d\u4e0e\u6587\u4ef6\u540d\u4e0d\u4e00\u81f4\uff0c\u5df2\u76f4\u63a5\u5168\u5c40\u5e94\u7528\u8be5\u5b57\u4f53\u751f\u6210\u3002")
                else:
                    st.success("\u5df2\u6309 PSD \u56fe\u5c42\u5339\u914d\u4e0a\u4f20\u5b57\u4f53")
                _show_layer_font_detail(per_layer_fonts, layer_font_map, _fallback_font_label)
                _missing2 = []
                for _lname, _psname in layer_font_map.items():
                    if per_layer_fonts.get(_lname) == _fallback_font_path:
                        _missing2.append((_lname, _psname))
                if _missing2 and not (_all_fallback and len(uploaded_only) == 1):
                    st.warning("\u90e8\u5206\u56fe\u5c42\u5b57\u4f53\u4ecd\u672a\u5339\u914d\uff0c\u5982\u9700 1:1 \u8fd8\u539f\uff0c\u8bf7\u7ee7\u7eed\u4e0a\u4f20\u7f3a\u5931\u5b57\u4f53\u6587\u4ef6\u3002")
                _upload_map = {}
                for _lname, _fpath in (per_layer_fonts or {}).items():
                    _font_key = Path(_fpath).stem if _fpath else _fallback_font_label
                    _upload_map.setdefault(_font_key, []).append(_lname)
                if _upload_map:
                    st.markdown("\u4e0a\u4f20\u5b57\u4f53\u5e94\u7528\u7ed3\u679c\uff1a")
                    for _font_name, _layers in _upload_map.items():
                        _layers_txt = "\u3001".join(_layers)
                        st.info(f"{_font_name} -> \u56fe\u5c42\uff1a{_layers_txt}")
                for d, p in uploaded_only.items():
                    push_font_history(d, p)
            else:
                _pick = st.selectbox("\u4e0a\u4f20\u5b57\u4f53\u9009\u62e9", up_names, index=0)
                font_path = uploaded_only[_pick]
                push_font_history(_pick, font_path)
                st.success(f"\u5df2\u4f7f\u7528\u4e0a\u4f20\u5b57\u4f53\uff1a{_pick}")

    # keep preferred keys for cross-rerun
    st.session_state["_preferred_font_path"] = font_path
    st.session_state["_preferred_font_label"] = Path(font_path).name if font_path else ""
    if is_psd and not per_layer_fonts:
        per_layer_fonts = resolve_per_layer_font_path(layer_font_map, _fidx, font_path)

    # ── re-calibrate positions for the chosen font ──
    if is_psd and font_path:
        positions = get_text_layer_positions(psd, font_path, per_layer_fonts=per_layer_fonts)
        positions = calibrate_stroke_weights(
            psd, positions, original_img, bg, font_path, font_color, img_width)
        if enable_company and company_layer and company_layer in positions:
            company_y = positions[company_layer][0]
            company_fsize = positions[company_layer][1]
            company_stroke = positions[company_layer][3]
            company_align = positions[company_layer][5] if len(positions[company_layer]) > 5 else "center"
            company_font_path = per_layer_fonts.get(company_layer, font_path)
            company_color = per_layer_colors.get(company_layer, font_color)
            if company_font_hint is not None:
                company_font_hint.caption(f"\u2192 \u5c06\u4f7f\u7528\u5b57\u4f53\uff1a{Path(company_font_path).stem}")
        if enable_name and name_layer and name_layer in positions:
            name_y = positions[name_layer][0]
            name_fsize = positions[name_layer][1]
            name_stroke = positions[name_layer][3]
            name_align = positions[name_layer][5] if len(positions[name_layer]) > 5 else "center"
            name_font_path = per_layer_fonts.get(name_layer, font_path)
            name_color = per_layer_colors.get(name_layer, font_color)
            if name_font_hint is not None:
                name_font_hint.caption(f"\u2192 \u5c06\u4f7f\u7528\u5b57\u4f53\uff1a{Path(name_font_path).stem}")

    # ── font weight ──
    auto_stroke = max(company_stroke, name_stroke)
    stroke_override = st.number_input(
        "字体粗细微调",
        min_value=0.0, max_value=10.0, value=float(auto_stroke), step=0.001, format="%.3f",
        help="像素校准默认值为 {}，数值越大越粗。Pillow 渲染时取整数部分，小数部分通过透明度模拟。".format(auto_stroke),
    )
    company_stroke = int(round(stroke_override))
    name_stroke = int(round(stroke_override))

    if not is_psd:
        fcol1, fcol2 = st.columns(2)
        with fcol1:
            font_size = st.number_input("\u5b57\u53f7", 10, 200, int(font_size))
        with fcol2:
            color_hex = st.color_picker("\u6587\u5b57\u989c\u8272", "#FFFFFF")
            r, g, b = int(color_hex[1:3], 16), int(color_hex[3:5], 16), int(color_hex[5:7], 16)
            font_color = (r, g, b, 255)

    # ── QR replacement ──
    if qr_file and not qr_box:
        if _HAS_CV2:
            try:
                _detect_arr = np.array(original_img.convert("RGB"))
                _detect_gray = cv2.cvtColor(_detect_arr, cv2.COLOR_RGB2GRAY)
                _det = cv2.QRCodeDetector()
                _, _pts, _ = _det.detectAndDecode(_detect_gray)
                if _pts is not None and len(_pts) > 0:
                    _xs, _ys = _pts[0][:, 0], _pts[0][:, 1]
                    _pad = 20
                    qr_box = (
                        max(0, int(_xs.min()) - _pad),
                        max(0, int(_ys.min()) - _pad),
                        min(img_width, int(_xs.max()) + _pad),
                        min(img_height, int(_ys.max()) + _pad),
                    )
                    st.info(f"通过图像分析自动检测到二维码区域 ({qr_box[0]},{qr_box[1]})-({qr_box[2]},{qr_box[3]})")
            except Exception:
                pass
        if not qr_box:
            st.warning("未能检测到模板中的二维码区域，无法自动替换。请确认 PSD 图层名称包含「二维码」「QR」等关键词，或确认模板图片中包含可识别的二维码。")
    if qr_file and qr_box:
        qr_file.seek(0)
        qr_raw = qr_file.read()
        qr_img = Image.open(io.BytesIO(qr_raw)).convert("RGBA")
        tw = qr_box[2] - qr_box[0]
        th = qr_box[3] - qr_box[1]
        qr_resized = qr_img.resize((tw, th), Image.LANCZOS).convert("RGBA")

        corner_radius = 0
        try:
            ori_crop = np.array(original_img.crop(qr_box).convert("RGB"))
            center_brightness = float(ori_crop[th // 2, tw // 2].mean())
            bg_brightness = float(ori_crop[0, 0].mean())
            is_light_qr = center_brightness > bg_brightness + 30
            if is_light_qr:
                gray = ori_crop.mean(axis=2)
                threshold = (center_brightness + bg_brightness) / 2
                top_row = gray[0, :]
                for px in range(min(tw // 4, 60)):
                    if top_row[px] > threshold:
                        corner_radius = max(px, 4)
                        break
                if corner_radius == 0:
                    left_col = gray[:, 0]
                    for py in range(min(th // 4, 60)):
                        if left_col[py] > threshold:
                            corner_radius = max(py, 4)
                            break
        except Exception:
            corner_radius = max(4, min(tw, th) // 15)

        if corner_radius > 0:
            mask = Image.new("L", (tw, th), 0)
            d = ImageDraw.Draw(mask)
            d.rounded_rectangle([0, 0, tw - 1, th - 1], radius=corner_radius, fill=255)
            r, g, b, a = qr_resized.split()
            merged_alpha = Image.fromarray(
                np.minimum(np.array(a), np.array(mask)).astype(np.uint8)
            )
            qr_resized = Image.merge("RGBA", (r, g, b, merged_alpha))

        bg.paste(qr_resized, (qr_box[0], qr_box[1]), qr_resized)
        st.success(f"二维码已替换（区域 {tw}x{th}px，圆角 {corner_radius}px）")

    def build_text_items(row):
        items = []
        if enable_company and company_field:
            items.append((row[company_field], company_y, company_fsize, company_stroke, company_font_path or font_path, company_color or font_color, company_align))
        if enable_name and name_field:
            items.append((row[name_field], name_y, name_fsize, name_stroke, name_font_path or font_path, name_color or font_color, name_align))
        return items

    def build_filename(row):
        parts = []
        if enable_company and company_field:
            parts.append(row[company_field])
        if enable_name and name_field:
            parts.append(row[name_field])
        return "_".join(parts) if parts else f"row"

    # ── preview: original vs first (always freshly generated) ──
    st.markdown('<div class="apple-section-title">第三步：预览与生成</div>', unsafe_allow_html=True)
    first = rows[0]
    preview = generate_one(bg, build_text_items(first), img_width, font_color, font_path)

    pcol1, pcol2 = st.columns(2)
    with pcol1:
        st.image(original_img, caption="原始模板", use_container_width=True)
    with pcol2:
        st.image(preview, caption=f"替换效果: {build_filename(first)}",
                 use_container_width=True)

    prev_act1, prev_act2 = st.columns(2)
    with prev_act1:
        if st.button("🔄 刷新预览", use_container_width=True, key="btn_refresh_preview"):
            st.session_state["_preview_refresh"] = True
    with prev_act2:
        if st.button("输入问题", use_container_width=True, key="btn_open_issue_dialog"):
            st.session_state["_dlg_preview_img"] = preview
            preview_issue_dialog()

    st.caption("预览会在字体、粗细、颜色等参数变化时自动刷新。请仔细对比两张图，确认字体大小、位置、间距、二维码是否与原图一致。")




    # ── generate all / check / download ──
    st.markdown("---")
    st.markdown("### 生成全部")
    total = len(rows)

    st.markdown(f"共 **{total}** 条名单，点击下方按钮一键生成并打包下载。")

    gen_all_col1, gen_all_col2 = st.columns(2)
    with gen_all_col1:
        if st.button(f"生成全部 {total} 张", type="primary", use_container_width=True, key="btn_gen_all"):
            progress2 = st.progress(0, text="正在生成...")
            all_img_data = []
            for i, row in enumerate(rows):
                fname = build_filename(row)
                img = generate_one(bg, build_text_items(row), img_width, font_color, font_path)
                img_buf = io.BytesIO()
                img.save(img_buf, format="PNG")
                all_img_data.append((f"{fname}.png", img_buf.getvalue()))
                progress2.progress((i + 1) / total, text=f"正在生成 [{i+1}/{total}] {fname}")
            progress2.progress(1.0, text=f"全部生成完成! 共 {total} 张")
            st.session_state["all_img_data"] = all_img_data
            st.session_state["check_done"] = None

    with gen_all_col2:
        preview_count = st.radio(
            "可选预览",
            [0, 5, 10],
            horizontal=True,
            format_func=lambda x: "不预览" if x == 0 else f"预览前{x}张",
            key="preview_count_radio",
        )

    if preview_count and preview_count > 0:
        _pcount = min(preview_count, total)
        prev_btn_col, prev_close_col = st.columns(2)
        with prev_btn_col:
            if st.button("生成预览", use_container_width=True, key="btn_gen_preview"):
                preview_imgs = []
                progress = st.progress(0, text="正在生成预览...")
                for i in range(_pcount):
                    row = rows[i]
                    img = generate_one(bg, build_text_items(row), img_width, font_color, font_path)
                    fname = build_filename(row)
                    preview_imgs.append((img.copy(), fname))
                    progress.progress((i + 1) / _pcount, text=f"预览 [{i+1}/{_pcount}]")
                progress.progress(1.0, text=f"预览完成! 共 {_pcount} 张")
                st.session_state["preview_imgs"] = preview_imgs
        with prev_close_col:
            if st.session_state.get("preview_imgs"):
                if st.button("关闭预览", use_container_width=True, key="btn_close_preview"):
                    st.session_state["preview_imgs"] = None

        stored_preview = st.session_state.get("preview_imgs")
        if stored_preview:
            for i in range(0, len(stored_preview), 3):
                chunk = stored_preview[i:i+3]
                cols = st.columns(len(chunk))
                for col, (pimg, caption) in zip(cols, chunk):
                    with col:
                        st.image(pimg, caption=caption, use_container_width=True)

    # ── download + check + redownload ──
    all_img_data = st.session_state.get("all_img_data")
    if all_img_data:
        st.markdown("---")
        total_gen = len(all_img_data)
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for filename, data in all_img_data:
                zf.writestr(filename, data)

        st.download_button(
            label=f"下载全部 ({total_gen} 张 ZIP)",
            data=zip_buf.getvalue(),
            file_name="邀请函批量生成.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
            key="btn_download_zip",
        )

        check_col, regen_col = st.columns(2)
        with check_col:
            do_check = st.button("一件检查全部", use_container_width=True, key="btn_check_all")
        with regen_col:
            do_regen = st.button("重新下载全部", use_container_width=True, key="btn_regen_all")

        # ── execute quality check (all images, full 14-point suite) ──
        if do_check:
            check_count = len(all_img_data)
            all_check_issues = []
            list_issues = check_list_generation_quality(
                rows, enable_company, company_field,
                enable_name, name_field, build_filename,
            )
            _use_custom = bool(st.session_state.get("_custom_font_active"))
            check_progress = st.progress(0, text="正在检查全部图片...")
            for i in range(check_count):
                fname = all_img_data[i][0]
                chk_img = Image.open(io.BytesIO(all_img_data[i][1])).convert("RGBA")
                row = rows[i] if i < len(rows) else rows[-1]
                text_items = build_text_items(row)
                _qr_ref = None
                if qr_file:
                    qr_file.seek(0)
                    _qr_ref = Image.open(io.BytesIO(qr_file.read())).convert("RGBA")
                issues = check_image_quality(
                    chk_img, text_items, img_width, qr_box, font_path,
                    original_img=original_img,
                    replaced_qr_img=_qr_ref,
                )
                issues += compare_preview_quality(
                    original_img, chk_img, text_items, img_width,
                    qr_box, font_path, use_custom_font=_use_custom,
                )
                all_check_issues.append((fname, issues))
                check_progress.progress((i + 1) / check_count, text=f"检查 [{i+1}/{check_count}] {fname}")
            check_progress.progress(1.0, text=f"全部 {check_count} 张检查完成!")

            st.session_state["check_issues"] = all_check_issues
            st.session_state["list_issues"] = list_issues

            has_errors = False
            has_warnings = False
            info_count = 0

            if list_issues:
                st.markdown("#### 名单信息检查")
                for level, msg in list_issues:
                    if level == "error":
                        st.error(msg)
                        has_errors = True
                    elif level == "warning":
                        st.warning(msg)
                        has_warnings = True
                    elif level == "info":
                        info_count += 1
                    else:
                        st.success(msg)

            for fname, issues in all_check_issues:
                for level, msg in issues:
                    if level == "error":
                        st.error(f"**{fname}**: {msg}")
                        has_errors = True
                    elif level == "warning":
                        st.warning(f"**{fname}**: {msg}")
                        has_warnings = True
                    elif level == "info":
                        info_count += 1

            if info_count > 0:
                st.info(f"有 {info_count} 项提示信息（不影响实际使用质量），详见纠错日志。")

            fix_log = build_fix_log(all_check_issues, list_issues)
            st.download_button(
                "下载纠错日志",
                data=fix_log,
                file_name=f"纠错日志_{_dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                mime="text/plain",
                use_container_width=True,
                key="btn_download_fix_log",
            )

            if not has_errors and not has_warnings:
                st.markdown(
                    '<div class="check-pass-banner">✅ 全部检查通过！字体居中 / 间距正常 / 二维码可识别</div>',
                    unsafe_allow_html=True,
                )
                st.session_state["check_done"] = True
            elif not has_errors:
                st.info("检查完成，有轻微警告但不影响使用，可以下载。")
                st.session_state["check_done"] = True
            else:
                st.error("发现错误，建议点击「重新下载全部」调整后重新生成。")
                st.session_state["check_done"] = False

        elif st.session_state.get("check_done") is True:
            st.markdown(
                '<div class="check-pass-banner">✅ 上次检查已通过，可放心下载</div>',
                unsafe_allow_html=True,
            )
        elif st.session_state.get("check_done") is False:
            st.warning("上次检查发现错误，建议调整后点击「重新下载全部」重新生成。")

        # ── regenerate all ──
        if do_regen:
            regen_progress = st.progress(0, text="正在重新生成...")
            regen_data = []
            for i, row in enumerate(rows):
                fname = build_filename(row)
                regen_img = generate_one(bg, build_text_items(row), img_width, font_color, font_path)
                img_buf = io.BytesIO()
                regen_img.save(img_buf, format="PNG")
                regen_data.append((f"{fname}.png", img_buf.getvalue()))
                regen_progress.progress((i + 1) / total, text=f"重新生成 [{i+1}/{total}] {fname}")
            regen_progress.progress(1.0, text=f"重新生成完成! 共 {total} 张")
            st.session_state["all_img_data"] = regen_data
            st.session_state["check_done"] = None
            st.session_state["check_issues"] = None
            st.session_state["list_issues"] = None

            regen_zip = io.BytesIO()
            with zipfile.ZipFile(regen_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                for filename, data in regen_data:
                    zf.writestr(filename, data)
            st.download_button(
                label=f"重新生成完成，下载全部 ({len(regen_data)} 张 ZIP)",
                data=regen_zip.getvalue(),
                file_name="邀请函批量生成.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
                key="btn_regen_download",
            )

else:
    st.info("\u8bf7\u5148\u4e0a\u4f20\u6a21\u677f\u6587\u4ef6\u548c\u540d\u5355\u6587\u4ef6")

st.markdown("---")
with st.expander("基础问题解说", expanded=False):
    st.markdown(
        "- 若出现 `TypeError: Failed to fetch dynamically imported module`，通常是浏览器缓存或网络拦截。\n"
        "- 请按顺序处理：关闭页面重开 → 强制刷新（Windows: `Ctrl+Shift+R`，macOS: `Cmd+Shift+R`）→ 无痕窗口重试。\n"
        "- 仍无法访问时，请清除 `streamlit.app` 站点数据后再试。\n"
        "- 本工具为 Streamlit Cloud 公网部署，不同网络/在家均可使用；若仅某些网络失败，请联系 IT 放行 `*.streamlit.app` HTTPS 访问。\n"
        "- **建议使用浏览器无痕模式访问，速度更快、不受缓存干扰。**\n"
        '- **如果页面卡住无响应，请点击 <a href="javascript:window.location.reload();" style="color:#0071e3;text-decoration:none;font-weight:600;">刷新</a> ，'
        "页面会自动修复过往出现的所有问题并重新加载。**",
        unsafe_allow_html=True,
    )

st.caption(_BUILD_TAG)
if _heal_log:
    with st.expander(f"自愈日志 ({len(_heal_log)})", expanded=False):
        st.code("\n".join(_heal_log[-30:]), language="text")
        st.download_button(
            "下载自愈日志",
            "\n".join(_heal_log),
            file_name="self_heal_log.txt",
            mime="text/plain",
        )
