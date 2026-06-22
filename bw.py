import os
import re
import sys
import time
import json
import base64
import hashlib
import logging
from datetime import datetime
from urllib.parse import urljoin, urlparse, urlsplit, urlunsplit, parse_qsl, urlencode

import requests
from bs4 import BeautifulSoup

# Optional: load secrets from a local .env file if python-dotenv is installed.
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Optional heavy deps used for Excel export / duplicate tracking.
try:
    import pandas as pd
    import openpyxl
    _XLSX_AVAILABLE = True
except ImportError:
    _XLSX_AVAILABLE = False

# Optional heavy deps used for paraphrase quality gating.
try:
    import language_tool_python
    from sentence_transformers import SentenceTransformer, util as st_util
    _NLP_AVAILABLE = True
except ImportError:
    _NLP_AVAILABLE = False

# =============================================================================
#  CONFIG
# =============================================================================

BASE_URL  = "https://jobsbotswana.info"
SITE_HOST = "jobsbotswana.info"

# WP Job Manager custom-post-type REST endpoint + single-job rewrite slug.
WPJM_REST       = f"{BASE_URL}/wp-json/wp/v2/job-listings"
JOB_SLUG_PREFIX = "/jobs/"

# auto = try REST then fall back to HTML; rest = REST only; html = HTML only.
SOURCE_MODE = os.environ.get("SOURCE_MODE", "auto").strip().lower()

REQUEST_DELAY = float(os.environ.get("REQUEST_DELAY", "1.0"))  # polite delay between requests, seconds
MAX_JOBS = int(os.environ.get("MAX_JOBS", "0"))                # 0 = no cap, otherwise stop after N new jobs

# Cap on source pages crawled (REST pages of 100, or sitemap batches). 0 = all.
_scrape_pages_raw = int(os.environ.get("SCRAPE_PAGES", "0"))
SCRAPE_PAGES = _scrape_pages_raw if _scrape_pages_raw > 0 else None

REST_PER_PAGE = 100

OUTPUT_FILE = "jobsbotswana_jobs.xlsx"
PROCESSED_IDS_FILE = "jobsbotswana_processed.csv"

# ── WordPress (secrets via environment variables — see header docstring) ────
WP_URL      = os.environ.get("WP_BASE_URL", "")
WP_USER     = os.environ.get("WP_USERNAME", "")
WP_PASSWORD = os.environ.get("WP_APP_PASSWORD", "")
WP_BASE      = WP_URL.rstrip("/")
WP_JOBS_URL  = f"{WP_BASE}/job-listings"
WP_MEDIA_URL = f"{WP_BASE}/media"

# ── Mistral (secret via environment variable — see header docstring) ────────
MISTRAL_API_KEY = os.environ.get("MISTRAL_API_KEY", "")
MISTRAL_MODEL   = "mistral-small-latest"
MISTRAL_URL     = "https://api.mistral.ai/v1/chat/completions"

ENABLE_PARAPHRASE = True   # set False to skip paraphrasing entirely

# ── Startup checks: warn (don't crash) if secrets are missing ───────────────
for _var, _val, _feature in [
    ("MISTRAL_API_KEY", MISTRAL_API_KEY, "paraphrasing"),
    ("WP_USERNAME",     WP_USER,         "WordPress posting"),
    ("WP_APP_PASSWORD", WP_PASSWORD,     "WordPress posting"),
]:
    if not _val:
        logging.getLogger(__name__).warning(
            f"Environment variable {_var} is not set — {_feature} will be disabled/skipped."
        )

JOB_TYPE_MAPPING = {
    "full-time": "full-time", "full time": "full-time",
    "part-time": "part-time", "part time": "part-time",
    "contract":  "contract",  "temporary": "temporary",
    "internship":"internship","freelance": "freelance",
    "volunteer": "volunteer",
}

# JSON-LD employmentType (FULL_TIME, ...) -> human label used elsewhere.
EMPLOYMENT_TYPE_HUMAN = {
    "FULL_TIME": "Full Time", "PART_TIME": "Part Time", "CONTRACTOR": "Contract",
    "CONTRACT": "Contract", "TEMPORARY": "Temporary", "INTERN": "Internship",
    "INTERNSHIP": "Internship", "VOLUNTEER": "Volunteer", "PER_DIEM": "Part Time",
    "OTHER": "",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/json",
    "Accept-Charset": "utf-8",
}

REQUEST_TIMEOUT = 25

# Reuse one TCP/TLS connection where possible for every request this run makes.
SESSION = requests.Session()
SESSION.headers.update(HEADERS)

# ── Source landmarks (stable WP-Job-Manager / JSON-LD hooks) ────────────────
SOCIAL_HOST_RE = re.compile(r"(facebook|instagram|twitter|linkedin|youtube|wa\.me|whatsapp|t\.me)", re.I)

# WP Job Manager auto-generates these apply lines when an employer's apply
# method is an email / URL — they carry the application target.
APPLY_TEXT_EMAIL_RE = re.compile(r"to apply for this job email your details to\s+([^\s<]+@[^\s<]+)", re.I)
APPLY_TEXT_URL_RE   = re.compile(r"to apply for this job please visit\s+(https?://[^\s<]+)", re.I)

# UI / template boilerplate that can leak into description text.
BOILERPLATE_PATTERNS = [
    re.compile(r"Apply for job.*$", re.I | re.S),
    re.compile(r"To apply for this job.*$", re.I | re.S),
    re.compile(r"Before applying for this position.*$", re.I | re.S),
    re.compile(r"Login to apply.*$", re.I | re.S),
    re.compile(r"Upload your CV/?resume.*$", re.I | re.S),
    re.compile(r"Related Jobs.*$", re.I | re.S),
    re.compile(r"Email Me Jobs Like These.*$", re.I | re.S),
    re.compile(r"Send to friend.*$", re.I | re.S),
]

# =============================================================================
#  LOGGING / COLOUR
# =============================================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log_ = logging.getLogger(__name__)   # logger instance (.info/.warning/.error)

_USE_COLOUR = sys.stdout.isatty()

def _c(code, text):
    return f"\033[{code}m{text}\033[0m" if _USE_COLOUR else text

C_HEADER  = lambda t: _c("1;36",  t)
C_LABEL   = lambda t: _c("1;33",  t)
C_VALUE   = lambda t: _c("97",    t)
C_DIM     = lambda t: _c("2",     t)
C_GREEN   = lambda t: _c("1;32",  t)
C_RED     = lambda t: _c("1;31",  t)
C_BLUE    = lambda t: _c("1;34",  t)
C_DIVIDER = lambda: _c("2", "─" * 80)

def log(msg):
    """Plain console print (kept distinct from the log_ logger instance)."""
    print(msg, flush=True)

# Matches a plain email address inside free text job descriptions / company details.
EMAIL_PATTERN = re.compile(r"[A-Za-z0-9.+_-]+@[A-Za-z0-9-]+\.[A-Za-z0-9.-]+")

# =============================================================================
#  TEXT CLEANUP / SANITIZATION
# =============================================================================

_MOJIBAKE = [
    ("Â", ""), ("â€™", "'"), ("â€œ", '"'), ("â€\x9d", '"'), ("â€", '"'),
    ("â€¢", "•"), ("â„¢", "™"), ("\u00a0", " "), ("\u200b", ""), ("\ufeff", ""),
]

def _fix_mojibake(text: str) -> str:
    for pattern, replacement in _MOJIBAKE:
        text = text.replace(pattern, replacement)
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)
    return text

def sanitize_text(text, is_url=False) -> str:
    """Light cleanup pass used right before sending a field to WordPress."""
    if not isinstance(text, str):
        text = str(text) if (text is not None and str(text) not in ("nan", "None", "NaN")) else ""
    text = text.strip()
    if text in ("nan", "None", "NaN", "", "N/A", "n/a", "NA", "na"):
        return ""
    text = _fix_mojibake(text)
    if is_url:
        return re.sub(r"[ \t\r\n\f\v]+", " ", text).strip()
    text = re.sub(r"#+\s*", "", text)
    text = re.sub(r"\*\*", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

def clean_text(el):
    if el is None:
        return ""
    return re.sub(r"\s+", " ", el.get_text(" ", strip=True)).strip()

def _clean_description(text):
    """Strip trailing boilerplate while PRESERVING line breaks (so the body
    keeps its paragraph/bullet structure for paraphrasing & WP rendering)."""
    if not text:
        return ""
    for pattern in BOILERPLATE_PATTERNS:
        text = pattern.sub("", text)
    out = []
    for ln in text.split("\n"):
        s = ln.strip()
        if not s:
            continue
        out.append(re.sub(r"[ \t]+", " ", s))
    return "\n".join(out).strip()

def html_to_text(html):
    """Convert rendered HTML (REST content.rendered / JSON-LD description /
    .job_description) into clean multi-line text, preserving paragraphs and
    bullet points, then strip boilerplate."""
    if not html:
        return ""
    soup = BeautifulSoup(html, "lxml")
    parts, seen = [], set()
    blocks = soup.find_all(["h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "blockquote", "tr"])
    if not blocks:
        # plain text with <br>/newlines only
        for br in soup.find_all("br"):
            br.replace_with("\n")
        for ln in soup.get_text("\n").split("\n"):
            s = ln.strip()
            if s and s not in seen:
                seen.add(s)
                parts.append(s)
        return _clean_description("\n".join(parts))
    for el in blocks:
        t = el.get_text(" ", strip=True)
        if not t:
            continue
        line = ("- " + t) if el.name == "li" else t
        key = (el.name, t)
        if key in seen:
            continue
        seen.add(key)
        parts.append(line)
    return _clean_description("\n".join(parts))

# =============================================================================
#  BASIC HTTP / PARSING HELPERS
# =============================================================================

def get_soup(url):
    resp = SESSION.get(url, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    resp.encoding = resp.encoding or "utf-8"
    return BeautifulSoup(resp.text, "lxml")

def absolute_url(href):
    if not href:
        return ""
    if href.startswith("http"):
        return href
    return urljoin(BASE_URL + "/", href.lstrip("/"))

def extract_email(text):
    if not text:
        return ""
    m = EMAIL_PATTERN.search(text)
    return m.group(0) if m else ""

def first_external_email(text):
    """First email in text that is NOT a jobsbotswana.info address."""
    for m in EMAIL_PATTERN.finditer(text or ""):
        e = m.group(0)
        if SITE_HOST not in e.lower():
            return e
    return ""

def classify_application(value):
    """Classify a WP Job Manager _application value (or any string) into
    (apply_url, apply_email). Site-own addresses/URLs are rejected."""
    v = (value or "").strip()
    if not v:
        return "", ""
    if v.lower().startswith("mailto:"):
        v = v.split(":", 1)[1].split("?")[0].strip()
    if re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", v):
        return ("", v) if SITE_HOST not in v.lower() else ("", "")
    if v.lower().startswith(("http://", "https://", "www.")):
        url = v if v.lower().startswith("http") else "https://" + v
        host = urlparse(url).netloc.lower()
        if SITE_HOST in host or SOCIAL_HOST_RE.search(host):
            return "", ""
        return url, ""
    em = first_external_email(v)               # email embedded in free text
    return ("", em) if em else ("", "")

# =============================================================================
#  COMPANY LOGO HELPERS
# =============================================================================

PLACEHOLDER_LOGO_RE = re.compile(r"default|placeholder|avatar|no-?image|blank|generic", re.I)

def clean_logo_url(raw: str) -> str:
    if not raw:
        return ""
    raw = raw.strip()
    if not raw.startswith("http"):
        raw = absolute_url(raw)
    return re.sub(r"[\"')\s]+$", "", raw)

def is_placeholder_logo(url: str) -> bool:
    if not url:
        return True
    return bool(PLACEHOLDER_LOGO_RE.search(url))

# =============================================================================
#  NLP TOOLS (lazy init, optional)
# =============================================================================

_grammar_tool = None
_sim_model    = None

def _get_grammar_tool():
    global _grammar_tool
    if _grammar_tool is None and _NLP_AVAILABLE:
        try:
            _grammar_tool = language_tool_python.LanguageTool(
                "en-US", remote_server="https://api.languagetool.org")
        except Exception as e:
            log_.warning(f"LanguageTool init failed: {e}")
    return _grammar_tool

def _get_sim_model():
    global _sim_model
    if _sim_model is None and _NLP_AVAILABLE:
        try:
            _sim_model = SentenceTransformer("all-MiniLM-L6-v2", device="cpu")
        except Exception as e:
            log_.warning(f"SentenceTransformer init failed: {e}")
    return _sim_model

def grammar_correct(text: str) -> str:
    tool = _get_grammar_tool()
    if tool:
        try:
            return language_tool_python.utils.correct(text, tool.check(text))
        except Exception:
            pass
    return text

def similarity_score(a: str, b: str) -> float:
    model = _get_sim_model()
    if model:
        try:
            emb = model.encode([a, b], convert_to_tensor=True)
            return float(st_util.pytorch_cos_sim(emb[0], emb[1]))
        except Exception:
            pass
    def tokens(s):
        return set(re.sub(r"[^a-z0-9 ]", " ", s.lower()).split())
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb: return 0.0
    return len(ta & tb) / max(len(ta), len(tb))

def clean_output(text: str) -> str:
    text = _fix_mojibake(text)
    for pat in [r"\[/?INST\]", r"</?s>",
                r"(?i)(rewritten?|rephrased?|output|paraphrase[d]?)[:\s]+",
                r"\*\*", r"###", r"---"]:
        text = re.sub(pat, "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return grammar_correct(text.strip())

# =============================================================================
#  MISTRAL API
# =============================================================================

def mistral_generate(prompt: str, max_tokens: int = 400, temperature: float = 0.7) -> str:
    if not MISTRAL_API_KEY:
        log_.warning("MISTRAL_API_KEY not set — skipping paraphrase")
        return ""
    try:
        response = requests.post(
            MISTRAL_URL,
            headers={
                "Authorization": f"Bearer {MISTRAL_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": MISTRAL_MODEL,
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": max_tokens,
                "temperature": temperature,
            },
            timeout=30,
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        log_.error(f"Mistral API error: {e}")
        return ""

# =============================================================================
#  PARAPHRASE FUNCTIONS
# =============================================================================

def _print_wrapped(text: str, prefix: str = "   ", width: int = 100):
    words = text.split()
    line  = []
    for w in words:
        line.append(w)
        if len(" ".join(line)) >= width:
            print(f"{prefix}{' '.join(line)}")
            line = []
    if line:
        print(f"{prefix}{' '.join(line)}")

def paraphrase_title(title: str) -> str:
    if not ENABLE_PARAPHRASE:
        return title
    clean = sanitize_text(title)
    if not clean:
        return title

    print(f"\n ┌─ TITLE PARAPHRASE {'─'*45}")
    print(f" │ Original : \"{clean}\"")
    print(f" │ {'─'*60}")

    best_result = None
    best_sim    = 0.0

    for attempt in range(4):
        temp = round(0.68 + attempt * 0.06, 2)
        print(f" │ Attempt {attempt+1} (temp={temp}):")

        prompt = (
            f"Rewrite this job title professionally using different words. "
            f"Output ONLY the rewritten title, nothing else. "
            f"Keep it between 4 and 12 words.\n\nJob title: {clean}"
        )

        raw    = mistral_generate(prompt, max_tokens=50, temperature=temp)
        result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")

        wc     = len(result.split()) if result else 0
        sim    = similarity_score(clean, result) if result else 0.0
        is_dup = result.lower().strip() == clean.lower().strip()

        print(f" │    Output  : \"{result}\"")
        print(f" │    Words   : {wc} | Similarity: {sim:.3f} | Duplicate: {'Yes ⚠️' if is_dup else 'No'}")

        valid = bool(result) and 4 <= wc <= 14 and sim >= 0.55 and not is_dup

        if not valid:
            reasons = []
            if not result:  reasons.append("empty output")
            if wc < 4:      reasons.append(f"too short ({wc} words, min=4)")
            if wc > 14:     reasons.append(f"too long ({wc} words, max=14)")
            if sim < 0.55:  reasons.append(f"sim={sim:.3f} < 0.55")
            if is_dup:      reasons.append("identical to original")
            print(f" │    → ❌ REJECTED — {', '.join(reasons)}")
        else:
            if sim > best_sim:
                best_sim    = sim
                best_result = result
                print(f" │    → ✅ ACCEPTED — new best candidate (sim={sim:.3f})")
            else:
                print(f" │    → ✅ VALID but not better than current best (best sim={best_sim:.3f})")

        print(f" │ {'─'*60}")
        time.sleep(1)

    if best_result:
        print(f" │ 🏆 FINAL SELECTED : \"{best_result}\"")
        print(f" │    Similarity     : {best_sim:.3f}")
        print(f" └{'─'*65}")
        return best_result
    else:
        print(f" │ ⚠️  No valid paraphrase found → Keeping original: \"{clean}\"")
        print(f" └{'─'*65}")
        return clean

def paraphrase_description(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text

    paragraphs  = [p.strip() for p in re.split(r"\n+", clean) if p.strip()]
    if not paragraphs:
        paragraphs = [clean]
    rewritten   = []
    success_count = 0

    print(f"\n ┌─ DESCRIPTION PARAPHRASE ({len(paragraphs)} paragraph(s)) {'─'*15}")

    for i, para in enumerate(paragraphs):
        orig_wc = len(para.split())

        print(f"\n │ ┌─ Paragraph {i+1}/{len(paragraphs)} {'─'*50}")
        print(f" │ │ ORIGINAL ({orig_wc} words):")
        _print_wrapped(para, prefix=" │ │    ")
        print(f" │ │ {'─'*60}")

        prompt = (
            f"Rewrite this job description paragraph professionally. "
            f"Keep ALL facts, requirements, and responsibilities. "
            f"Use different sentence structure and vocabulary. "
            f"Output ONLY the rewritten paragraph — no labels, no explanation.\n\n"
            f"Original:\n{para}"
        )

        best_result = None
        best_sim    = 0.0
        accepted_text = None

        for attempt in range(3):
            temp = round(0.65 + attempt * 0.08, 2)
            print(f" │ │ Attempt {attempt+1}/3 (temp={temp}):")

            raw    = mistral_generate(prompt, max_tokens=500, temperature=temp)
            result = clean_output(raw).strip()

            rw  = len(result.split()) if result else 0
            sim = similarity_score(para, result) if result and rw >= 5 else 0.0

            if result:
                print(f" │ │    Paraphrased ({rw} words, sim={sim:.3f}):")
                _print_wrapped(result, prefix=" │ │       ")
            else:
                print(f" │ │    Paraphrased : (no output from model)")

            valid = bool(result) and rw >= 8 and sim >= 0.48

            if not valid:
                reasons = []
                if not result: reasons.append("empty output")
                if rw < 8:     reasons.append(f"too short ({rw} words, min=8)")
                if sim < 0.48: reasons.append(f"sim={sim:.3f} < 0.48")
                print(f" │ │    → ❌ REJECTED — {', '.join(reasons)}")
                if result and sim > best_sim:
                    best_sim    = sim
                    best_result = result
                    print(f" │ │       (stored as best fallback, sim={sim:.3f})")
            else:
                print(f" │ │    → ✅ ACCEPTED on attempt {attempt+1}")
                rewritten.append(result)
                success_count += 1
                accepted_text = result
                break

            print(f" │ │ {'─'*60}")
            time.sleep(1)

        if accepted_text is None:
            print(f" │ │ {'─'*60}")
            if best_result and best_sim >= 0.40:
                print(f" │ │ 🔁 FALLBACK — Using best attempt (sim={best_sim:.3f}):")
                _print_wrapped(best_result, prefix=" │ │    ")
                rewritten.append(best_result)
                success_count += 1
            else:
                print(f" │ │ ⚠️  KEPT ORIGINAL — no acceptable paraphrase (best sim={best_sim:.3f})")
                rewritten.append(para)

        print(f" │ └{'─'*62}")

    print(f"\n │ SUMMARY: {success_count}/{len(paragraphs)} paragraphs successfully paraphrased")
    print(f" └{'─'*80}\n")

    return "\n\n".join(rewritten)

def paraphrase_company(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text

    print(f"\n ┌─ COMPANY BLURB PARAPHRASE {'─'*37}")
    orig_wc = len(clean.split())
    print(f" │ Original ({orig_wc} words):")
    _print_wrapped(clean, prefix=" │    ")
    print(f" │ {'─'*60}")

    prompt = (
        f"Rewrite this company description professionally. "
        f"Preserve all facts. Use different wording. "
        f"Output ONLY the rewritten description.\n\nOriginal:\n{clean}"
    )

    raw    = mistral_generate(prompt, max_tokens=600, temperature=0.68)
    result = clean_output(raw)
    rw     = len(result.split()) if result else 0
    sim    = similarity_score(clean, result) if result and rw >= 10 else 0.0

    if result and rw >= 10:
        print(f" │ Paraphrased ({rw} words, sim={sim:.3f}):")
        _print_wrapped(result, prefix=" │    ")
        print(f" │ → ✅ ACCEPTED")
        print(f" └{'─'*65}")
        time.sleep(1)
        return result
    else:
        reasons = []
        if not result: reasons.append("empty output")
        if rw < 10:    reasons.append(f"too short ({rw} words, min=10)")
        print(f" │ → ❌ REJECTED — {', '.join(reasons)} — keeping original")
        print(f" └{'─'*65}")
        time.sleep(1)
        return clean

# =============================================================================
#  DUPLICATE TRACKER (persists across runs)
# =============================================================================

def _init_tracker():
    if not _XLSX_AVAILABLE:
        return
    if not os.path.exists(PROCESSED_IDS_FILE):
        pd.DataFrame(columns=[
            "Job ID", "Job URL", "Job Title", "Company Name",
            "Status", "Timestamp", "WP ID",
        ]).to_csv(PROCESSED_IDS_FILE, index=False)

def load_processed_ids() -> tuple:
    if not _XLSX_AVAILABLE:
        log_.warning("pandas not installed — duplicate tracking is in-run only, not persisted")
        return set(), set()
    _init_tracker()
    df = pd.read_csv(PROCESSED_IDS_FILE)
    return (
        set(df["Job ID"].fillna("").astype(str)),
        set(df.get("Job URL", pd.Series()).fillna("").astype(str)),
    )

def _upsert_row(job_id: str, updates: dict):
    if not _XLSX_AVAILABLE:
        return
    _init_tracker()
    df   = pd.read_csv(PROCESSED_IDS_FILE)
    mask = df["Job ID"].astype(str) == str(job_id)
    if mask.any():
        for col, val in updates.items():
            if col in df.columns:
                df.loc[mask, col] = val
        df.loc[mask, "Timestamp"] = datetime.now().isoformat()
    else:
        row = {"Job ID": job_id, "Timestamp": datetime.now().isoformat()}
        row.update(updates)
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_csv(PROCESSED_IDS_FILE, index=False)

def make_job_id(job_url: str, title: str = "", company: str = "") -> str:
    # Every job has a unique /jobs/<slug>/ permalink — the stable primary key.
    if job_url:
        return hashlib.md5(job_url.encode()).hexdigest()[:16]
    if title or company:
        seed = f"{title}|{company}"
        return hashlib.md5(seed.encode()).hexdigest()[:16]
    return hashlib.md5(b"unknown").hexdigest()[:16]

def mark_scraped(job_id, job_url, title, company):
    _upsert_row(job_id, {"Job URL": job_url, "Job Title": title,
                          "Company Name": company, "Status": "scraped"})

def mark_paraphrased(job_id):
    _upsert_row(job_id, {"Status": "paraphrased"})

def mark_posted(job_id, wp_id, wp_url):
    _upsert_row(job_id, {"Status": "posted", "WP ID": wp_id})

def mark_failed(job_id, reason):
    _upsert_row(job_id, {"Status": f"failed|{reason}"})

# =============================================================================
#  WORDPRESS POSTING (destination site — unchanged)
# =============================================================================

def _wp_auth_headers() -> dict:
    token = base64.b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()
    return {"Authorization": f"Basic {token}", "Content-Type": "application/json"}

def get_or_create_term(taxonomy_url: str, name: str):
    if not name or not name.strip():
        return None
    slug = re.sub(r"[^a-z0-9-]", "-", name.lower().strip())
    h = _wp_auth_headers()
    try:
        r = requests.get(f"{taxonomy_url}?slug={slug}", headers=h, timeout=10, verify=False)
        terms = r.json()
        if isinstance(terms, list) and terms:
            return terms[0]["id"]
    except Exception:
        pass
    try:
        r = requests.post(taxonomy_url, json={"name": name, "slug": slug},
                          headers=h, auth=(WP_USER, WP_PASSWORD), timeout=10, verify=False)
        return r.json().get("id")
    except Exception as e:
        log_.error(f"Term create error '{name}': {e}")
        return None

def post_job_to_wordpress(job: dict) -> tuple:
    if not WP_USER or not WP_PASSWORD:
        log_.warning("WP_USERNAME / WP_APP_PASSWORD not set — skipping WordPress post")
        return None, None

    h = _wp_auth_headers()

    title       = sanitize_text(job.get("jobTitle", ""))
    description = sanitize_text(job.get("jobDescription", ""))
    if not title or not description:
        return None, None

    slug = re.sub(r"[^a-z0-9-]", "-", title.lower())[:80]
    try:
        r = requests.get(f"{WP_JOBS_URL}?slug={slug}", headers=h, timeout=10, verify=False)
        posts = r.json()
        if isinstance(posts, list) and posts:
            log_.info(f"⏭ Job already on WP: {title}")
            return posts[0]["id"], posts[0].get("link")
    except Exception:
        pass

    logo_url    = sanitize_text(job.get("companyLogo", ""), is_url=True)
    location    = sanitize_text(job.get("jobLocation", ""))
    raw_type    = sanitize_text(job.get("jobType", "")) or "Full-time"
    job_type_s  = JOB_TYPE_MAPPING.get(raw_type.lower().strip(), "full-time")
    company     = sanitize_text(job.get("companyName", ""))
    application = sanitize_text(job.get("application", ""), is_url=True)
    company_url = sanitize_text(job.get("companyUrl", ""), is_url=True)
    deadline    = sanitize_text(job.get("deadline", ""))
    co_website  = sanitize_text(job.get("companyWebsite", ""), is_url=True)
    qualif      = sanitize_text(job.get("jobQualifications", ""))
    experience  = sanitize_text(job.get("jobExperience", ""))
    co_address  = sanitize_text(job.get("companyAddress", ""))
    job_field   = sanitize_text(job.get("jobField", ""))
    salary      = sanitize_text(job.get("salaryRange", ""))
    about       = sanitize_text(job.get("companyDetails", ""))

    is_email = bool(re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", application))
    is_url_v = bool(re.match(r"^https?://[^\s]+$", application))
    if not (is_email or is_url_v):
        application = ""

    # Upload logo
    attachment_id = None
    if logo_url:
        try:
            img_r = requests.get(logo_url, headers={"User-Agent": HEADERS["User-Agent"]}, timeout=15)
            if img_r.status_code == 200:
                ct  = img_r.headers.get("Content-Type", "image/jpeg")
                ext = "png" if "png" in ct else "jpg"
                fn  = re.sub(r"[^a-z0-9]", "-", company.lower()) + "-logo." + ext
                up_h = dict(_wp_auth_headers())
                up_h["Content-Disposition"] = f"attachment; filename={fn}"
                up_h["Content-Type"] = ct
                up_r = requests.post(WP_MEDIA_URL, headers=up_h, data=img_r.content,
                                     auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
                if up_r.status_code in (200, 201):
                    attachment_id = up_r.json().get("id")
        except Exception as e:
            log_.warning(f"Logo upload failed: {e}")

    region_term_id   = get_or_create_term(f"{WP_BASE}/job_listing_region", location)
    job_type_term_id = get_or_create_term(f"{WP_BASE}/job_listing_type",
                                           job_type_s.replace("-", " ").title())

    payload = {
        "title":          title,
        "content":        description,
        "status":         "publish",
        "featured_media": attachment_id or 0,
        "meta": {
            "_job_title":          title,
            "_job_location":       location,
            "_job_type":           job_type_s,
            "_job_description":    description,
            "_application":        application,
            "_company_url":        company_url,
            "_job_expires":        deadline,
            "_company_name":       company,
            "_company_website":    co_website,
            "_company_logo":       str(attachment_id) if attachment_id else "",
            "_company_address":    co_address,
            "_company_details":    about,
            "_job_qualifications": qualif,
            "_job_experiences":    experience,
            "_job_field":          job_field,
            "_job_salary":         salary,
        },
    }
    if region_term_id:   payload["job_listing_region"] = [region_term_id]
    if job_type_term_id: payload["job_listing_type"]   = [job_type_term_id]

    for attempt in range(3):
        try:
            r = requests.post(WP_JOBS_URL, json=payload, headers=h,
                              auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
            r.raise_for_status()
            post = r.json()
            log_.info(f"✅ Job posted: '{title}' → WP ID {post.get('id')}")
            return post.get("id"), post.get("link")
        except Exception as e:
            log_.error(f"Job post attempt {attempt+1} failed: {e}")
            if attempt < 2:
                time.sleep(2 ** attempt)
    return None, None

# =============================================================================
#  SOURCE — PATH A: WP REST API  (/wp-json/wp/v2/job-listings)
# =============================================================================

def _term_names(item, taxonomy):
    out = []
    for group in item.get("_embedded", {}).get("wp:term", []) or []:
        if isinstance(group, list):
            for term in group:
                if isinstance(term, dict) and term.get("taxonomy") == taxonomy:
                    name = (term.get("name") or "").strip()
                    if name:
                        out.append(name)
    return out

def _featured_logo(item):
    for m in item.get("_embedded", {}).get("wp:featuredmedia", []) or []:
        if not isinstance(m, dict):
            continue
        src = m.get("source_url")
        if not src:
            sizes = (m.get("media_details", {}) or {}).get("sizes", {}) or {}
            for k in ("medium", "full", "thumbnail"):
                if sizes.get(k, {}).get("source_url"):
                    src = sizes[k]["source_url"]; break
        if src and not is_placeholder_logo(src):
            return clean_logo_url(src)
    return ""

def fetch_rest_jobs():
    """Enumerate all job_listing posts via the WP REST API. Returns a list of
    raw API items, or None if the endpoint is unavailable (→ HTML fallback)."""
    items = []
    page = 1
    while True:
        url = (f"{WPJM_REST}?per_page={REST_PER_PAGE}&page={page}"
               f"&_embed=1&orderby=date&order=desc")
        try:
            r = SESSION.get(url, timeout=REQUEST_TIMEOUT)
        except Exception as e:
            log(C_RED(f"  REST request error: {e}"))
            return None if page == 1 else items

        if r.status_code in (400, 404):
            # 400 with code rest_post_invalid_page_number = past the last page
            try:
                code = (r.json() or {}).get("code", "")
            except Exception:
                code = ""
            if page > 1 and "page_number" in code:
                break
            if page == 1:
                log(C_RED(f"  REST endpoint not available (HTTP {r.status_code}, code={code or '—'})"))
                return None
            break

        if r.status_code != 200:
            if page == 1:
                log(C_RED(f"  REST endpoint returned HTTP {r.status_code}"))
                return None
            break

        try:
            data = r.json()
        except Exception:
            return None if page == 1 else items

        if isinstance(data, dict) and data.get("code"):
            return None if page == 1 else items
        if not isinstance(data, list) or not data:
            break

        items.extend(data)
        total_pages = int(r.headers.get("X-WP-TotalPages", "0") or "0")
        total_jobs  = r.headers.get("X-WP-Total", "?")
        log(f"  REST page {page}"
            + (f"/{total_pages}" if total_pages else "")
            + f": {len(data)} listing(s)"
            + (f"  (total {total_jobs})" if page == 1 else ""))

        if total_pages and page >= total_pages:
            break
        if SCRAPE_PAGES and page >= SCRAPE_PAGES:
            break
        page += 1
        time.sleep(REQUEST_DELAY)

    return items

def parse_rest_item(item):
    """Map one WP REST job_listing object to the standard raw_job dict."""
    meta = item.get("meta", {}) or {}
    link = item.get("link", "") or ""

    title = clean_text(BeautifulSoup((item.get("title", {}) or {}).get("rendered", ""), "lxml"))
    content_html = (item.get("content", {}) or {}).get("rendered", "")
    description = html_to_text(content_html)
    # Raw (un-stripped) body text — used for apply detection, since the
    # "To apply for this job ..." line is removed from `description` as boilerplate.
    content_text = BeautifulSoup(content_html, "lxml").get_text(" ", strip=True) if content_html else ""

    job_types  = _term_names(item, "job_listing_type")
    categories = _term_names(item, "job_listing_category")
    regions    = _term_names(item, "job_listing_region")

    location = (meta.get("_job_location") or "").strip() or (regions[0] if regions else "")
    job_type = (job_types[0] if job_types else "")
    field    = (categories[0] if categories else "")

    company       = (meta.get("_company_name") or "").strip()
    company_site  = (meta.get("_company_website") or "").strip()
    company_blurb = (meta.get("_company_tagline") or "").strip()
    logo          = _featured_logo(item) or clean_logo_url(meta.get("_company_logo") or "")
    if is_placeholder_logo(logo):
        logo = ""

    deadline = (meta.get("_job_expires") or "").strip()
    salary   = (meta.get("_job_salary") or "").strip()
    posted   = (item.get("date") or "")[:10]

    apply_url, apply_email = classify_application(meta.get("_application") or "")
    if not apply_url and not apply_email:
        apply_url, apply_email = _apply_from_text(content_text)
    if not apply_url and not apply_email:
        apply_email = first_external_email(content_text)

    return {
        "title":          title,
        "job_url":        link,
        "job_type":       job_type,
        "qualification":  "",
        "experience":     "",
        "location":       location,
        "city":           location,
        "field":          field,
        "posted_date":    posted,
        "deadline":       deadline,
        "description":    description,
        "apply_url":      apply_url,
        "apply_email":    apply_email,
        "apply_raw":      (meta.get("_application") or ""),
        "company_name":   company,
        "company_url":    company_site,
        "company_blurb":  company_blurb,
        "company_logo":   logo,
        "company_address": "",
        "salary":         salary,
        "source_page":    WPJM_REST,
        "_source_id":     item.get("id"),
    }

# =============================================================================
#  SOURCE — PATH B: HTML + JSON-LD  (single /jobs/<slug>/ pages)
# =============================================================================

def _is_single_job_url(href):
    try:
        p = urlparse(href)
    except Exception:
        return False
    if p.netloc and SITE_HOST not in p.netloc.lower():
        return False
    path = p.path
    return JOB_SLUG_PREFIX in path and path.rstrip("/") not in ("/jobs", "")

def discover_job_urls_via_sitemap():
    """Find /jobs/<slug>/ URLs from the job_listing sitemap (Yoast/RankMath/
    core), recursing through a sitemap index if needed."""
    urls = set()
    candidates = [
        f"{BASE_URL}/job_listing-sitemap.xml",
        f"{BASE_URL}/job_listing-sitemap1.xml",
        f"{BASE_URL}/job-sitemap.xml",
        f"{BASE_URL}/sitemap_index.xml",
        f"{BASE_URL}/wp-sitemap.xml",
        f"{BASE_URL}/sitemap.xml",
    ]
    loc_re = re.compile(r"<loc>\s*([^<]+?)\s*</loc>", re.I)
    for sm in candidates:
        try:
            r = SESSION.get(sm, timeout=REQUEST_TIMEOUT)
        except Exception:
            continue
        if r.status_code != 200 or "<loc" not in r.text.lower():
            continue
        locs = loc_re.findall(r.text)
        for loc in locs:
            low = loc.lower()
            if low.endswith(".xml") and ("job_listing" in low or "job-sitemap" in low):
                try:
                    rr = SESSION.get(loc, timeout=REQUEST_TIMEOUT)
                    for u in loc_re.findall(rr.text):
                        if _is_single_job_url(u):
                            urls.add(u.strip())
                except Exception:
                    pass
            elif _is_single_job_url(loc):
                urls.add(loc.strip())
        if urls:
            break
    return sorted(urls)

def discover_job_urls_via_listing():
    """Fallback discovery: scrape /jobs/<slug>/ links off the homepage and the
    main jobs listing page (only the first non-AJAX batch, but better than
    nothing)."""
    urls = set()
    for path in ("/jobs/", "/"):
        try:
            soup = get_soup(BASE_URL + path)
        except Exception:
            continue
        for a in soup.find_all("a", href=True):
            full = absolute_url(a["href"])
            if _is_single_job_url(full):
                urls.add(full)
    return sorted(urls)

def _iter_jsonld(data):
    if isinstance(data, list):
        for x in data:
            yield from _iter_jsonld(x)
    elif isinstance(data, dict):
        yield data
        if "@graph" in data:
            yield from _iter_jsonld(data["@graph"])

def parse_jsonld_jobposting(soup):
    """Return the JobPosting JSON-LD object on a single job page, or None."""
    for tag in soup.find_all("script", attrs={"type": "application/ld+json"}):
        raw = tag.string or tag.get_text() or ""
        if not raw.strip():
            continue
        try:
            data = json.loads(raw)
        except Exception:
            try:
                data = json.loads(re.sub(r",\s*([}\]])", r"\1", raw))   # tolerate trailing commas
            except Exception:
                continue
        for obj in _iter_jsonld(data):
            if isinstance(obj, dict) and "JobPosting" in str(obj.get("@type", "")):
                return obj
    return None

def _jsonld_value(v):
    """JSON-LD fields are sometimes str, sometimes {url/name/value}."""
    if isinstance(v, dict):
        return v.get("url") or v.get("name") or v.get("value") or ""
    if isinstance(v, list) and v:
        return _jsonld_value(v[0])
    return v or ""

def _apply_from_text(text):
    """Pull an apply email / URL out of WP Job Manager's auto-generated
    'To apply for this job ...' line in free text."""
    if not text:
        return "", ""
    m = APPLY_TEXT_EMAIL_RE.search(text)
    if m and SITE_HOST not in m.group(1).lower():
        return "", m.group(1).rstrip(".")
    m = APPLY_TEXT_URL_RE.search(text)
    if m and SITE_HOST not in urlparse(m.group(1)).netloc.lower():
        return m.group(1).rstrip("."), ""
    return "", ""

def extract_application_from_page(soup):
    """Apply email / URL from the WP Job Manager application widget or the
    auto-generated apply line."""
    for sel in (".application_details a[href]", "a.application_button[href]",
                ".application a[href]", ".wpjb-apply a[href]"):
        for a in soup.select(sel):
            href = (a.get("href") or "").strip()
            low = href.lower()
            if low.startswith("mailto:"):
                em = href.split(":", 1)[1].split("?")[0].strip()
                if em and SITE_HOST not in em.lower():
                    return "", em
            elif low.startswith("http"):
                host = urlparse(href).netloc.lower()
                if SITE_HOST not in host and not SOCIAL_HOST_RE.search(host):
                    return href, ""
    return _apply_from_text(soup.get_text(" ", strip=True))

def parse_single_page(url):
    """Parse one /jobs/<slug>/ page via JobPosting JSON-LD with WP Job Manager
    CSS hooks as fallback. Returns the standard raw_job dict (or None)."""
    try:
        soup = get_soup(url)
    except Exception as e:
        log(C_RED(f"    ✗ fetch failed: {url} ({e})"))
        return None

    jp = parse_jsonld_jobposting(soup) or {}

    # ---- title ----
    title = clean_text(BeautifulSoup(str(_jsonld_value(jp.get("title"))), "lxml")) if jp.get("title") else ""
    if not title:
        el = soup.select_one("h1.entry-title") or soup.select_one(".single_job_listing h1") or soup.find("h1")
        title = clean_text(el) if el else ""
    if not title and soup.title:
        title = re.split(r"\s*[|\-–]\s*", soup.title.get_text(strip=True))[0].strip()

    # ---- description ----
    description = ""
    if jp.get("description"):
        description = html_to_text(str(jp["description"]))
    if not description:
        el = (soup.select_one("div.job_description")
              or soup.select_one(".single_job_listing .job_description")
              or soup.select_one("[itemprop='description']")
              or soup.select_one(".entry-content"))
        if el:
            description = html_to_text(el.decode_contents())

    # ---- company ----
    company = _jsonld_value((jp.get("hiringOrganization") or {})) if isinstance(jp.get("hiringOrganization"), dict) else ""
    if not company:
        el = soup.select_one(".company .fn") or soup.select_one(".company strong") or soup.select_one(".company-name")
        company = clean_text(el) if el else ""

    company_site = ""
    el = soup.select_one(".company a[href^='http']")
    if el:
        href = el.get("href", "")
        if SITE_HOST not in urlparse(href).netloc.lower():
            company_site = href

    # ---- logo ----
    logo = ""
    ho = jp.get("hiringOrganization") or {}
    if isinstance(ho, dict) and ho.get("logo"):
        logo = clean_logo_url(_jsonld_value(ho.get("logo")))
    if not logo or is_placeholder_logo(logo):
        img = soup.select_one("img.company_logo") or soup.select_one(".company img")
        if img:
            logo = clean_logo_url(img.get("src") or img.get("data-src") or "")
    if is_placeholder_logo(logo):
        logo = ""

    # ---- meta ----
    job_type = EMPLOYMENT_TYPE_HUMAN.get(str(_jsonld_value(jp.get("employmentType"))).upper().strip(),
                                         str(_jsonld_value(jp.get("employmentType"))).strip())
    if not job_type:
        el = soup.select_one("li.job-type") or soup.select_one(".job-type")
        job_type = clean_text(el) if el else ""

    location = ""
    jl = jp.get("jobLocation")
    if jl:
        addr = (jl[0] if isinstance(jl, list) and jl else jl)
        if isinstance(addr, dict):
            a = addr.get("address", addr)
            if isinstance(a, dict):
                location = (a.get("addressLocality") or a.get("addressRegion")
                            or a.get("streetAddress") or "").strip()
    if not location:
        el = soup.select_one("li.location") or soup.select_one(".location") or soup.select_one(".job-geo")
        location = clean_text(el) if el else ""

    deadline = str(_jsonld_value(jp.get("validThrough")))[:10]
    posted   = str(_jsonld_value(jp.get("datePosted")))[:10]
    if not posted:
        t = soup.select_one("time[datetime]")
        posted = (t.get("datetime", "")[:10] if t else "")

    salary = ""
    bs = jp.get("baseSalary")
    if isinstance(bs, dict):
        val = bs.get("value")
        if isinstance(val, dict):
            salary = str(val.get("value") or val.get("minValue") or "").strip()

    # ---- application ----
    apply_url, apply_email = extract_application_from_page(soup)
    if not apply_url and not apply_email:
        apply_url, apply_email = _apply_from_text(description)
    if not apply_url and not apply_email:
        apply_email = first_external_email(description)

    return {
        "title":          title,
        "job_url":        url,
        "job_type":       job_type,
        "qualification":  "",
        "experience":     "",
        "location":       location,
        "city":           location,
        "field":          "",
        "posted_date":    posted,
        "deadline":       deadline,
        "description":    description,
        "apply_url":      apply_url,
        "apply_email":    apply_email,
        "apply_raw":      "",
        "company_name":   company,
        "company_url":    company_site,
        "company_blurb":  "",
        "company_logo":   logo,
        "company_address": "",
        "salary":         salary,
        "source_page":    url,
    }

# =============================================================================
#  SOURCE ORCHESTRATION
# =============================================================================

def collect_and_parse_jobs(known_ids=None, known_urls=None):
    """Collect + parse jobs via REST (preferred) or HTML fallback. Jobs already
    in the tracker are skipped. Returns (raw_jobs, source_pages)."""
    known_ids  = known_ids  or set()
    known_urls = known_urls or set()

    raw_jobs, sources = [], []

    rest_items = None
    if SOURCE_MODE in ("auto", "rest"):
        log(f"\n{'=' * 80}\nSOURCE: WP REST API → {WPJM_REST}\n{'=' * 80}")
        rest_items = fetch_rest_jobs()

    if rest_items:
        sources.append(WPJM_REST)
        log(f"\n  REST returned {len(rest_items)} listing(s); parsing …")
        for i, it in enumerate(rest_items, 1):
            link = it.get("link", "") or ""
            if link and (link in known_urls or make_job_id(link) in known_ids):
                log(C_DIM(f"  ⧳ [{i}/{len(rest_items)}] already in tracker — skipped: {link}"))
                continue
            raw = parse_rest_item(it)

            # Enrich from the single page if REST meta was sparse.
            needs = (not raw["description"]
                     or (not raw["apply_url"] and not raw["apply_email"])
                     or not raw["company_name"])
            if raw.get("job_url") and needs:
                page = parse_single_page(raw["job_url"])
                if page:
                    for k in ("description", "apply_url", "apply_email", "company_name",
                              "company_url", "company_logo", "company_blurb",
                              "job_type", "location", "deadline", "salary", "posted_date"):
                        if not raw.get(k) and page.get(k):
                            raw[k] = page[k]
                    raw["city"] = raw["location"]
                time.sleep(REQUEST_DELAY)

            if raw and (raw["title"] or raw["description"]):
                raw_jobs.append(raw)
        return raw_jobs, sources

    if SOURCE_MODE == "rest":
        log(C_RED("  REST unavailable and SOURCE_MODE=rest — nothing to do."))
        return [], []

    # ---- HTML fallback ----
    log(f"\n{'=' * 80}\nSOURCE: HTML + JSON-LD fallback\n{'=' * 80}")
    urls = discover_job_urls_via_sitemap()
    if not urls:
        log(C_DIM("  No sitemap URLs — falling back to listing-page link discovery"))
        urls = discover_job_urls_via_listing()
    if SCRAPE_PAGES:
        urls = urls[: SCRAPE_PAGES * REST_PER_PAGE]
    log(f"  Discovered {len(urls)} job URL(s)")
    sources.append("(html) " + (urls[0] if urls else "no urls"))

    for i, u in enumerate(urls, 1):
        if u in known_urls or make_job_id(u) in known_ids:
            log(C_DIM(f"  ⧳ [{i}/{len(urls)}] already in tracker — skipped"))
            continue
        log(C_DIM(f"  → [{i}/{len(urls)}] {u}"))
        raw = parse_single_page(u)
        if raw and (raw["title"] or raw["description"]):
            raw_jobs.append(raw)
        else:
            log(C_RED(f"    ✗ no usable content parsed: {u}"))
        time.sleep(REQUEST_DELAY)

    return raw_jobs, sources

# =============================================================================
#  STEP — DEDUPLICATE + PARAPHRASE
# =============================================================================

def process_job(raw_job: dict, processed_ids: set, processed_urls: set, seen_content: set):
    """
    Applies persistent + in-run duplicate detection, then paraphrases the
    title/description/company blurb via Mistral, and returns the standardized
    job dict ready for WordPress posting / Excel export.
    Returns None if the job was a duplicate (and should be skipped).
    """
    job_url  = raw_job.get("job_url", "")
    title    = raw_job.get("title", "")
    company  = raw_job.get("company_name", "")
    location = raw_job.get("location") or raw_job.get("city", "")

    job_id = make_job_id(job_url, title, company)

    if job_id in processed_ids:
        log(C_DIM(f"  ⧳ Already processed (tracker) — skipped: {title} @ {company}"))
        return None

    fingerprint = (title.lower().strip(), company.lower().strip(), location.lower().strip())
    if fingerprint in seen_content:
        log(C_DIM(f"  ⧳ Duplicate content this run — skipped: {title}"))
        return None
    seen_content.add(fingerprint)

    mark_scraped(job_id, job_url, title, company)
    processed_ids.add(job_id)
    processed_urls.add(job_url)

    description = raw_job.get("description", "")
    blurb       = raw_job.get("company_blurb", "")

    paraphrased_title = title
    paraphrased_desc  = description
    paraphrased_blurb = blurb

    if ENABLE_PARAPHRASE and MISTRAL_API_KEY:
        print(C_BLUE(f"\n  ✍️  Paraphrasing '{title}' ..."))
        paraphrased_title = paraphrase_title(title)
        paraphrased_desc  = paraphrase_description(description)
        if blurb:
            paraphrased_blurb = paraphrase_company(blurb)
        mark_paraphrased(job_id)
    else:
        print(C_DIM("  ⚠️  Paraphrasing skipped (ENABLE_PARAPHRASE=False or MISTRAL_API_KEY not set)"))

    apply_url   = raw_job.get("apply_url", "")
    apply_email = raw_job.get("apply_email", "")
    application = apply_url or apply_email

    apply_method = ("resolved_url" if apply_url
                    else ("description_email" if apply_email else "not_found"))

    return {
        # Paraphrased fields
        "jobTitle":          paraphrased_title,
        "jobDescription":    paraphrased_desc,
        "companyDetails":    paraphrased_blurb,
        # Original fields (audit / duplicate detection)
        "originalTitle":     title,
        "originalDesc":      description,
        # Structured fields
        "jobType":           raw_job.get("job_type", ""),
        "jobQualifications": raw_job.get("qualification", ""),
        "jobExperience":     raw_job.get("experience", ""),
        "jobLocation":       location,
        "jobField":          raw_job.get("field", ""),
        "datePosted":        raw_job.get("posted_date", ""),
        "deadline":          raw_job.get("deadline", ""),
        "application":       application,
        "companyUrl":        raw_job.get("company_url", ""),
        "companyName":       company,
        "companyLogo":       raw_job.get("company_logo", ""),
        "companyWebsite":    raw_job.get("company_url", ""),
        "companyAddress":    raw_job.get("company_address") or raw_job.get("city", ""),
        "jobUrl":            job_url,
        "salaryRange":       raw_job.get("salary", ""),
        "_jobId":            job_id,
        "_apply_method":     apply_method,
        "_apply_raw":        raw_job.get("apply_raw", ""),
    }

# =============================================================================
#  VERBOSE PRINTER
# =============================================================================

def print_job_verbose(index, job):
    desc = job.get("jobDescription", "")
    desc_preview = (desc[:400] + " [...]") if len(desc) > 400 else desc

    print()
    print(C_DIVIDER())
    print(C_HEADER(f"  JOB #{index}"))
    print(C_DIVIDER())
    print(f"  {C_LABEL('Title (original)')}    : {C_VALUE(job.get('originalTitle',''))}")
    print(f"  {C_LABEL('Title (paraphrased)')} : {C_GREEN(job.get('jobTitle',''))}")
    print(f"  {C_LABEL('Job Type')}             : {job.get('jobType','') or C_DIM('—')}")
    print(f"  {C_LABEL('Experience')}           : {job.get('jobExperience','') or C_DIM('—')}")
    print(f"  {C_LABEL('Location')}             : {job.get('jobLocation','') or C_DIM('—')}")
    print(f"  {C_LABEL('Field/Category')}       : {job.get('jobField','') or C_DIM('—')}")
    print(f"  {C_LABEL('Salary')}               : {job.get('salaryRange','') or C_DIM('—')}")
    print(f"  {C_LABEL('Posted')}               : {job.get('datePosted','') or C_DIM('—')}")
    print(f"  {C_LABEL('Deadline')}             : {job.get('deadline','') or C_DIM('—')}")

    application = job.get("application", "")
    print(f"  {C_LABEL('Apply')}                : {C_GREEN(application) if application else C_DIM('— not found —')}")
    print(f"  {C_LABEL('Apply Method')}         : {C_DIM(job.get('_apply_method',''))}")

    print()
    print(f"  {C_BLUE('── COMPANY ──────────────────────────────────────────')}")
    print(f"  {C_LABEL('Name')}      : {C_VALUE(job.get('companyName','') or C_DIM('—'))}")
    print(f"  {C_LABEL('Website')}   : {job.get('companyWebsite','') or C_DIM('—')}")
    print(f"  {C_LABEL('Logo')}      : {job.get('companyLogo','') or C_DIM('— none —')}")
    about = job.get("companyDetails", "")
    if about:
        preview = (about[:200] + " [...]") if len(about) > 200 else about
        print(f"  {C_LABEL('About')}     : {preview}")

    print()
    print(f"  {C_BLUE('── DESCRIPTION PREVIEW ─────────────────────────────')}")
    print(desc_preview if desc_preview else C_DIM("   — no description —"))
    print(f"  {C_LABEL('Job URL')}   : {job.get('jobUrl','')}")
    print(C_DIVIDER())

# =============================================================================
#  EXCEL SAVE (standardized column order)
# =============================================================================

EXCEL_HEADERS = [
    "Job Title", "Job Type", "Job Qualifications", "Job Experience",
    "Job Location", "Job Field", "Date Posted", "Deadline",
    "Job Description", "Application", "Company URL", "Company Name",
    "Company Logo", "Company Website", "Company Address",
    "Company Details", "Job URL", "Salary Range",
]

def _save_excel(jobs: list):
    if not _XLSX_AVAILABLE:
        log_.warning("pandas/openpyxl not installed — skipping Excel export")
        return
    if not jobs:
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(EXCEL_HEADERS)
    for job in jobs:
        ws.append([
            job["jobTitle"], job["jobType"], job["jobQualifications"], job["jobExperience"],
            job["jobLocation"], job["jobField"], job["datePosted"], job["deadline"],
            job["jobDescription"], job["application"], job["companyUrl"], job["companyName"],
            job["companyLogo"], job["companyWebsite"], job["companyAddress"],
            job["companyDetails"], job["jobUrl"], job["salaryRange"],
        ])
    wb.save(OUTPUT_FILE)
    log_.info(f"Saved {len(jobs)} rows → {OUTPUT_FILE}")

# =============================================================================
#  MAIN
# =============================================================================

def main():
    start_time = datetime.now()

    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  JOBSBOTSWANA.INFO SCRAPER + MISTRAL PARAPHRASE + WORDPRESS POSTING"))
    print(C_HEADER("=" * 80))
    print(f"  Source mode     : {SOURCE_MODE}")
    print(f"  Source pages cap: {SCRAPE_PAGES if SCRAPE_PAGES else 'all'}")
    print(f"  Request delay   : {REQUEST_DELAY}s")
    print(f"  Max new jobs    : {'unlimited' if not MAX_JOBS else MAX_JOBS}")
    print(f"  Paraphrase      : {'✅ enabled' if (ENABLE_PARAPHRASE and MISTRAL_API_KEY) else '❌ disabled'}")
    print(f"  WordPress post  : {'✅ enabled' if (WP_USER and WP_PASSWORD) else '❌ disabled'}")
    print(f"  Excel export    : {'✅ enabled' if _XLSX_AVAILABLE else '❌ disabled (pip install pandas openpyxl)'}")
    print(f"  NLP gating      : {'✅' if _NLP_AVAILABLE else '⚠️  no sentence-transformers / language-tool'}")
    print(f"  Started         : {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(C_HEADER("=" * 80))

    _init_tracker()
    processed_ids, processed_urls = load_processed_ids()
    print(f"  Tracker loaded: {len(processed_ids)} previously processed job IDs")

    raw_jobs, source_pages = collect_and_parse_jobs(processed_ids, processed_urls)

    jobs_out = []
    seen_content = set()
    total_raw_jobs = 0
    posted_count = 0
    errors = 0

    for raw_job in raw_jobs:
        total_raw_jobs += 1
        try:
            job = process_job(raw_job, processed_ids, processed_urls, seen_content)
        except Exception as e:
            errors += 1
            log(C_RED(f"  ✗ ERROR processing job: {e}"))
            continue

        if job is None:
            continue

        jobs_out.append(job)
        print_job_verbose(len(jobs_out), job)

        print(C_BLUE("\n  📤 Posting to WordPress …"))
        wp_id, wp_url = post_job_to_wordpress(job)
        if wp_id:
            mark_posted(job["_jobId"], wp_id, wp_url or "")
            posted_count += 1
            print(C_GREEN(f"  ✅ WP ID={wp_id}  🔗 {wp_url}"))
        else:
            mark_failed(job["_jobId"], "wp_post_failed_or_skipped")
            print(C_RED("  ❌ WordPress post failed / skipped"))

        if len(jobs_out) % 25 == 0:
            _save_excel(jobs_out)

        if MAX_JOBS and len(jobs_out) >= MAX_JOBS:
            log(f"\nMAX_JOBS limit ({MAX_JOBS}) reached, stopping.")
            break

    _save_excel(jobs_out)

    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds() / 60.0
    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  SCRAPE COMPLETE"))
    print(C_HEADER("=" * 80))
    print(f"  {C_LABEL('Sources used')}              : {', '.join(source_pages) if source_pages else '—'}")
    print(f"  {C_LABEL('Raw jobs found')}             : {total_raw_jobs}")
    print(f"  {C_LABEL('New jobs processed')}         : {C_GREEN(str(len(jobs_out)))}")
    print(f"  {C_LABEL('Posted to WordPress')}        : {C_GREEN(str(posted_count))}")
    print(f"  {C_LABEL('Errors')}                     : {C_RED(str(errors)) if errors else '0'}")
    print(f"  {C_LABEL('Duration')}                   : ~{duration:.1f} min")
    print(f"  {C_LABEL('Output file')}                : {OUTPUT_FILE}")
    print(f"  {C_LABEL('Tracker file')}               : {PROCESSED_IDS_FILE}")

    if jobs_out:
        with_apply = sum(1 for j in jobs_out if j.get("application"))
        with_email = sum(1 for j in jobs_out if "@" in (j.get("application") or ""))
        with_url   = with_apply - with_email
        no_apply   = len(jobs_out) - with_apply
        print(f"\n  {C_LABEL('Application links:')}")
        print(f"    URL found    : {with_url}")
        print(f"    Email found  : {with_email}")
        print(f"    Not found    : {no_apply}")

        para_count = sum(1 for j in jobs_out if j.get("jobTitle") != j.get("originalTitle"))
        print(f"\n  {C_LABEL('Paraphrased titles')} : {para_count}/{len(jobs_out)}")

        with_logo = sum(1 for j in jobs_out if j.get("companyLogo"))
        print(f"  {C_LABEL('Logos found')}        : {with_logo}/{len(jobs_out)}")

    print(C_HEADER("=" * 80))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)
