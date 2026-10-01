"""Read every candidate posting and note what the tracker does not keep.

The tracker keeps a title and a season label. This script reads the full
posting text of every role shortlist.py could hand out and records, per posting:
  - whether the posting is still up (a board that answers "not found" means it
    was taken down, even when the tracker has not noticed yet)
  - whether the text talks about a summer internship
  - the internship date range, when the posting states one ("June 1 - August 14")
  - who may apply: undergraduates, or PhD / Master's students only
  - the graduation window, when the posting states one, checked against
    GRADUATION (the applicant graduates May 2028)

Results are kept in data/term_cache.json. Every posting is re-read once a day,
since a posting can be taken down at any time; a failed read keeps the facts
from the last good read. shortlist.py reads the cache.

Run:   .venv\\Scripts\\python.exe term_dates.py           read what was not read today
       .venv\\Scripts\\python.exe term_dates.py --report  show what the cache holds
"""

import asyncio
import html as htmllib
import json
import os
import re
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, "src"))

import httpx  # noqa: E402

import shortlist  # noqa: E402
from intern_engine import enrich, sponsorship  # noqa: E402
from intern_engine.models import Job  # noqa: E402
from intern_engine.net import HostLimiter, Net  # noqa: E402

STORE = os.path.join(ROOT, "data", "jobs.json")
CACHE = os.path.join(ROOT, "data", "term_cache.json")
YEAR = 2027
GRADUATION = (2028, 5)  # the applicant's expected graduation, (year, month)
VERSION = 5  # bump when the facts below change, so every posting is re-read

# ----------------------------------------------------------------- dates

MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8,
          "sep": 9, "oct": 10, "nov": 11, "dec": 12}
_MONTH = (r"(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|"
          r"aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)")
# "June 1st, 2027", "June 2027", "June", "early June", "mid-August"
_DATE = (r"(?:(?:early|mid|late|the\s+end\s+of|the\s+beginning\s+of)[\s-]+)?"
         r"(" + _MONTH + r")\.?(?:\s+(\d{1,2})(?!\d)(?:st|nd|rd|th)?)?(?:,?\s*(20\d\d))?")
_SEP = r"\s*(?:-|–|—|to|through|thru|until|till|and\s+(?:end(?:s|ing)?|conclud\w+|run\w*\s+through)\s*(?:in|on)?)\s*"
RE_RANGE = re.compile(_DATE + _SEP + r"(?:the\s+end\s+of\s+)?" + _DATE, re.I)
RE_START = re.compile(
    r"(?:start(?:s|ing)?(?:\s+dates?)?|begin(?:s|ning)?|commenc\w+|kick(?:s|ing)?\s+off)"
    r"(?:\s+(?:is|are|will\s+be|of))?\s*(?:[:\-]\s*)?(?:on|in|from|around|as\s+early\s+as)?\s*"
    r"(?:(?:monday|tuesday|wednesday|thursday|friday),?\s+)?" + _DATE, re.I)
RE_NUMERIC = re.compile(r"(?<![\d/])(\d{1,2})/(\d{1,2})(?:/(?:20)?(\d\d))?\s*(?:-|–|—|to|through)\s*(\d{1,2})/(\d{1,2})(?:/(?:20)?(\d\d))?(?![\d/])")
RE_CONTEXT = re.compile(r"intern|program|summer|session|\bterm\b|start|dates?\b|duration|weeks|cohort|\bruns?\b", re.I)
RE_SUMMER = re.compile(
    r"summer\s+(?:20\d\d\s+)?(?:\w+\s+){0,3}(?:intern|program|analyst|associate|co-?op|session|term)"
    r"|intern\w*[^.!?]{0,40}\bsummer\b|\bsummer\s+20\d\d\b|\b20\d\d\s+summer\b", re.I)
RE_OTHER_TERM = re.compile(
    r"\b(fall|autumn|spring|winter)\s+(?:20\d\d\s+)?(?:intern|co-?op|term|semester|session)"
    r"|\b(?:year[\s-]round|off[\s-]cycle|part[\s-]time\s+intern|6[\s-]month|six[\s-]month|12[\s-]month)", re.I)


def _point(month, day, year):
    """(month, day or None) for a date that is in YEAR or carries no year; None otherwise."""
    if year and int(year) != YEAR:
        return None
    return [MONTHS[month[:3].lower()], int(day) if day else None]


def extract(text):
    """Term facts from plain posting text."""
    out = {"summer": bool(RE_SUMMER.search(text)),
           "other_terms": sorted({m.group(0).lower()[:24] for m in RE_OTHER_TERM.finditer(text)}),
           "start": None, "end": None, "quote": ""}
    for m in RE_RANGE.finditer(text):
        context = text[max(0, m.start() - 140):m.end() + 60]
        if not RE_CONTEXT.search(context):
            continue
        a, b = _point(m[1], m[2], m[3] or m[6]), _point(m[4], m[5], m[6])
        if not a or not b or a[0] > b[0] or b[0] - a[0] > 7:
            continue
        if a[0] == b[0] and (a[1] is None or b[1] is None):
            continue  # "June to June" style noise
        out.update(start=a, end=b, quote=re.sub(r"\s+", " ", m.group(0)).strip(" .,;"))
        return out
    for m in RE_NUMERIC.finditer(text):  # "5/26 - 8/7"
        if not RE_CONTEXT.search(text[max(0, m.start() - 140):m.end() + 60]):
            continue
        years = [y for y in (m[3], m[6]) if y]
        if any(int(y) != YEAR % 100 for y in years):
            continue
        a, b = [int(m[1]), int(m[2])], [int(m[4]), int(m[5])]
        if 1 <= a[0] <= 12 and 1 <= b[0] <= 12 and a[1] <= 31 and b[1] <= 31 and a[0] < b[0] <= a[0] + 7:
            out.update(start=a, end=b, quote=m.group(0).strip())
            return out
    for m in RE_START.finditer(text):
        a = _point(m[1], m[2], m[3])
        if a:
            out.update(start=a, quote=re.sub(r"\s+", " ", m.group(0)).strip(" .,;"))
            return out
    return out


# ----------------------------------------------------------------- who may apply

_BLOCK_TAG = re.compile(r"<\s*/?\s*(?:br|p|li|ul|ol|div|h\d|tr|td|section|article)\b[^>]*>", re.I)
_ABBREV_END = re.compile(r"(?:\b(?:[A-Za-z]{1,2}\.){1,3}|\be\.g\.|\bi\.e\.|\bvs\.|\bapprox\.|\bincl\.)$")


def statements(raw):
    """Posting HTML or text as a list of sentences, keeping bullet boundaries."""
    text = raw or ""
    for _ in range(2):
        text = htmllib.unescape(text)
    text = _BLOCK_TAG.sub("\n", text)
    text = re.sub(r"<[^>]+>", " ", text)
    text = htmllib.unescape(text).replace("’", "'")
    out = []
    for line in text.splitlines():
        line = re.sub(r"[ \t\xa0]+", " ", line).strip(" •·-–*")
        if not line:
            continue
        pieces = re.split(r"(?<=[.!?;])\s+(?=[A-Z•(])", line)
        merged = []
        for p in pieces:
            if merged and _ABBREV_END.search(merged[-1]):
                merged[-1] += " " + p
            else:
                merged.append(p)
        out.extend(merged)
    return out


RE_UG = re.compile(
    r"bachelor|undergrad|\bb\.?\s?[sa]\.?(?=[\s/,)]|$)|\bbs/|\bba/|sophomore|freshm[ae]n"
    r"|\bjuniors?\b(?!\s+(?:developer|engineer|software|analyst|associate|data|quant|trader|researcher|level|role))"
    r"|rising\s+seniors?|four[- ]year|4[- ]year|college\s+(?:students?|seniors?|juniors?)|university\s+students?", re.I)
RE_GR = re.compile(
    r"ph\.?\s?d|doctora\w*|master'?s|\bmasters\b"
    r"|\bm\.?s\.?(?![\s-]*(?:office|excel|word|power|teams|sql|access|project|dynamics|azure|fabric))(?=[\s/,)]|$)"
    r"|\bmba\b|(?<!under)graduate\s+(?:students?|degree|program|studies|school|level)|post-?graduate", re.I)
RE_ANCHOR = re.compile(
    r"pursu\w*|enrol\w*|currently\b|working\s+toward\w*|studying|candidates?\b|students?\b|penultimate"
    r"|rising\s+(?:junior|senior|sophomore)s?|(?:final|first|second|third|fourth|1st|2nd|3rd|4th)[\s-]+year"
    r"|class\s+of\s+20\d\d|in\s+progress|matriculat\w*|eligib\w*|(?:degree|program)\s+in\b"
    r"|requir\w*|qualifications?|must\b|minimum", re.I)
RE_GENERIC = re.compile(  # "currently pursuing a degree in CS": any level may apply
    r"(?:pursu\w*|enrol\w*|working\s+toward\w*|currently\s+(?:attending|studying\s+at))\s+(?:in\s+|at\s+)?"
    r"(?:a|an|your)?\s*(?:accredited\s+|full[- ]time\s+|relevant\s+)?(?:college|university|degree|academic|school)"
    r"|degree[\s-]seeking", re.I)
RE_SOFT = re.compile(r"prefer\w*|ideal\w*|\bplus\b|bonus|nice[\s-]to[\s-]have|desired|desirable|advantage\w*|"
                     r"strongly\s+encouraged|welcome", re.I)
RE_PRIOR_BEFORE = re.compile(r"(?:complet\w*|hold(?:s|ing)?|earned|obtained|possess\w*|with)\s+(?:a|an|your|their)?\s*(?:\w+\s+)?$", re.I)
RE_PRIOR_AFTER = re.compile(r"^\w*\s+(?:studies|coursework|career|education|background|gpa|years)"
                            r"|^[^.;]{0,60}?\bseeking\s+(?:an?\s+)?(?:master|m\.?s\b|ph\.?d|graduate|advanced)", re.I)
SEASON_MONTH = {"spring": 5, "summer": 8, "fall": 12, "autumn": 12, "winter": 12}
_MONTH_PAIR = _MONTH + r"(?:\s*/\s*" + _MONTH + r")?"  # "May/June"
_SEASON = r"(?:spring|summer|fall|autumn|winter)"
_SEASON_PAIR = _SEASON + r"(?:\s*/\s*" + _SEASON + r")?"  # "spring/summer"
# "June 2028", "May/June 2028", "Spring 2028", "winter of 2027", "spring/summer of 2028"
_MY = r"(?:(" + _MONTH_PAIR + r"|" + _SEASON_PAIR + r")\.?\s*,?\s*(?:of\s+)?)?(20\d\d)"
_MY_NOT_SUMMER = (r"(?:(" + _MONTH_PAIR + r"|(?:spring|fall|autumn|winter)(?:\s*/\s*" + _SEASON + r")?)"
                  r"\.?\s*,?\s*(?:of\s+)?)?(20\d\d)")  # a bare "Summer 2027" is the internship, not a graduation
_GRAD_CUE = (r"(?:graduat\w*|class\s+of|full[\s-]time\s+(?:employment|work|roles?|positions?)|degree\s+(?:completion|conferral)"
             r"|complet\w+\s+(?:your|their|the|a)?\s*(?:degree|studies)|awarded|commencement)")
_GRAD_CUE_STRICT = r"(?:graduat\w*|class\s+of|degree\s+(?:completion|conferral)|commencement)"
_FILLER = r"(?:\s+(?!program|internship|intern\b|summer|runs?\b|start)[\w']+){0,5}?"
_FILLER_SHORT = r"(?:\s+(?!for\b|the\b|this\b|our\b|summer\b|program|internship|intern\b|between\b|from\b)[\w']+){0,3}?"
RE_GRAD_RANGE = re.compile(_GRAD_CUE + _FILLER + r"\s*(?:between|from)?\s*" + _MY + r"\s*(?:-|–|—|to|through|and|until)\s*" + _MY, re.I)
RE_GRAD_BEFORE = re.compile(_GRAD_CUE + _FILLER + r"\s*(by|before|no\s+later\s+than|prior\s+to|on\s+or\s+before)\s+" + _MY, re.I)
RE_GRAD_AFTER = re.compile(_GRAD_CUE + _FILLER + r"\s*(after|no\s+earlier\s+than|on\s+or\s+after)\s+" + _MY, re.I)
RE_GRAD_IN = re.compile(_GRAD_CUE_STRICT + _FILLER_SHORT + r"\s*(?:in|of|during|is|date\s+of|:|-)?\s*" + _MY_NOT_SUMMER
                        + r"(?:\s*(?:,|or|and|/|&)\s*" + _MY_NOT_SUMMER + r")?", re.I)
RE_OR_LATER = re.compile(r"^\s*(?:(?:or|and)\s+(?:later|after|beyond)|onwards?|\+)", re.I)
RE_OR_EARLIER = re.compile(r"^\s*(?:or|and)\s+(?:earlier|before|sooner)", re.I)
RE_ALSO = re.compile(r"^\s*(?:,|or|and|/|&)\s*" + _MY, re.I)  # "by Summer 2027 or Summer 2028"


def _ym(month, year):
    if not month:
        return (int(year), None)
    key = month.split("/")[0].strip().lower()  # "May/June" -> May
    m = SEASON_MONTH.get(key) or MONTHS[key[:3]]
    return (int(year), m)


def _also(sentence, m):
    """A second date right after a match: "by Summer 2027 or Summer 2028"."""
    extra = RE_ALSO.match(sentence[m.end():])
    return [_ym(extra[1], extra[2])] if extra else []


def _grad_check(sentence):
    """(True / False / None, quote): does GRADUATION fit the window this sentence states?"""
    g = GRADUATION
    m = RE_GRAD_RANGE.search(sentence)
    if m:
        lo, hi = _ym(m[1], m[2]), _ym(m[3], m[4])
        lo = (lo[0], lo[1] or 1)
        hi = (hi[0], hi[1] or 12)
        internship_term = lo[0] == hi[0] == YEAR and lo[1] >= 5 and hi[1] <= 9
        if lo <= hi and hi[0] - lo[0] <= 3 and not internship_term:
            return lo <= g <= hi, m.group(0)
    m = RE_GRAD_BEFORE.search(sentence)
    if m:
        hi = max((y, mo or 12) for y, mo in [_ym(m[2], m[3])] + _also(sentence, m))
        return (g <= hi if m[1].lower().startswith(("by", "no", "on")) else g < hi), m.group(0)
    m = RE_GRAD_AFTER.search(sentence)
    if m:
        lo = min((y, mo or 1) for y, mo in [_ym(m[2], m[3])] + _also(sentence, m))
        return (g >= lo if m[1].lower().startswith(("no", "on")) else g > lo), m.group(0)
    m = RE_GRAD_IN.search(sentence)
    if m:
        points = [_ym(m[1], m[2])] + ([_ym(m[3], m[4])] if m[4] else [])
        after = sentence[m.end():m.end() + 20]
        if RE_OR_LATER.search(after):  # "December 2027 or later"
            fits = g >= min((y, mo or 1) for y, mo in points)
        elif RE_OR_EARLIER.search(after):
            fits = g <= max((y, mo or 12) for y, mo in points)
        else:
            fits = any(y == g[0] and (mo is None or abs(mo - g[1]) <= 1) for y, mo in points)
        return fits, m.group(0)
    return None, ""


def requirements(raw):
    """Who may apply, from posting HTML or text."""
    ok_stmt, grad_stmt, pref_stmt = [], [], []
    grad_votes = []
    for s in statements(raw):
        if not RE_ANCHOR.search(s):
            continue
        ug = [m for m in RE_UG.finditer(s)
              if not (RE_PRIOR_BEFORE.search(s[max(0, m.start() - 40):m.start()])
                      and "pursu" not in s[max(0, m.start() - 40):m.start()].lower())
              and not RE_PRIOR_AFTER.search(s[m.end():m.end() + 80])]
        gr = list(RE_GR.finditer(s))
        if ug or (RE_GENERIC.search(s) and not gr):
            ok_stmt.append(s)
        elif gr:
            soft = any(RE_SOFT.search(s[max(0, m.start() - 60):m.end() + 30]) for m in gr)
            (pref_stmt if soft else grad_stmt).append(s)
        verdict, quote = _grad_check(s)
        if verdict is not None:
            grad_votes.append((verdict, quote))
    if ok_stmt:
        degree, dq = "undergrad-ok", ok_stmt[0]
    elif grad_stmt:
        degree, dq = "grad-only", grad_stmt[0]
    else:
        degree, dq = "unknown", ""
    grad_ok = None
    if grad_votes:
        good = [q for v, q in grad_votes if v]
        grad_ok = bool(good)
        gq = good[0] if good else grad_votes[0][1]
    else:
        gq = ""
    return {
        "degree": degree, "degree_quote": dq[:200],
        # "Ideal candidates are PhD students" next to "we may consider bachelor's": open, but a long shot.
        "grad_preferred": degree == "undergrad-ok" and bool(pref_stmt),
        "grad_ok": grad_ok, "grad_quote": re.sub(r"\s+", " ", gq)[:120],
    }


# ----------------------------------------------------------------- fetching

class Gone(Exception):
    """The board says this posting no longer exists."""


async def _lever(job, net):
    posting_id = job.id.rsplit(":", 1)[-1]
    data = await net.get_json(f"https://api.lever.co/v0/postings/{job.company_slug}/{posting_id}?mode=json")
    lists = " ".join("<p>" + str(x.get("text", "")) + "</p>" + str(x.get("content", "")) for x in data.get("lists") or [])
    return "\n".join([str(data.get("descriptionPlain") or data.get("description") or ""), lists,
                      str(data.get("additionalPlain") or data.get("additional") or "")])


_ashby_boards = {}


async def _ashby(job, net):
    slug = job.company_slug
    if slug not in _ashby_boards:
        _ashby_boards[slug] = asyncio.ensure_future(
            net.get_json(f"https://api.ashbyhq.com/posting-api/job-board/{slug}"))
    board = await _ashby_boards[slug]
    posting_id = job.id.rsplit(":", 1)[-1]
    for j in board.get("jobs") or []:
        if j.get("id") == posting_id:
            return str(j.get("descriptionHtml") or j.get("descriptionPlain") or "")
    raise Gone()


async def _rippling(job, net):
    posting_id = job.id.rsplit(":", 1)[-1]
    data = await net.get_json(
        f"https://api.rippling.com/platform/api/ats/v1/board/{job.company_slug}/jobs/{posting_id}")
    desc = data.get("description") or ""
    if isinstance(desc, dict):
        desc = "\n".join(str(v or "") for v in desc.values())
    return str(desc)


async def _amazon(job, net):
    posting_id = job.id.rsplit(":", 1)[-1]
    data = await net.get_json("https://www.amazon.jobs/en/search.json",
                              params={"base_query": posting_id, "result_limit": 10})
    for j in data.get("jobs") or []:
        if str(j.get("id_icims")) == posting_id:
            return "\n".join(str(j.get(k) or "") for k in
                             ("description", "basic_qualifications", "preferred_qualifications"))
    return None


async def _oracle(job, net):
    if not enrich._ORACLE_RE.match(job.url):
        return None
    text = await enrich._oracle(job, net)
    if text is None:  # the board answers, but has no published requisition with this id
        raise Gone()
    return text


FETCHERS = dict(enrich._FETCHERS, lever=_lever, ashby=_ashby, rippling=_rippling, amazon=_amazon, oracle=_oracle)


def candidates():
    """Open postings shortlist.py could hand out, as Job objects."""
    api = json.load(open(shortlist.JOBS, encoding="utf-8"))["jobs"]
    store = json.load(open(STORE, encoding="utf-8"))
    today = datetime.now().date()
    jobs = []
    for j in api:
        if not shortlist.classify(j) or shortlist.why_ineligible(j, today):
            continue
        rec = store.get(j["id"]) or {}
        jobs.append(Job(id=j["id"], source=j.get("source") or rec.get("source") or "",
                        company=j["company"], company_slug=rec.get("company_slug") or "",
                        title=j["title"], location=j.get("location") or "", url=j["url"]))
    return jobs


async def read_one(job, net):
    """("ok", html) / ("gone", None) / ("error", None)."""
    fetcher = FETCHERS.get(job.source)
    if fetcher is None:
        return "error", None
    try:
        html = await fetcher(job, net)
    except Gone:
        return "gone", None
    except httpx.HTTPStatusError as exc:
        code = exc.response.status_code
        if code in (404, 410):
            return "gone", None
        # Workday answers 403 "S22 permission denied" for a posting that was taken down.
        if job.source == "workday" and code == 403 and '"S22"' in exc.response.text:
            return "gone", None
        return "error", None
    except Exception:  # noqa: BLE001  a dead page must not stop the run
        return "error", None
    return ("ok", html) if html and html.strip() else ("error", None)


async def fetch_all(jobs):
    limiter = HostLimiter(4)
    common = dict(timeout=httpx.Timeout(20.0, connect=10.0), follow_redirects=True,
                  headers={"User-Agent": "Mozilla/5.0 (internship-shortlist; personal use)"})
    async with httpx.AsyncClient(**common) as client:
        net = Net(client, limiter)
        results = await asyncio.gather(*(read_one(j, net) for j in jobs))
    return {j.id: r for j, r in zip(jobs, results)}


def load_cache():
    if not os.path.exists(CACHE):
        return {}
    with open(CACHE, encoding="utf-8") as fh:
        return json.load(fh)


def main():
    cache = load_cache()
    if "--report" in sys.argv:
        live = [e for e in cache.values() if e.get("status") == "ok"]
        print(f"{len(cache)} postings cached; {len(live)} read; "
              f"{sum(1 for e in cache.values() if e.get('status') == 'gone')} taken down; "
              f"{sum(1 for e in live if e.get('start'))} state dates; "
              f"{sum(1 for e in live if e.get('degree') == 'grad-only')} PhD / Master's only; "
              f"{sum(1 for e in live if e.get('grad_ok') is False)} graduation window misses May 2028.")
        return 0

    stamp = datetime.now().strftime("%Y-%m-%d")
    jobs = candidates()
    open_ids = {j.id for j in jobs}
    todo = [j for j in jobs
            if (cache.get(j.id) or {}).get("checked") != stamp or (cache.get(j.id) or {}).get("v") != VERSION]
    print(f"term_dates: {len(jobs)} candidate postings, {len(todo)} to read.")
    if todo:
        results = asyncio.run(fetch_all(todo))
        for job in todo:
            status, raw = results[job.id]
            old = cache.get(job.id) or {}
            if status == "error":
                if not old:
                    cache[job.id] = {"title": job.title, "status": "unread", "has_text": False}
                continue  # keep the last good facts; tried again next run
            entry = {"checked": stamp, "v": VERSION, "title": job.title, "status": status,
                     "has_text": status == "ok"}
            if status == "ok":
                entry.update(extract(job.title + ". " + sponsorship.strip_html(raw)))
                entry.update(requirements(raw))
            cache[job.id] = entry
    cache = {i: e for i, e in cache.items() if i in open_ids}  # closed postings leave the cache
    with open(CACHE, "w", encoding="utf-8") as fh:
        json.dump(cache, fh, indent=1, sort_keys=True)
    counts = {s: sum(1 for e in cache.values() if e.get("status") == s) for s in ("ok", "gone", "unread")}
    print(f"term_dates: {counts['ok']} read, {counts['gone']} taken down, {counts['unread']} unreadable; "
          f"{sum(1 for e in cache.values() if e.get('start'))} state dates, "
          f"{sum(1 for e in cache.values() if e.get('degree') == 'grad-only')} PhD / Master's only.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
