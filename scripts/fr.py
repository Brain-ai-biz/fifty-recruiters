#!/usr/bin/env python3
"""50 מגייסים: המנוע. פייתון 3.8 ומעלה, בלי שום התקנה.

  python3 fr.py start CV_FILE [--job AD_FILE_OR_TEXT]   פותח הרצה וקורא את קורות החיים
  python3 fr.py start --sample                          הרצה על קורות החיים לדוגמה
  python3 fr.py panel 1|2                               בודק את הסימנים ומחשב את 50 ההחלטות
  python3 fr.py lock                                    נעילת עובדות: השכתוב לא המציא כלום
  python3 fr.py build [--open]                          דף התוצאה, קובץ Word וקובץ להדפסה
  python3 fr.py status                                  מה נעשה ומה הצעד הבא

כל פקודה מקבלת --run DIR. בלי זה: ההרצה האחרונה תחת ./fifty-recruiters-output
"""
import argparse
import base64
import datetime as dt
import html
import json
import re
import shutil
import subprocess
import sys
import unicodedata
import webbrowser
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
KIT = HERE.parent
REF = KIT / "reference"
ASSETS = KIT / "assets"
OUT_ROOT = Path.cwd() / "fifty-recruiters-output"

TEXT_SIGNALS = ["headline", "summary", "numbers", "numbers_top", "action", "order", "dates", "short",
                "bullets", "skills", "keywords", "contact", "clean", "relevant_first"]
FACT_SIGNALS = ["years", "domain", "stability", "english", "degree", "link", "lead"]
ALL_SIGNALS = TEXT_SIGNALS + FACT_SIGNALS
SCRIPT_SIGNALS = {"short", "contact", "keywords"}
SHORT_WORDS = 650

LABEL = {
    "headline": "כותרת ברורה בראש", "summary": "תקציר קצר בהתחלה", "numbers": "תוצאות עם מספרים",
    "numbers_top": "מספר חזק כבר למעלה", "action": "מה נעשה ומה יצא מזה", "order": "מהחדש לישן",
    "dates": "תאריכים לכל תפקיד", "short": "עד שני עמודים", "bullets": "שורות קצרות שקל לסרוק",
    "skills": "כלים במקום אחד", "keywords": "מילות המפתח של התפקיד", "contact": "טלפון ומייל",
    "clean": "בלי פרטים מיותרים", "relevant_first": "הרלוונטי ראשון", "years": "שנות ניסיון",
    "domain": "ניסיון בתחום", "stability": "יציבות", "english": "אנגלית גבוהה", "degree": "השכלה",
    "link": "קישור לפרופיל או לעבודות", "lead": "ניהול אנשים",
}
NO = {
    "headline": "לא הבנתי מהשורה הראשונה לאיזה תפקיד זה מכוון.",
    "summary": "חסרות לי שתיים-שלוש שורות בהתחלה שאומרות מה מביאים.",
    "numbers": "לא ראיתי תוצאות עם מספרים, רק תיאור תפקיד.",
    "numbers_top": "המספרים הטובים קבורים למטה, ובשש שניות לא מגיעים אליהם.",
    "action": "יש כאן רשימת אחריות, לא מה נעשה בפועל ומה יצא מזה.",
    "order": "הניסיון מתחיל מהתפקיד הישן, ואת העדכני צריך לחפש.",
    "dates": "לא ברור מהתאריכים כמה זמן היה בכל מקום.",
    "short": "ארוך מדי בשביל קריאה ראשונה.",
    "bullets": "פסקאות ארוכות, קשה לסרוק אותן.",
    "skills": "הכלים מפוזרים בטקסט ואין מקום אחד לראות אותם.",
    "keywords": "לא מצאתי את מילות המפתח של התפקיד.",
    "contact": "לא מצאתי מהר טלפון ומייל.",
    "clean": "יש פרטים אישיים שלא צריך, והם תופסים את המקום של העיקר.",
    "relevant_first": "הניסיון הכי רלוונטי למשרה לא בולט ראשון.",
    "years": "חסרות לי שנות ניסיון למשרה הזו.",
    "domain": "אין ניסיון בתחום של המשרה, ואצלי זה תנאי.",
    "stability": "יותר מדי תפקידים קצרים.",
    "english": "לא ראיתי אנגלית ברמה גבוהה, ואצלנו זה תנאי.",
    "degree": "חסרה לי השכלה רלוונטית.",
    "link": "אין קישור לפרופיל או לתיק עבודות.",
    "lead": "לא ראיתי ניסיון בניהול אנשים, ובמשרה הזו זה תנאי.",
}
YES = {
    "headline": "בשורה הראשונה הבנתי בדיוק לאן זה מכוון.",
    "summary": "התקציר למעלה אמר לי מה מביאים, בשלוש שורות.",
    "numbers": "יש תוצאות עם מספרים, לא רק תיאור תפקיד.",
    "numbers_top": "כבר בחצי העליון ראיתי מספר שמשכנע.",
    "action": "רואים מה נעשה בפועל ומה יצא מזה.",
    "order": "התפקיד העדכני ראשון, בדיוק איפה שמסתכלים.",
    "dates": "הרצף ברור, עם תאריכים לכל תפקיד.",
    "short": "קצר וממוקד, קראתי הכול.",
    "bullets": "שורות קצרות, סרקתי בשניות.",
    "skills": "כל הכלים במקום אחד.",
    "keywords": "מילות המפתח של התפקיד מופיעות.",
    "contact": "הטלפון והמייל מול העיניים.",
    "clean": "מסמך נקי, רק מה שחשוב.",
    "relevant_first": "הניסיון הרלוונטי למשרה בולט ראשון.",
    "years": "יש בדיוק את שנות הניסיון שחיפשתי.",
    "domain": "יש ניסיון בתחום שלנו.",
    "stability": "תפקידים יציבים, לא קפיצות.",
    "english": "יש אנגלית ברמה גבוהה.",
    "degree": "יש השכלה מתאימה.",
    "link": "יש קישור, ונכנסתי לראות.",
    "lead": "יש ניסיון בניהול אנשים.",
}
CALL = {"f": ("הייתי מתקשרת", "לא הייתי מתקשרת"), "m": ("הייתי מתקשר", "לא הייתי מתקשר")}
DISCLAIMER = ("זו סימולציה: 50 פרופילי גיוס בדויים שמופעלים על ידי מודל AI, לא מגייסים אמיתיים. "
              "התוצאה היא כיוון לשיפור, לא הבטחה לראיון או לעבודה.")


# ---------------------------------------------------------------- טקסט
def norm(s):
    s = unicodedata.normalize("NFC", s or "")
    s = re.sub(r"[‎‏‪-‮⁦-⁩﻿­]", "", s)
    s = s.translate(str.maketrans({"״": '"', "“": '"', "”": '"', "„": '"', "׳": "'", "’": "'", "‘": "'", "`": "'",
                                   "–": "-", "—": "-", "‐": "-", "‑": "-", "−": "-", "•": " ", "·": " ", " ": " "}))
    return re.sub(r"\s+", " ", s).strip()


def loose(s):
    """להשוואת ציטוטים: בלי רווחים ובלי סימני פיסוק, כדי ששבירת שורה או נקודה לא יפילו ציטוט אמיתי."""
    return re.sub(r"[\s\"'.,;:!?()\[\]{}\-–—|/\\*•·_]+", "", norm(s)).lower()


def in_text(quote, text_loose):
    q = loose(quote)
    return bool(q) and q in text_loose


def hebrew_reversed(text):
    """טקסט מ-PDF שיצא הפוך: אות סופית בתחילת מילה במקום בסופה."""
    words = re.findall(r"[א-ת]{2,}", text)
    if len(words) < 20:
        return False
    start = sum(1 for w in words if w[0] in "ךםןףץ")
    end = sum(1 for w in words if w[-1] in "ךםןףץ")
    return start > end * 1.5 and start >= 5


def unreverse(text):
    out = []
    for line in text.splitlines():
        r = line[::-1]
        r = re.sub(r"[A-Za-z0-9@._%:/+#&=?-]+", lambda m: m.group(0)[::-1], r)
        out.append(r)
    return "\n".join(out)


def read_docx(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    paras = []
    for p in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S):
        p = re.sub(r"<w:tab/>", "\t", p)
        p = re.sub(r"<w:br[^>]*/>", "\n", p)
        parts = re.findall(r"<w:t(?: [^>]*)?>(.*?)</w:t>|(\t)|(\n)", p, re.S)
        line = "".join(a or b or c for a, b, c in parts)
        paras.append(html.unescape(line))
    return "\n".join(paras)


def read_pdf(path):
    try:
        from pypdf import PdfReader  # לא חובה. אם מותקן, משתמשים בו
        return "\n".join((pg.extract_text() or "") for pg in PdfReader(str(path)).pages), "pypdf"
    except Exception:
        pass
    if shutil.which("pdftotext"):
        r = subprocess.run(["pdftotext", "-layout", str(path), "-"], capture_output=True, text=True)
        if r.returncode == 0:
            return r.stdout, "pdftotext"
    return "", None


def extract(path):
    ext = path.suffix.lower()
    if ext == ".docx":
        return read_docx(path), "docx"
    if ext in (".txt", ".md", ".text"):
        return path.read_text(encoding="utf-8", errors="replace"), "text"
    if ext == ".pdf":
        return read_pdf(path)
    if ext == ".doc":
        if shutil.which("textutil"):  # מק
            r = subprocess.run(["textutil", "-convert", "txt", "-stdout", str(path)], capture_output=True, text=True)
            if r.returncode == 0:
                return r.stdout, "textutil"
        return "", None
    return "", None


# ---------------------------------------------------------------- הרצה
def run_dir(a, must=True):
    if getattr(a, "run", None):
        return Path(a.run).expanduser().resolve()
    if OUT_ROOT.exists():
        runs = sorted([p for p in OUT_ROOT.iterdir() if p.is_dir()], key=lambda p: p.stat().st_mtime)
        if runs:
            return runs[-1]
    if must:
        sys.exit("לא מצאתי הרצה. מתחילים ב: python3 fr.py start <קובץ קורות החיים>")
    return None


def load(p):
    try:
        return json.loads(Path(p).read_text(encoding="utf-8"))
    except FileNotFoundError:
        sys.exit(f"חסר קובץ: {p}")
    except json.JSONDecodeError as e:
        sys.exit(f"הקובץ {Path(p).name} אינו JSON תקין: {e}")


def save(p, obj):
    Path(p).write_text(json.dumps(obj, ensure_ascii=False, indent=1), encoding="utf-8")


def cmd_start(a):
    if a.sample:
        src = REF / "sample" / "cv-sample.docx"
        job = REF / "sample" / "job-sample.txt"
    else:
        if not a.cv:
            sys.exit("צריך קובץ קורות חיים: python3 fr.py start <קובץ> (או --sample לדוגמה)")
        src = Path(a.cv).expanduser()
        job = a.job
        if not src.exists():
            sys.exit(f"לא מצאתי את הקובץ: {src}")
    stamp = dt.datetime.now().strftime("%Y-%m-%d-%H%M%S")
    run = Path(a.run).expanduser().resolve() if a.run else OUT_ROOT / stamp
    run.mkdir(parents=True, exist_ok=True)
    shutil.copy(src, run / ("original" + src.suffix.lower()))
    text, how = extract(src)
    if how == "pdftotext" or how == "pypdf":
        text = re.sub(r"[\u200e\u200f\u202a-\u202e\u2066-\u2069]", "", text)
        text = "\n".join(re.sub(r"[ \t]{2,}", " ", ln).strip() for ln in text.splitlines())
        text = re.sub(r"\n{3,}", "\n\n", text)
    if text and hebrew_reversed(text):
        text = unreverse(text)
        how = f"{how}, תוקן כיוון"
    if job:
        jp = Path(str(job)).expanduser()
        is_file = len(str(job)) < 400 and "\n" not in str(job) and jp.is_file()
        ad = jp.read_text(encoding="utf-8", errors="replace") if is_file else str(job)
        (run / "job_ad.txt").write_text(ad.strip() + "\n", encoding="utf-8")
    print(f"תיקיית ההרצה: {run}")
    if text.strip():
        (run / "cv_v1.txt").write_text(text.strip() + "\n", encoding="utf-8")
        print(f"קורות החיים נקראו ({how}): {len(text.split())} מילים -> {run / 'cv_v1.txt'}")
        print("בדוק שהטקסט נקרא נכון (עברית מימין לשמאל, בלי מילים הפוכות). אם לא: כתוב את cv_v1.txt בעצמך מתוך הקובץ.")
    else:
        print("לא הצלחתי לחלץ טקסט מהקובץ הזה אוטומטית.")
        print(f"קרא את הקובץ {run / ('original' + src.suffix.lower())} בעצמך וכתוב את הטקסט מילה במילה ל-{run / 'cv_v1.txt'}.")
        print("אם אין דרך לקרוא אותו: בקש לשמור את קורות החיים כקובץ Word (docx) או להדביק את הטקסט.")
    print("מודעת הדרושים נשמרה ב-job_ad.txt" if (run / "job_ad.txt").exists() else "אין מודעת דרושים: התפקיד יוסק מקורות החיים.")
    print("הצעד הבא: לכתוב job.json ו-signals_v1.json לפי reference/signals.md, ואז: python3 fr.py panel 1")


# ---------------------------------------------------------------- סימנים
YEAR = re.compile(r"^(19|20)\d\d$")


def has_metric(q):
    toks = re.findall(r"\d[\d,.]*%?", q)
    return any(not YEAR.match(t.replace(",", "").rstrip("%")) for t in toks)


def check_signals(run, n):
    text = (run / f"cv_v{n}.txt").read_text(encoding="utf-8")
    tl = loose(text)
    job = load(run / "job.json")
    raw = load(run / f"signals_v{n}.json")
    sig = raw.get("signals", raw)
    errs, warn, out = [], [], {}
    for k in ALL_SIGNALS:
        if k in SCRIPT_SIGNALS or (n == 2 and k in FACT_SIGNALS):
            continue
        v = sig.get(k)
        if not isinstance(v, dict) or "met" not in v:
            errs.append(f"חסר הסימן {k}")
            continue
        met = bool(v["met"])
        quotes = v.get("quote") or []
        if isinstance(quotes, str):
            quotes = [quotes]
        if met:
            if not quotes:
                errs.append(f"{k}: מסומן כקיים בלי ציטוט")
            bad = [q for q in quotes if not in_text(q, tl)]
            if bad:
                errs.append(f"{k}: הציטוט לא נמצא מילה במילה בקורות החיים: «{bad[0][:70]}»")
            if k == "numbers" and sum(1 for q in quotes if has_metric(q)) < 2:
                errs.append("numbers: צריך לפחות שני ציטוטים שבכל אחד מספר של תוצאה (לא רק שנה)")
            if k == "numbers_top":
                pos = [loose(text).find(loose(q)) for q in quotes if has_metric(q) and in_text(q, tl)]
                if not pos:
                    errs.append("numbers_top: הציטוט צריך לכלול מספר של תוצאה")
                elif min(pos) > len(tl) * 0.5:
                    met = False
                    warn.append("numbers_top: המספר נמצא בחצי התחתון, ולכן הסימן לא מתקיים")
        out[k] = {"met": met, "quote": quotes, "why": v.get("why", "")}
    # סימנים שהסקריפט מחשב בעצמו
    words = len(re.findall(r"\S+", text))
    out["short"] = {"met": words <= SHORT_WORDS, "quote": [], "why": f"{words} מילים"}
    phone = re.search(r"(\+972|0)\s?-?\d{1,2}\s?-?\d{3}\s?-?\d{3,4}", norm(text))
    email = re.search(r"[\w.+-]+@[\w-]+\.[\w.]+", text)
    out["contact"] = {"met": bool(phone and email), "quote": [], "why": "טלפון ומייל" if phone and email else "חסר טלפון או מייל"}
    kws = [k for k in job.get("keywords", []) if str(k).strip()]
    found = [k for k in kws if loose(k) in tl]
    ratio = len(found) / len(kws) if kws else 0
    out["keywords"] = {"met": ratio >= 0.6, "quote": found, "why": f"{len(found)} מתוך {len(kws)} מילות מפתח"}
    if len(kws) != 8:
        errs.append(f"job.json: צריך בדיוק 8 מילות מפתח ב-keywords (יש {len(kws)}). ראו reference/signals.md סעיף 1")
    if n == 2:  # עובדות לא משתנות בשכתוב
        f1 = load(run / "signals_v1.final.json")["signals"]
        for k in FACT_SIGNALS:
            if k in f1:
                out[k] = dict(f1[k], why="ננעל מהסבב הראשון: עובדה, לא ניסוח")
    return out, errs, warn, {"words": words, "keywords_found": found, "keywords": kws}


def decide(rec, sig):
    miss_must = [k for k in rec["must"] if not sig[k]["met"]]
    got_want = [k for k in rec["want"] if sig[k]["met"]]
    miss_want = [k for k in rec["want"] if not sig[k]["met"]]
    call = not miss_must and len(got_want) >= rec["min"]
    if call:
        drivers = [k for k in rec["must"]] + got_want
    else:
        drivers = miss_must or miss_want
    return call, drivers


def template_reason(rec, call, drivers):
    yes, no = CALL.get(rec["g"], CALL["m"])
    if call:
        return f"{yes}. {YES[drivers[0]] if drivers else 'הכול כאן.'}"
    return f"{no}. {NO[drivers[0]] if drivers else 'משהו לא הסתדר לי.'}"


def cmd_panel(a):
    run = run_dir(a)
    n = a.round
    if n == 2 and not (run / "cv_v2.txt").exists():
        sys.exit("אין cv_v2.txt. קודם כותבים cv_v2.json ומריצים: python3 fr.py lock")
    if n == 2 and not (run / "signals_v1.final.json").exists():
        sys.exit("קודם מריצים את הסבב הראשון: python3 fr.py panel 1")
    sig, errs, warn, meta = check_signals(run, n)
    for w in warn:
        print("הערה:", w)
    if errs:
        print("יש לתקן את signals_v%d.json לפני שממשיכים:" % n)
        for e in errs:
            print("  -", e)
        sys.exit(1)
    save(run / f"signals_v{n}.final.json", {"signals": sig, "meta": meta})
    recs = load(REF / "recruiters.json")["recruiters"]
    custom = {}
    rp = run / f"reasons_{n}.json"
    if rp.exists():
        custom = {str(k): v for k, v in load(rp).items()}
    rows, bad = [], []
    for r in recs:
        call, drivers = decide(r, sig)
        reason = template_reason(r, call, drivers)
        c = custom.get(str(r["id"]))
        if c:
            c = norm(c)
            yes, no = CALL.get(r["g"], CALL["m"])
            ok_prefix = c.startswith(yes) if call else c.startswith(no)
            if call and c.startswith(no):
                ok_prefix = False
            if ok_prefix and len(c.split()) <= 28 and "—" not in c and "–" not in c:
                reason = c
            else:
                bad.append(r["id"])
        rows.append({"id": r["id"], "call": call, "drivers": drivers, "reason": reason})
    save(run / f"round{n}.json", {"round": n, "calls": sum(x["call"] for x in rows), "rows": rows})
    k = sum(x["call"] for x in rows)
    print(f"סבב {n}: {k} מתוך 50 היו מתקשרים.")
    met = [LABEL[s] for s in ALL_SIGNALS if sig[s]["met"]]
    miss = [LABEL[s] for s in ALL_SIGNALS if not sig[s]["met"]]
    print("קיים:", ", ".join(met) or "-")
    print("חסר:", ", ".join(miss) or "-")
    if n == 2 and (run / "round1.json").exists():
        r1 = {x["id"]: x["call"] for x in load(run / "round1.json")["rows"]}
        flipped = [x["id"] for x in rows if x["call"] and not r1.get(x["id"])]
        print(f"נדלקו אחרי השכתוב: {len(flipped)}")
    print("\nההחלטות (מספר, שם, החלטה, הסיבה העיקרית):")
    byid = {r["id"]: r for r in recs}
    for x in rows:
        print(f"  {x['id']:>2} {byid[x['id']]['name']} ({byid[x['id']]['role']}, {byid[x['id']]['org']}): "
              f"{'כן' if x['call'] else 'לא'} | {', '.join(LABEL[d] for d in x['drivers'][:2])}")
    if custom:
        print(f"\nמשפטים מ-reasons_{n}.json: {len(custom) - len(bad)} התקבלו, {len(bad)} נפסלו.")
    if bad:
        print(f"\nמשפטים שנפסלו ב-reasons_{n}.json (לא מתחילים בהחלטה הנכונה או ארוכים מ-28 מילים), הוחלפו בנוסח קבוע: {bad}")
    if not custom:
        print(f"\nאופציונלי ומומלץ: כתוב reasons_{n}.json במשפט אחד לכל מגייס בקול שלו, ואז הרץ שוב את הפקודה הזו.")
    print("הצעד הבא:", "שכתוב לפי reference/rewrite.md ואז: python3 fr.py lock" if n == 1 else "python3 fr.py build --open")


# ---------------------------------------------------------------- נעילת עובדות
NUM = re.compile(r"\d[\d,.]*")
HEB = re.compile(r"[\u05d0-\u05ea]+(?:[\"'][\u05d0-\u05ea]+)?")
NUMW = {"שלוש", "שלושה", "ארבע", "ארבעה", "חמש", "חמישה", "שש", "שישה", "שבע", "שבעה", "שמונה", "תשע", "תשעה", "עשר",
        "עשרה", "עשרים", "שלושים", "ארבעים", "חמישים", "שישים", "שבעים", "שמונים", "תשעים", "מאה", "מאות", "אלף", "אלפי",
        "אלפים", "מיליון", "מיליוני", "מיליונים", "פי", "כפול", "פעמיים", "חצי", "רבע", "שליש", "עשרות", "אחוז", "אחוזים", "שניים", "שתיים"}
TEAMW = {"צוות", "צוותים", "עובדים", "כפופים", "כפופות", "עובדות", "מנהלים"}
INFLATE = {"הובלתי", "ניהלתי", "הקמתי", "יזמתי", "שיפרתי", "בניתי", "פיתחתי", "הגדלתי", "הכפלתי", "חסכתי", "גייסתי", "אחראית", "אחראי"}
STOP = {"של", "את", "עם", "על", "גם", "כל", "אני", "היא", "הוא", "זה", "זו", "או", "אבל", "כמו", "יותר", "בין", "אחרי", "לפני", "תוך", "מול", "דרך", "שלי", "היום", "כדי", "רק", "עוד", "מאז"}


def stems(w):
    out = {w}
    for i in range(1, 3):
        if len(w) - i >= 3 and all(c in "והבלמשכ" for c in w[:i]):
            out.add(w[i:])
    return out


def heb_words(text):
    return [w for w in HEB.findall(norm(text))]


def wordset(text):
    s = set()
    for w in heb_words(text):
        s |= stems(w)
    return s
LAT = re.compile(r"[A-Za-z][A-Za-z0-9+#&]*")


def cv2_items(cv):
    """כל הטקסטים בשכתוב, עם המקורות שלהם."""
    items = []
    def add(where, obj):
        if isinstance(obj, dict) and "text" in obj:
            items.append((where, obj.get("text", ""), obj.get("src") or []))
    add("headline", cv.get("headline"))
    add("summary", cv.get("summary"))
    for i, e in enumerate(cv.get("experience", [])):
        for j, b in enumerate(e.get("bullets", [])):
            add(f"experience[{i}].bullets[{j}]", b)
    for key in ("education", "languages", "extra", "military"):
        for i, b in enumerate(cv.get(key, []) or []):
            add(f"{key}[{i}]", b)
    sk = cv.get("skills") or {}
    if sk:
        items.append(("skills", ", ".join(sk.get("items", [])), sk.get("src") or []))
    return items


def cv2_text(cv):
    L = [cv.get("name", "")]
    if cv.get("headline"):
        L.append(cv["headline"]["text"])
    c = cv.get("contact", {})
    L.append(" | ".join(x for x in [c.get("phone"), c.get("email"), c.get("city"), c.get("link")] if x))
    if cv.get("summary"):
        L += ["", "תקציר", cv["summary"]["text"]]
    if cv.get("experience"):
        L += ["", "ניסיון"]
        for e in cv["experience"]:
            L.append(f"{e.get('title', '')} | {e.get('org', '')} | {e.get('dates', '')}")
            L += [f"- {b['text']}" for b in e.get("bullets", [])]
    sk = cv.get("skills") or {}
    if sk.get("items"):
        L += ["", "כלים ומיומנויות", ", ".join(sk["items"])]
    for key, title in (("education", "השכלה"), ("military", "שירות"), ("languages", "שפות"), ("extra", "עוד")):
        if cv.get(key):
            L += ["", title] + [b["text"] for b in cv[key]]
    return "\n".join(L).strip() + "\n"


def cmd_lock(a):
    run = run_dir(a)
    v1 = (run / "cv_v1.txt").read_text(encoding="utf-8")
    t1 = loose(v1)
    cv = load(run / "cv_v2.json")
    changes = load(run / "changes.json") if (run / "changes.json").exists() else []
    renames = cv.get("renames", [])
    v1n = norm(v1)
    nums1 = {x.replace(",", "").rstrip(".") for x in NUM.findall(v1n)}
    lat1 = {x.lower() for x in LAT.findall(v1n)}
    errs = []
    for r in renames:
        if not in_text(r.get("src", ""), t1):
            errs.append(f"renames: המקור של «{r.get('term')}» לא נמצא בקורות החיים המקוריים")
        else:
            lat1 |= {x.lower() for x in LAT.findall(r.get("term", ""))}
    v1w = wordset(v1)
    warns, neww = [], set()
    for where, text, src in cv2_items(cv):
        srcw = set()
        for x in src:
            srcw |= wordset(x)
        content = [w for w in heb_words(text) if len(w) >= 3 and w not in STOP]
        if src and content and not any(stems(w) & srcw for w in content):
            errs.append(f"{where}: אין אף מילה משותפת בין השורה למקור שלה. המקור חייב להיות מה שהשורה אומרת («{text[:50]}»)")
        for w in heb_words(text):
            st = stems(w)
            if st & NUMW and not st & v1w:
                errs.append(f"{where}: «{w}» הוא מספר במילים שלא מופיע במקור («{text[:50]}»)")
            if st & TEAMW and not st & srcw:
                errs.append(f"{where}: «{w}» (צוות או עובדים) לא מופיע במקור של השורה. ניהול אנשים לא מוסיפים בניסוח")
            if st & INFLATE and not st & srcw:
                warns.append(f"{where}: «{w}» לא כתוב במקור של השורה. ודא שזה לא מנפח («{text[:50]}»)")
            if len(w) >= 3 and w not in STOP and not st & v1w:
                neww.add(w)
        if not src:
            errs.append(f"{where}: אין src. כל שורה צריכה להצביע על המקור שלה בקורות החיים המקוריים")
        for s in src:
            if not in_text(s, t1):
                errs.append(f"{where}: המקור «{s[:60]}» לא נמצא מילה במילה בקורות החיים המקוריים")
        for x in NUM.findall(norm(text)):
            if x.replace(",", "").rstrip(".") not in nums1:
                errs.append(f"{where}: המספר {x} לא מופיע בקורות החיים המקוריים («{text[:60]}»)")
        for w in LAT.findall(text):
            if w.lower() not in lat1:
                errs.append(f"{where}: המילה {w} לא מופיעה במקור. אם זה שם אחר לאותו דבר, הוסף ל-renames עם מקור")
    for i, e in enumerate(cv.get("experience", [])):
        for f in ("title", "org"):
            if e.get(f) and loose(e[f]) not in t1:
                errs.append(f"experience[{i}].{f}: «{e[f]}» לא מופיע במקור. תפקיד ומקום עבודה לא משנים")
        for x in NUM.findall(norm(e.get("dates", ""))):
            if x.replace(",", "") not in nums1:
                errs.append(f"experience[{i}].dates: {x} לא מופיע במקור")
    c = cv.get("contact", {})
    for f in ("phone", "email", "link"):
        if c.get(f) and loose(c[f]) not in t1:
            errs.append(f"contact.{f}: «{c[f]}» לא מופיע במקור")
    t2 = loose(cv2_text(cv))
    for i, ch in enumerate(changes):
        if ch.get("before") and not in_text(ch["before"], t1):
            errs.append(f"changes[{i}].before לא נמצא במקור")
        if ch.get("after") and not in_text(ch["after"], t2):
            errs.append(f"changes[{i}].after לא נמצא בשכתוב")
    if not changes:
        errs.append("חסר changes.json: 3 עד 8 שינויים, לכל אחד before, after ו-why")
    if errs:
        print("הנעילה נכשלה. השכתוב חייב להישען רק על מה שכתוב במקור:")
        for e in errs[:40]:
            print("  -", e)
        sys.exit(1)
    (run / "cv_v2.txt").write_text(cv2_text(cv), encoding="utf-8")
    for w in warns:
        print("לבדיקה:", w)
    if neww:
        print("מילים שלא מופיעות במקור (עבור עליהן: שם של כלי, תחום, תואר או הישג שלא כתוב במקור יוצא מהשכתוב):")
        print("  " + ", ".join(sorted(neww)))
    print(f"הנעילה עברה: כל שורה מצביעה על המקור, אין מספר חדש, אין מקום עבודה או תפקיד חדש. {len(cv2_text(cv).split())} מילים.")
    if renames:
        print("שמות שהוחלפו (כדאי להראות למשתמש):", ", ".join(f"{r['src']} -> {r['term']}" for r in renames))
    print("הצעד הבא: signals_v2.json על cv_v2.txt, בעיניים חדשות, ואז: python3 fr.py panel 2")


# ---------------------------------------------------------------- Word
def _runs(text, bold=False, size=22, color=None):
    out = []
    for m in re.finditer(r"([A-Za-z][A-Za-z0-9.+#&/@:_\-]*(?:[ ]+[A-Za-z0-9][A-Za-z0-9.+#&/@:_\-]*)*)|([^A-Za-z]+)", text):
        lat, heb = m.group(1), m.group(2)
        t = html.escape(lat or heb, quote=False)
        rpr = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        rpr += ("<w:b/><w:bCs/>" if bold else "") + (f'<w:color w:val="{color}"/>' if color else "")
        rpr += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>' + ("" if lat else "<w:rtl/>")
        out.append(f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>')
    return "".join(out)


def _p(text, bold=False, size=22, color=None, after=80, before=0, border=False, ind=0):
    ppr = f'<w:bidi/><w:spacing w:before="{before}" w:after="{after}"/>'
    if ind:
        ppr += f'<w:ind w:start="{ind}" w:hanging="220"/>'
    if border:
        ppr += '<w:pBdr><w:bottom w:val="single" w:sz="8" w:space="2" w:color="5B8DEF"/></w:pBdr>'
    return f'<w:p><w:pPr>{ppr}</w:pPr>{_runs(text, bold, size, color)}</w:p>'


def write_docx(cv, path):
    body = [_p(cv.get("name", ""), True, 40, "0B1424", 40)]
    if cv.get("headline"):
        body.append(_p(cv["headline"]["text"], True, 26, "2F6FD6", 60))
    c = cv.get("contact", {})
    body.append(_p(" | ".join(x for x in [c.get("phone"), c.get("email"), c.get("city"), c.get("link")] if x), size=19, color="44506A", after=160))
    def section(t):
        body.append(_p(t, True, 24, "0B1424", 80, 200, border=True))
    if cv.get("summary"):
        section("תקציר")
        body.append(_p(cv["summary"]["text"], size=21))
    if cv.get("experience"):
        section("ניסיון")
        for e in cv["experience"]:
            body.append(_p(f"{e.get('title', '')}, {e.get('org', '')}", True, 22, "0B1424", 20, 100))
            if e.get("dates"):
                body.append(_p(e["dates"], size=19, color="44506A", after=40))
            for b in e.get("bullets", []):
                body.append(_p("•  " + b["text"], size=21, after=40, ind=360))
    sk = cv.get("skills") or {}
    if sk.get("items"):
        section("כלים ומיומנויות")
        body.append(_p(" · ".join(sk["items"]), size=21))
    for key, title in (("education", "השכלה"), ("military", "שירות"), ("languages", "שפות"), ("extra", "עוד")):
        if cv.get(key):
            section(title)
            for b in cv[key]:
                body.append(_p(b["text"], size=21, after=40))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
           + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000" w:header="0" w:footer="0" w:gutter="0"/><w:bidi/></w:sectPr>'
           '</w:body></w:document>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)


# ---------------------------------------------------------------- דף להדפסה
def cv_html(cv, fonts_css):
    e = lambda s: html.escape(s or "")
    c = cv.get("contact", {})
    parts = [f"<h1>{e(cv.get('name'))}</h1>"]
    if cv.get("headline"):
        parts.append(f"<p class='hl'>{e(cv['headline']['text'])}</p>")
    parts.append("<p class='ct'>" + " &nbsp;|&nbsp; ".join(f"<bdi>{e(x)}</bdi>" for x in [c.get("phone"), c.get("email"), c.get("city"), c.get("link")] if x) + "</p>")
    if cv.get("summary"):
        parts.append(f"<h2>תקציר</h2><p>{e(cv['summary']['text'])}</p>")
    if cv.get("experience"):
        parts.append("<h2>ניסיון</h2>")
        for x in cv["experience"]:
            parts.append(f"<div class='job'><div class='jt'><b>{e(x.get('title'))}</b>, {e(x.get('org'))}<span>{e(x.get('dates'))}</span></div><ul>"
                         + "".join(f"<li>{e(b['text'])}</li>" for b in x.get("bullets", [])) + "</ul></div>")
    sk = cv.get("skills") or {}
    if sk.get("items"):
        parts.append("<h2>כלים ומיומנויות</h2><p class='sk'>" + "".join(f"<span>{e(s)}</span>" for s in sk["items"]) + "</p>")
    for key, title in (("education", "השכלה"), ("military", "שירות"), ("languages", "שפות"), ("extra", "עוד")):
        if cv.get(key):
            parts.append(f"<h2>{title}</h2>" + "".join(f"<p>{e(b['text'])}</p>" for b in cv[key]))
    return f"""<!doctype html><html lang="he" dir="rtl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{e(cv.get('name'))} - קורות חיים</title><style>{fonts_css}
@page{{size:A4;margin:14mm}}
body{{font-family:Rubik,Arial,sans-serif;color:#0b1424;max-width:780px;margin:24px auto;padding:0 20px;line-height:1.55;font-size:15px}}
h1{{font-size:34px;margin:0;font-weight:800}} .hl{{color:#2f6fd6;font-weight:700;font-size:18px;margin:4px 0}}
.ct{{color:#44506a;font-size:13.5px;margin:0 0 14px}} h2{{font-size:16px;margin:18px 0 6px;padding-bottom:3px;border-bottom:2px solid #77CEBF}}
.jt{{display:flex;justify-content:space-between;gap:12px}} .jt span{{color:#44506a;font-size:13.5px;white-space:nowrap}}
ul{{margin:4px 0 10px;padding-inline-start:20px}} li{{margin:2px 0}} p{{margin:4px 0}}
.sk span{{display:inline-block;border:1px solid #c9d6ee;border-radius:99px;padding:1px 10px;margin:2px;font-size:13.5px}}
.bar{{position:sticky;top:0;background:#fff;padding:10px 0;text-align:center}} .bar button{{font:700 16px Rubik,Arial;padding:10px 22px;border-radius:12px;border:0;background:#0b1424;color:#fff;cursor:pointer}}
@media print{{.bar{{display:none}} body{{margin:0}}}}
</style></head><body><div class="bar"><button onclick="print()">שמירה כ-PDF או הדפסה</button></div>{''.join(parts)}</body></html>"""


# ---------------------------------------------------------------- דף התוצאה
def fonts_css():
    out = []
    for f, rng in (("rubik-hebrew.woff2", "U+0590-05FF,U+200C-2010,U+20AA,U+25CC,U+FB1D-FB4F"), ("rubik-latin.woff2", "U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+2000-206F,U+2074,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD")):
        p = ASSETS / f
        if p.exists():
            b = base64.b64encode(p.read_bytes()).decode()
            out.append(f"@font-face{{font-family:Rubik;src:url(data:font/woff2;base64,{b}) format('woff2');font-weight:300 900;unicode-range:{rng}}}")
    return "\n".join(out)


def cmd_build(a):
    run = run_dir(a)
    recs = load(REF / "recruiters.json")["recruiters"]
    job = load(run / "job.json")
    r1 = load(run / "round1.json")
    r2 = load(run / "round2.json") if (run / "round2.json").exists() else None
    cv = load(run / "cv_v2.json") if (run / "cv_v2.json").exists() else None
    if not (r2 and cv and (run / "cv_v2.txt").exists()):
        sys.exit("צריך את שני הסבבים ואת השכתוב הנעול. בדוק: python3 fr.py status")
    s1 = load(run / "signals_v1.final.json")["signals"]
    s2 = load(run / "signals_v2.final.json")["signals"]
    changes = load(run / "changes.json")
    css = fonts_css()
    write_docx(cv, run / "cv-new.docx")
    page = cv_html(cv, css)
    (run / "cv-new.html").write_text(page, encoding="utf-8")
    b1, b2 = {x["id"]: x for x in r1["rows"]}, {x["id"]: x for x in r2["rows"]}
    people = []
    for r in recs:
        people.append({"id": r["id"], "name": r["name"], "g": r["g"], "role": r["role"], "org": r["org"], "scan": r["scan"],
                       "c1": b1[r["id"]]["call"], "c2": b2[r["id"]]["call"],
                       "w1": b1[r["id"]]["reason"], "w2": b2[r["id"]]["reason"],
                       "d2": [LABEL[d] for d in b2[r["id"]]["drivers"]]})
    stayed = [p for p in people if not p["c2"]]
    gaps = {}
    for r in recs:
        if not b2[r["id"]]["call"]:
            d = (b2[r["id"]]["drivers"] or ["?"])[0]
            if d in LABEL:
                gaps[d] = gaps.get(d, 0) + 1
    data = {
        "name": cv.get("name", ""), "target": job.get("title", ""), "fromAd": bool(job.get("from_ad")),
        "n1": r1["calls"], "n2": r2["calls"], "people": people,
        "sig": [{"k": k, "label": LABEL[k], "v1": s1[k]["met"], "v2": s2[k]["met"], "fact": k in FACT_SIGNALS} for k in ALL_SIGNALS],
        "gaps": [{"label": LABEL[k], "n": v, "fact": k in FACT_SIGNALS, "k": k} for k, v in sorted(gaps.items(), key=lambda x: -x[1])],
        "changes": changes[:8], "v1": (run / "cv_v1.txt").read_text(encoding="utf-8")[:6000], "cv": cv,
        "cvHtml": page, "docx": base64.b64encode((run / "cv-new.docx").read_bytes()).decode(),
        "disclaimer": DISCLAIMER, "date": dt.date.today().strftime("%d.%m.%Y"),
        "file": re.sub(r"[^\w֐-׿-]+", "-", cv.get("name", "cv")).strip("-") or "cv",
    }
    tpl = (ASSETS / "wall.html").read_text(encoding="utf-8")
    tpl = tpl.replace("/*__FONTS__*/", css).replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False).replace("</", "<\\/"))
    (run / "wall.html").write_text(tpl, encoding="utf-8")
    print(f"מוכן. {r1['calls']} מתוך 50 בסבב הראשון, {r2['calls']} מתוך 50 אחרי השכתוב.")
    print(f"דף התוצאה: {run / 'wall.html'}")
    print(f"קורות החיים החדשים: {run / 'cv-new.docx'} (וורד) ו-{run / 'cv-new.html'} (לשמירה כ-PDF)")
    if stayed:
        print("נשארו כבויים בגלל:", ", ".join(f"{LABEL[k]} ({v})" for k, v in sorted(gaps.items(), key=lambda x: -x[1])[:4]))
    print(DISCLAIMER)
    if a.open:
        webbrowser.open((run / "wall.html").as_uri())


def cmd_status(a):
    run = run_dir(a, must=False)
    if not run:
        print("אין עדיין הרצה. הצעד הבא: python3 fr.py start <קובץ קורות החיים>")
        return
    steps = [("cv_v1.txt", "קורות החיים נקראו", "python3 fr.py start <קובץ>"),
             ("job.json", "התפקיד והמילים", "לכתוב job.json לפי reference/signals.md"),
             ("signals_v1.json", "סימנים, סבב 1", "לכתוב signals_v1.json לפי reference/signals.md"),
             ("round1.json", "סבב 1", "python3 fr.py panel 1"),
             ("cv_v2.json", "שכתוב", "לכתוב cv_v2.json ו-changes.json לפי reference/rewrite.md"),
             ("cv_v2.txt", "נעילת עובדות", "python3 fr.py lock"),
             ("signals_v2.json", "סימנים, סבב 2", "לכתוב signals_v2.json על cv_v2.txt"),
             ("round2.json", "סבב 2", "python3 fr.py panel 2"),
             ("wall.html", "דף התוצאה", "python3 fr.py build --open")]
    print(f"הרצה: {run}")
    for f, label, how in steps:
        ok = (run / f).exists()
        print(f"  {'V' if ok else ' '} {label}")
        if not ok:
            print(f"הצעד הבא: {how}")
            return
    print("הכול מוכן.")


def main():
    ap = argparse.ArgumentParser(description="50 מגייסים")
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("start")
    s.add_argument("cv", nargs="?")
    s.add_argument("--job")
    s.add_argument("--sample", action="store_true")
    s.add_argument("--run")
    p = sub.add_parser("panel")
    p.add_argument("round", type=int, choices=[1, 2])
    p.add_argument("--run")
    l = sub.add_parser("lock")
    l.add_argument("--run")
    b = sub.add_parser("build")
    b.add_argument("--open", action="store_true")
    b.add_argument("--run")
    st = sub.add_parser("status")
    st.add_argument("--run")
    a = ap.parse_args()
    {"start": cmd_start, "panel": cmd_panel, "lock": cmd_lock, "build": cmd_build, "status": cmd_status}[a.cmd](a)


if __name__ == "__main__":
    main()
