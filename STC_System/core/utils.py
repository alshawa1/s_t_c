"""
core/utils.py
─────────────
Shared helpers: date parsing, Arabic text utilities, constants, logging setup.
"""
import logging
import os
from datetime import date, datetime
from typing import Dict, List, Optional

import math
import re
import pandas as pd
from dateutil import parser as dateutil_parser

# ─── Logging ────────────────────────────────────────────────────────────────

LOG_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "logs")
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(LOG_DIR, "system.log"), encoding="utf-8"),
        logging.StreamHandler(),
    ],
)

logger = logging.getLogger("STC_OPS")


# ─── Date helpers ────────────────────────────────────────────────────────────

def get_today() -> date:
    """Return today's date."""
    return date.today()


def _is_nan(value) -> bool:
    """Check if value is null, None, NaN, NaT or blank string."""
    if value is None:
        return True
    if isinstance(value, float):
        return math.isnan(value)
    if isinstance(value, str):
        v = value.strip().lower()
        return v in ["", "nan", "nat", "null", "none"]
    return False


def normalize_date_to_iso(value, dayfirst: bool = True) -> str:
    """
    تحويل أي تاريخ إلى صيغة 'YYYY-MM-DD' بدقة متناهية.
    يدعم:
      - كائنات date و datetime
      - أرقام إكسيل التسلسلية (مثل 45510 و 45510.5)
      - صيغ ISO (YYYY-MM-DD و YYYY-MM-DD HH:MM:SS و YYYY-MM-DDTHH:MM:SS)
      - صيغ DD/MM/YYYY و MM/DD/YYYY حسب dayfirst (أو عند وضوح اليوم > 12)
    يرجع نص فارغ '' إذا كانت القيمة غير صالحة.
    """
    if _is_nan(value):
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    
    s = str(value).strip()
    if not s or s.lower() in ("none", "nan", "null", "nat", "-", "0", "0.0", "غير محدد"):
        return ""

    # 1. أرقام إكسيل التسلسلية (float / int / string)
    try:
        f = float(s)
        if 20000 <= f <= 75000:
            dt = pd.to_datetime(f, unit="D", origin="1899-12-30")
            return dt.strftime("%Y-%m-%d")
        elif 1980 <= f <= 2045:
            return f"{int(f):04d}-01-01"
    except (ValueError, OverflowError):
        pass

    # 2. صيغ ISO القياسية (مع إمكانية وجود وقت أو حرف T)
    m_iso = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)
    if m_iso:
        y, m, d = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
        if 1980 <= y <= 2045 and 1 <= m <= 12 and 1 <= d <= 31:
            return f"{y:04d}-{m:02d}-{d:02d}"

    # 3. إزالة الوقت من النص قبل الفحص
    s_date = re.split(r"[T\s]", s)[0].strip()

    # 4. صيغ DD/MM/YYYY أو MM/DD/YYYY
    m_dmy = re.match(r"^(\d{1,2})[-/\.](\d{1,2})[-/\.](\d{4})$", s_date)
    if m_dmy:
        part1, part2, y = int(m_dmy.group(1)), int(m_dmy.group(2)), int(m_dmy.group(3))
        if 1980 <= y <= 2045:
            # إذا كان الجزء الأول أكبر من 12 فهو يقيناً اليوم (DD/MM/YYYY)
            if part1 > 12 and 1 <= part2 <= 12 and 1 <= part1 <= 31:
                return f"{y:04d}-{part2:02d}-{part1:02d}"
            # إذا كان الجزء الثاني أكبر من 12 فهو يقيناً اليوم (MM/DD/YYYY)
            elif part2 > 12 and 1 <= part1 <= 12 and 1 <= part2 <= 31:
                return f"{y:04d}-{part1:02d}-{part2:02d}"
            # في حال اللبس (كلاهما <= 12): نستخدم dayfirst المكتشف للعمود
            elif dayfirst and 1 <= part1 <= 31 and 1 <= part2 <= 12:
                return f"{y:04d}-{part2:02d}-{part1:02d}"
            elif (not dayfirst) and 1 <= part1 <= 12 and 1 <= part2 <= 31:
                return f"{y:04d}-{part1:02d}-{part2:02d}"

    # 5. محاولة إضافية عبر dateutil
    try:
        dt = dateutil_parser.parse(s_date, dayfirst=dayfirst)
        if 1980 <= dt.year <= 2045:
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass

    return ""


def normalize_date_series_to_iso(values) -> List[str]:
    """
    توحيد عمود تواريخ كامل إلى صيغة 'YYYY-MM-DD' مع الكشف التلقائي الذكي
    لنمط العمود (هل هو DD/MM/YYYY أم MM/DD/YYYY) بناءً على القيم غير الملتبسة (> 12).
    """
    dmy_evidence = 0
    mdy_evidence = 0
    pat = re.compile(r"^(\d{1,2})[-/\.](\d{1,2})[-/\.](\d{4})")

    for v in values:
        if v is None or isinstance(v, (datetime, date)):
            continue
        s = str(v).strip()
        if not s:
            continue
        m = pat.match(s)
        if m:
            p1, p2 = int(m.group(1)), int(m.group(2))
            if p1 > 12 and 1 <= p2 <= 12:
                dmy_evidence += 1
            elif p2 > 12 and 1 <= p1 <= 12:
                mdy_evidence += 1

    prefer_dayfirst = (dmy_evidence >= mdy_evidence)
    return [normalize_date_to_iso(v, dayfirst=prefer_dayfirst) for v in values]


def parse_date(value) -> Optional[date]:
    """
    Safely parse any date-like value (datetime, str, Excel serial, None …)
    Returns a date object or None.
    """
    iso_str = normalize_date_to_iso(value)
    if not iso_str:
        return None
    try:
        parts = iso_str.split("-")
        return date(int(parts[0]), int(parts[1]), int(parts[2]))
    except Exception:
        return None


def days_since(value) -> Optional[int]:
    """Return number of days from *value* to today. None if unparsable."""
    d = parse_date(value)
    if d is None:
        return None
    return (get_today() - d).days


def is_future_or_today(value) -> bool:
    """Return True if *value* is today or a future date."""
    d = parse_date(value)
    if d is None:
        return False
    return d >= get_today()


# ─── Arabic text helpers ─────────────────────────────────────────────────────

def clean_text(value) -> str:
    """Return stripped string or empty string."""
    if _is_nan(value):
        return ""
    return str(value).strip()


def is_empty(value) -> bool:
    """True if value is None / NaN / blank string."""
    return _is_nan(value)


def is_unclear(value) -> bool:
    """
    True if *value* is empty OR is a trivial / unclear note.
    Extend the GENERIC_NOTES list to tune detection.
    """
    if is_empty(value):
        return True
    text = clean_text(value).lower()
    GENERIC_NOTES = [
        "لا يوجد", "بدون", "-", ".", "..", "...", "لايوجد",
        "تم التواصل", "تواصل", "ok", "okay", "n/a", "na",
    ]
    return text in GENERIC_NOTES


# ─── Task / file configuration ───────────────────────────────────────────────

# Human-readable task names (Arabic + English)
TASK_NAMES = {
    1: "اخطاء النظام — System Errors",
    2: "التوصل وعدم التوصل — Contact Status",
    3: "الاهمال — Neglect Analysis",
    4: "السدادات والتسوية — Payments & Reconciliation",
    5: "الجدولة — Scheduling",
    6: "السحب والتدوير — Withdrawal & Rotation",
    7: "العملاء المستهدفة — Target Customers",
    8: "تشغيل جميع المهام — Full Operations Report",
}

# Logical file keys
MAIN_PORTFOLIO  = "main_portfolio"
PROMISE_PAY     = "promise_pay"
MAHARAH_PAY     = "maharah_pay"
COMPANY_PAY     = "company_pay"

# Files required per task
TASK_FILE_REQUIREMENTS = {
    1: [MAIN_PORTFOLIO, PROMISE_PAY],
    2: [MAIN_PORTFOLIO],
    3: [MAIN_PORTFOLIO],
    4: [MAIN_PORTFOLIO, MAHARAH_PAY, COMPANY_PAY],
    5: [MAIN_PORTFOLIO, MAHARAH_PAY],
    6: [MAIN_PORTFOLIO],
    7: [MAIN_PORTFOLIO, PROMISE_PAY, MAHARAH_PAY],
    8: [MAIN_PORTFOLIO, PROMISE_PAY, MAHARAH_PAY, COMPANY_PAY],
}

FILE_LABELS = {
    MAIN_PORTFOLIO: "المحفظة الموزعة (Main Portfolio)",
    PROMISE_PAY:    "وعود السداد (Promise To Pay)",
    MAHARAH_PAY:    "سداد مهارة (Maharah Payments)",
    COMPANY_PAY:    "سداد الشركة (Parent Company Payments)",
}

# ─── Required columns per file (minimum validation set) ──────────────────────
# NOTE: Validation uses ANY-match logic — the file passes if at least ONE
# column from any alternative group is found. This allows portfolios with
# different column names (e.g. رقم المدينية vs رقم المديونية) to load fine.

# Alternative column sets — each inner list = one "slot"; file is valid if
# at least ONE column from the slot exists.
FILE_REQUIRED_COLUMNS_ALT: Dict[str, List[List[str]]] = {
    MAIN_PORTFOLIO: [
        # At least one ID-like column
        ["رقم الحساب", "رقم المديونية", "رقم المدينية", "رقم الهوية", "الهوية",
         "الرقم الرئيسي", "Account No.", "Account No", "Customer ID", "Debt No."],
    ],
    PROMISE_PAY: [
        ["رقم الحساب", "رقم المديونية", "رقم المدينية", "رقم الهوية",
         "الرقم الرئيسي", "Account No."],
    ],
    MAHARAH_PAY: [
        ["رقم الحساب", "رقم المديونية", "EmpSadad_Id", "Account No.",
         "رقم الهوية", "الهوية"],
    ],
    COMPANY_PAY: [
        ["Account No.", "Account No", "Service No.", "رقم الحساب"],
    ],
}

# Keep old dict for compatibility (not used for validation any more)
FILE_REQUIRED_COLUMNS: Dict[str, List[str]] = {
    MAIN_PORTFOLIO: ["رقم الحساب", "رقم المديونية", "الحالة الرئيسية", "المشرف"],
    PROMISE_PAY:    ["رقم الحساب", "رقم المديونية", "تاريخ وعد السداد"],
    MAHARAH_PAY:    ["رقم الحساب", "رقم المديونية", "مبلغ السداد"],
    COMPANY_PAY:    ["Account No.", "Payment Amount", "Current Balance Due"],
}

# ─── Business constants ───────────────────────────────────────────────────────

NEGLECT_THRESHOLD_DAYS = 7  # days without follow-up → مهمل

# Statuses that indicate a fully-paid customer
PAID_STATUSES = {
    "مسدد", "مسدد كليا", "مسدد كلياً", "تم السداد",
    "سداد كامل", "سداد تام", "مسدد بالكامل",
    "paid", "full payment",
}

# Statuses that indicate successful contact
CONTACTED_STATUSES = {
    "وعد السداد", "وعد سداد", "مسدد", "مسدد جزئي",
    "سداد جزئي", "تم التواصل", "قسط", "اتفاق سداد",
    "تسوية", "جدولة", "إعادة جدولة", "تحويل",
    "تم الرد", "رجع الاتصال", "رد علينا",
}

# Statuses that indicate no contact
NOT_CONTACTED_STATUSES = {
    "لا يرد", "لايرد", "مغلق", "خارج التغطية",
    "لا يوجد رد", "لا يتكلم عربي", "رسالة صوتية",
    "خط معطل", "لا يتم الاتصال", "ارجاء",
    "مرفوض", "لم يرد", "لم نتمكن من التواصل",
}
