# -*- coding: utf-8 -*-
"""
modules/module14_portfolio_analytics.py
──────────────────────────────────────
اللوحة التنفيذية لتحليل المحفظة الاستراتيجي وعمر الدين (Strategic Portfolio Analytics):

1. ربط المحفظة بشيت التحصيل برقم المديونية (Debt ID).
2. حساب عدد العملاء بالهويات الفريدة فقط (Unique National IDs).
3. تحليل حالات التواصل (تم التوصل / لا يرد ومغلق / عدم التوصل).
4. كاشف توفر الأرقام الرئيسية (الرقم الرئيسي) ونسبها.
5. تحليل سنة التعثر (عمر الدين) من تاريخ آخر إلغاء والربط بمبالغ السداد.
6. تحليل التحصيل حسب الحالة الرئيسية وتحديد الحالة الأكثر سداداً.
7. تقييم أداء المشرفين والمحصلين وكيرف الأداء المزدوج (التغطية + التحصيل).
8. تصدير تقرير HTML & CSS تفاعلي فخم ومستقل قابل للطباعة والحفظ.
9. تصدير تقرير Excel منسق واحترافي.
"""

from __future__ import annotations
import io
import json
import logging
import re
from datetime import datetime, date
from typing import Dict, List, Optional, Any
import polars as pl
import pandas as pd

_log = logging.getLogger("PORTFOLIO_ANALYTICS")

_DEBT_COLS         = ["رقم المديونية", "رقم المديونيه", "رقم العقد", "رقم الحساب", "المديونية", "debt_id", "Debt ID", "Debt No."]
_ID_COLS           = ["رقم الهوية", "الهوية", "هوية العميل", "السجل المدني", "ID", "customer_id", "Customer ID"]
_DEBT_AMT_COLS     = ["مبلغ الميدونية", "مبلغ المديونية", "مبلغ المديونيه", "المديونية", "debt_amount", "amount", "قيمة المديونية"]
_REMAIN_AMT_COLS   = ["متبقي سداد موثق", "متبقى سداد موثق", "المتبقي الموثق", "المتبقي", "remaining_amount", "متبقي السداد", "remaining"]
_SUPERVISOR_COLS   = ["المشرف", "اسم المشرف", "supervisor", "Supervisor"]
_COLLECTOR_COLS    = ["المحصل", "اسم المحصل", "الموظف", "محصل", "collector", "Collector"]

_PHONE_COLS        = ["الرقم الرئيسي", "الرقم الرئيسى", "ارقام العميل", "أرقام العميل", "الهاتف", "رقم الهاتف", "الجوال", "رقم الجوال", "Service No.", "phone", "mobile"]
_PORTFOLIO_COLS    = ["المحافظ", "المحفظة", "المحفظه", "اسم المحفظة", "اسم المحفظه", "نوع المحفظة", "الحاوية", "اسم الحاوية", "نوع المديونية", "portfolio", "Portfolio", "bucket"]
_CANCEL_DATE_COLS  = [
    "تاريخ فصل الخدمة", "تاريخ فصل الخدمه",
    "تاريخ الغاء الخدمة", "تاريخ إلغاء الخدمة",
    "تاريخ اخر الغاء", "تاريخ آخر إلغاء", "تاريخ أخر الغاء", "تاريخ آخر الغاء",
    "سنة التعثر", "عمر الدين", "تاريخ إنتهاء الحساب", "تاريخ انتهاء الحساب",
    "end_date", "expiry_date", "account_end_date", "disconnect_date"
]

_CANCEL_DATE_COLS  = [
    "تاريخ اخر الغاء", "تاريخ آخر إلغاء", "تاريخ أخر الغاء", "تاريخ آخر الغاء",
    "تاريخ الغاء الخدمة", "تاريخ فصل الخدمة", "تاريخ فصل الخدمه",
    "تاريخ إنتهاء الحساب", "تاريخ انتهاء الحساب", "تاريخ الانتهاء", "تاريخ الإنتهاء",
    "end_date", "expiry_date", "account_end_date", "disconnect_date"
]

_FOLLOWUP_DATE_COLS = [
    "تاريخ المتابعة", "تاريخ المتابعه",
    "آخر متابعة للعميل", "أخر متابعة للعميل", "اخر متابعة للعميل",
    "آخر متابعة", "أخر متابعة", "اخر متابعة",
    "تاريخ آخر متابعة", "تاريخ أخر متابعة", "تاريخ اخر متابعة",
    "تاريخ التواصل", "تاريخ التغطية", "followup_date"
]

_STATUS_MAIN_COLS  = ["الحالة الرئيسية", "الحاله الرئيسيه", "الحالة", "main_status", "status"]
_STATUS_SUB_COLS   = ["الحالة الفرعية", "الحاله الفرعيه", "sub_status"]

_PAY_AMT_COLS      = ["مبلغ السداد", "قيمة السداد", "المبلغ", "amount", "Amount", "paid_amount", "السدادات الموثقة", "Payment Amount"]
_PAY_DATE_COLS     = ["تاريخ السداد", "تاريخ الدفع", "تاريخ الحركة", "التاريخ", "payment_date", "date", "Date", "Payment Date"]


def _norm_str(s: str) -> str:
    s = re.sub(r'[\s_]+', ' ', str(s or '')).strip().lower()
    s = re.sub(r'[إأآا]', 'ا', s)
    s = s.replace('ة', 'ه').replace('ى', 'ي')
    return s


def _detect_col(df: Any, candidates: List[str]) -> Optional[str]:
    cols = df.columns if hasattr(df, "columns") else []
    for c in candidates:
        if c in cols:
            return c
    norm_candidates = [_norm_str(cand) for cand in candidates]
    for c in cols:
        c_norm = _norm_str(c)
        for cand_norm in norm_candidates:
            if c_norm == cand_norm:
                return c
    for c in cols:
        c_norm = _norm_str(c)
        for cand_norm in norm_candidates:
            if len(cand_norm) >= 4 and cand_norm in c_norm:
                return c
    return None


def clean_id(val: Any) -> str:
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def to_clean_num(series: Any) -> pd.Series:
    if series is None:
        return pd.Series(dtype=float)
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0.0)
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.replace(" ", "", regex=False).str.replace("﷼", "", regex=False),
        errors="coerce"
    ).fillna(0.0)


def to_clean_num_scalar(val: Any) -> float:
    """تحويل أي قيمة عددية فردية إلى float بأمان تام"""
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace(",", "").replace(" ", "").replace("﷼", "").strip()
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def extract_year_smart(val: Any) -> str:
    """استخراج سنة التعثر من صيغ التواريخ المختلفة أو أرقام إكسيل"""
    if val is None or pd.isna(val):
        return "غير محدد"
    s = str(val).strip()
    if not s or s.lower() in ("none", "nan", "null", "-", "غير محدد", "0", "0.0"):
        return "غير محدد"
    
    # 1. رقم إكسيل تسلسلي أو سنة مباشرة
    try:
        f = float(s)
        if 20000 <= f <= 75000:
            dt = pd.to_datetime(f, unit="D", origin="1899-12-30")
            return str(dt.year)
        elif 1980 <= f <= 2035:
            return str(int(f))
    except:
        pass

    # 2. بحث عن سنة رباعية صريحة
    m = re.search(r"\b(19[89]\d|20[0-3]\d)\b", s)
    if m:
        return m.group(1)

    # 3. محاولة بالتحويل الزمني
    for dfmt in [True, False]:
        try:
            dt = pd.to_datetime(s.split(" ")[0], errors="coerce", dayfirst=dfmt)
            if pd.notna(dt) and 1980 <= dt.year <= 2035:
                return str(dt.year)
        except:
            pass

    return "غير محدد"


class StrategicPortfolioModule:

    def run(
        self,
        portfolio: Any,
        payments: Optional[Any] = None,
        calls: Optional[Any] = None,
        calls_duration_col: Optional[str] = None,
        calls_success_threshold: int = 30,
        date_from: Optional[Any] = None,
        date_to: Optional[Any] = None,
        supervisors: Optional[List[str]] = None,
        **kwargs
    ) -> Dict[str, Any]:

        # ── 1. تجهيز المحفظة كـ Pandas DataFrame ──
        if isinstance(portfolio, pl.DataFrame):
            df_p = portfolio.to_pandas()
        elif isinstance(portfolio, pd.DataFrame):
            df_p = portfolio.copy()
        else:
            df_p = pd.DataFrame(portfolio)

        if len(df_p) == 0:
            raise ValueError("ملف المحفظة فارغ!")

        # ── 2. كشف الأعمدة الأساسية ──
        col_debt      = _detect_col(df_p, _DEBT_COLS)
        col_id        = _detect_col(df_p, _ID_COLS)
        col_debt_a    = _detect_col(df_p, _DEBT_AMT_COLS)
        col_rem_a     = _detect_col(df_p, _REMAIN_AMT_COLS)
        col_sup       = _detect_col(df_p, _SUPERVISOR_COLS)
        col_coll      = _detect_col(df_p, _COLLECTOR_COLS)
        col_phone     = _detect_col(df_p, _PHONE_COLS)
        col_portfolio = _detect_col(df_p, _PORTFOLIO_COLS)
        col_cancel    = _detect_col(df_p, _CANCEL_DATE_COLS)
        col_followup  = _detect_col(df_p, _FOLLOWUP_DATE_COLS)
        col_stat_main = _detect_col(df_p, _STATUS_MAIN_COLS)
        col_stat_sub  = _detect_col(df_p, _STATUS_SUB_COLS)

        # تنظيف الحقول التحليلية
        df_p["_debt_clean"] = df_p[col_debt].apply(clean_id) if col_debt else [""] * len(df_p)
        df_p["_cid_clean"]  = df_p[col_id].apply(clean_id) if col_id else [""] * len(df_p)
        df_p["_debt_amt"]   = to_clean_num(df_p[col_debt_a]) if col_debt_a else pd.Series(0.0, index=df_p.index)
        df_p["_remain_amt"] = to_clean_num(df_p[col_rem_a]) if col_rem_a else pd.Series(0.0, index=df_p.index)
        df_p["_supervisor"] = df_p[col_sup].fillna("غير محدد").astype(str).str.strip() if col_sup else pd.Series("غير محدد", index=df_p.index)
        df_p["_collector"]  = df_p[col_coll].fillna("غير محدد").astype(str).str.strip() if col_coll else pd.Series("غير محدد", index=df_p.index)

        # فلترة المشرفين إذا تم تحديدهم
        if supervisors and len(supervisors) > 0:
            clean_sups = set(str(s).strip() for s in supervisors if str(s).strip())
            if clean_sups:
                df_p = df_p[df_p["_supervisor"].isin(clean_sups)].copy()
                if df_p.empty:
                    raise ValueError("لا توجد سجلات مطابقة للمشرفين المحددين في ملف المحفظة!")

        # ── 3. كاشف التغطية (من عمود تاريخ المتابعة) ──
        if col_followup and col_followup in df_p.columns:
            f_vals = df_p[col_followup].fillna("").astype(str).str.strip()
            # فحص إذا كان التاريخ غير فارغ ولا يحوي رموز الفراغ
            df_p["_is_covered"] = (f_vals != "") & (~f_vals.str.lower().isin(["nan", "none", "null", "-", "0", "0.0"]))
        else:
            df_p["_is_covered"] = False

        # ── 4. تصنيف حالات التواصل (من الحالة الرئيسية والفرعية كطلب المستخدم) ──
        main_c = df_p[col_stat_main].fillna("").astype(str).str.strip() if col_stat_main else pd.Series("", index=df_p.index)
        sub_c  = df_p[col_stat_sub].fillna("").astype(str).str.strip() if col_stat_sub else pd.Series("", index=df_p.index)
        combined_status = main_c + " " + sub_c

        is_no_contact = (
            combined_status.str.contains(r"عدم توصل|عدم تتوصل|الرقم لا يخص|لا يوجد ارقام|غير مستعمل|غير صحيح|لا يخص|لايخص", regex=True) |
            ((main_c == "") & (sub_c == "")) |
            main_c.isin(["---", "-", "null", "none"])
        )
        is_closed_no_ans = (~is_no_contact) & (combined_status.str.contains(r"لايرد|لا يرد|مغلق", regex=True))
        is_contacted = (~is_no_contact) & (~is_closed_no_ans)

        df_p["_contact_status"] = "عدم التوصل"
        df_p.loc[is_closed_no_ans, "_contact_status"] = "لا يرد ومغلق"
        df_p.loc[is_contacted, "_contact_status"] = "تم التوصل"

        prio_map = {"تم التوصل": 1, "لا يرد ومغلق": 2, "عدم التوصل": 3}
        df_p["_prio"] = df_p["_contact_status"].map(prio_map).fillna(3)

        # ── 5. كاشف توفر الأرقام الرئيسية ──
        if col_phone and col_phone in df_p.columns:
            raw_phone = df_p[col_phone].fillna("").astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)
            has_valid_num = (raw_phone.str.len() >= 7) & (~raw_phone.str.lower().isin(["", "0", "0.0", "none", "nan", "null", "-"]))
            df_p["_phone_status"] = has_valid_num.map({True: "يوجد رقم رئيسي", False: "لا يوجد رقم مسجل"})
        else:
            df_p["_phone_status"] = "لا يوجد رقم مسجل"

        # ── 6. استخراج سنة التعثر (من تاريخ فصل الخدمة / تاريخ إلغاء الخدمة حصراً) ──
        # التحقق من وجود كولوم فصل الخدمة أو كولوم تاريخ آخر إلغاء في ملف المحفظة
        col_disc = _detect_col(df_p, ["تاريخ فصل الخدمة", "تاريخ فصل الخدمه"])
        col_canc = _detect_col(df_p, ["تاريخ الغاء الخدمة", "تاريخ إلغاء الخدمة", "تاريخ اخر الغاء", "تاريخ آخر إلغاء", "تاريخ أخر الغاء"])
        col_vint_direct = _detect_col(df_p, ["سنة التعثر", "عمر الدين"])

        if col_disc and col_disc in df_p.columns:
            s_yr = df_p[col_disc].apply(extract_year_smart)
            # دمج مع كولوم الإلغاء إذا كانت هناك قيم فارغة في كولوم الفصل
            if col_canc and col_canc in df_p.columns:
                s_canc_yr = df_p[col_canc].apply(extract_year_smart)
                s_yr = s_yr.where(s_yr != "غير محدد", s_canc_yr)
            df_p["سنة التعثر"] = s_yr
            vintage_desc = f"{col_disc}" + (f" / {col_canc}" if col_canc else "")
        elif col_canc and col_canc in df_p.columns:
            df_p["سنة التعثر"] = df_p[col_canc].apply(extract_year_smart)
            vintage_desc = col_canc
        elif col_vint_direct and col_vint_direct in df_p.columns:
            df_p["سنة التعثر"] = df_p[col_vint_direct].apply(extract_year_smart)
            vintage_desc = col_vint_direct
        elif col_cancel and col_cancel in df_p.columns:
            df_p["سنة التعثر"] = df_p[col_cancel].apply(extract_year_smart)
            vintage_desc = col_cancel
        else:
            df_p["سنة التعثر"] = "غير محدد"
            vintage_desc = "غير متوفر"

        df_p["_main_status_clean"] = main_c.replace("", "غير محدد")

        # ── 7. الربط بشيت التحصيل برقم المديونية ──
        debt_to_col: Dict[str, str] = {}
        cid_to_col: Dict[str, str] = {}
        for did, cid, cname in zip(df_p["_debt_clean"], df_p["_cid_clean"], df_p["_collector"]):
            if did and did not in debt_to_col:
                debt_to_col[did] = cname
            if cid and cid not in cid_to_col:
                cid_to_col[cid] = cname

        debt_collection_map: Dict[str, float] = {}
        collector_collection: Dict[str, float] = {c: 0.0 for c in df_p["_collector"].unique()}
        total_pay_amt = 0.0
        has_payments = False

        if payments is not None:
            if isinstance(payments, pl.DataFrame):
                df_pay = payments.to_pandas()
            elif isinstance(payments, pd.DataFrame):
                df_pay = payments.copy()
            else:
                df_pay = pd.DataFrame(payments).copy()

            if not df_pay.empty:
                pay_debt_col = _detect_col(df_pay, _DEBT_COLS)
                pay_cid_col  = _detect_col(df_pay, _ID_COLS)
                pay_amt_col  = _detect_col(df_pay, _PAY_AMT_COLS)
                pay_date_col = _detect_col(df_pay, _PAY_DATE_COLS)
                pay_col_col  = _detect_col(df_pay, _COLLECTOR_COLS)

                if pay_amt_col:
                    has_payments = True
                    df_pay["_amt"] = to_clean_num(df_pay[pay_amt_col])

                    if pay_date_col:
                        df_pay["_p_date"] = pd.to_datetime(df_pay[pay_date_col], errors="coerce")
                        if date_from:
                            df_pay = df_pay[df_pay["_p_date"] >= pd.Timestamp(date_from)]
                        if date_to:
                            df_pay = df_pay[df_pay["_p_date"] <= pd.Timestamp(date_to)]

                    p_debts = df_pay[pay_debt_col].apply(clean_id) if pay_debt_col else [""] * len(df_pay)
                    p_cids  = df_pay[pay_cid_col].apply(clean_id) if pay_cid_col else [""] * len(df_pay)
                    p_cols  = df_pay[pay_col_col].astype(str).str.strip() if pay_col_col else [""] * len(df_pay)
                    p_amts  = df_pay["_amt"].values

                    for amt, did, cid, direct_col in zip(p_amts, p_debts, p_cids, p_cols):
                        if amt <= 0.0:
                            continue
                        total_pay_amt += amt
                        if did:
                            debt_collection_map[did] = debt_collection_map.get(did, 0.0) + amt
                        assigned_col = debt_to_col.get(did) or cid_to_col.get(cid) or (direct_col if direct_col in collector_collection else None)
                        if assigned_col:
                            collector_collection[assigned_col] = collector_collection.get(assigned_col, 0.0) + amt

        df_p["_collected_amt"] = df_p["_debt_clean"].map(debt_collection_map).fillna(0.0)

        # ── 8. تجميع ملخص حالات التواصل ──
        total_records = len(df_p)
        cnt_contacted = int((df_p["_contact_status"] == "تم التوصل").sum())
        cnt_closed    = int((df_p["_contact_status"] == "لا يرد ومغلق").sum())
        cnt_not       = int((df_p["_contact_status"] == "عدم التوصل").sum())
        cnt_total_not = cnt_closed + cnt_not

        pct_contacted = round(cnt_contacted / total_records * 100, 1) if total_records > 0 else 0.0
        pct_closed    = round(cnt_closed / total_records * 100, 1) if total_records > 0 else 0.0
        pct_not       = round(cnt_not / total_records * 100, 1) if total_records > 0 else 0.0
        pct_total_not = round(cnt_total_not / total_records * 100, 1) if total_records > 0 else 0.0

        cust_agg = df_p.groupby("_cid_clean")["_prio"].min().reset_index()
        u_total     = len(cust_agg)
        u_contacted = int((cust_agg["_prio"] == 1).sum())
        u_closed    = int((cust_agg["_prio"] == 2).sum())
        u_not       = int((cust_agg["_prio"] == 3).sum())
        u_total_not = u_closed + u_not

        u_pct_contacted = round(u_contacted / u_total * 100, 1) if u_total > 0 else 0.0
        u_pct_closed    = round(u_closed / u_total * 100, 1) if u_total > 0 else 0.0
        u_pct_not       = round(u_not / u_total * 100, 1) if u_total > 0 else 0.0
        u_pct_total_not = round(u_total_not / u_total * 100, 1) if u_total > 0 else 0.0

        contact_summary_df = pd.DataFrame([
            {
                "حالة التواصل": "تم التوصل",
                "عدد الحسابات (السجلات)": cnt_contacted,
                "نسبة الحسابات %": f"{pct_contacted}%",
                "عدد العملاء (الهويات)": u_contacted,
                "نسبة العملاء %": f"{u_pct_contacted}%"
            },
            {
                "حالة التواصل": "لا يرد ومغلق",
                "عدد الحسابات (السجلات)": cnt_closed,
                "نسبة الحسابات %": f"{pct_closed}%",
                "عدد العملاء (الهويات)": u_closed,
                "نسبة العملاء %": f"{u_pct_closed}%"
            },
            {
                "حالة التواصل": "عدم التوصل",
                "عدد الحسابات (السجلات)": cnt_not,
                "نسبة الحسابات %": f"{pct_not}%",
                "عدد العملاء (الهويات)": u_not,
                "نسبة العملاء %": f"{u_pct_not}%"
            },
            {
                "حالة التواصل": "إجمالي عدم التوصل (مغلق + عدم توصل)",
                "عدد الحسابات (السجلات)": cnt_total_not,
                "نسبة الحسابات %": f"{pct_total_not}%",
                "عدد العملاء (الهويات)": u_total_not,
                "نسبة العملاء %": f"{u_pct_total_not}%"
            },
            {
                "حالة التواصل": "الإجمالي العام",
                "عدد الحسابات (السجلات)": total_records,
                "نسبة الحسابات %": "100.0%",
                "عدد العملاء (الهويات)": u_total,
                "نسبة العملاء %": "100.0%"
            }
        ])

        # ── 9. تجميع كاشف توفر الأرقام الرئيسية ──
        phone_rows = []
        for p_label in ["يوجد رقم رئيسي", "لا يوجد رقم مسجل"]:
            sub_p = df_p[df_p["_phone_status"] == p_label]
            p_cnt = len(sub_p)
            p_cust = sub_p["_cid_clean"].nunique()
            p_debt = sub_p["_debt_amt"].sum()
            p_rem  = sub_p["_remain_amt"].sum()
            p_coll = sub_p["_collected_amt"].sum()
            p_rate = round(p_coll / p_debt * 100, 2) if p_debt > 0 else 0.0

            phone_rows.append({
                "حالة الرقم": p_label,
                "عدد الحسابات": p_cnt,
                "نسبة الحسابات %": f"{round(p_cnt/total_records*100, 1)}%",
                "عدد العملاء": p_cust,
                "نسبة العملاء %": f"{round(p_cust/u_total*100, 1)}%",
                "مبلغ المديونية": round(p_debt, 2),
                "متبقي السداد": round(p_rem, 2),
                "إجمالي التحصيل": round(p_coll, 2),
                "نسبة التحصيل %": p_rate
            })

        phone_summary_df = pd.DataFrame(phone_rows)

        # ── 10. تجميع جدول سنة التعثر (عمر الدين) ──
        tot_debt_all   = df_p["_debt_amt"].sum()
        tot_remain_all = df_p["_remain_amt"].sum()
        tot_coll_rate  = round(total_pay_amt / tot_debt_all * 100, 2) if tot_debt_all > 0 else 0.0

        v_rows = []
        for v_yr, grp in df_p.groupby("سنة التعثر"):
            n_recs = len(grp)
            n_cust = grp["_cid_clean"].nunique()
            d_amt  = grp["_debt_amt"].sum()
            r_amt  = grp["_remain_amt"].sum()
            c_amt  = grp["_collected_amt"].sum()
            c_pct  = round(c_amt / d_amt * 100, 2) if d_amt > 0 else 0.0

            yr_contacted = (grp["_contact_status"] == "تم التوصل").sum()
            yr_contact_pct = round(yr_contacted / n_recs * 100, 1) if n_recs > 0 else 0.0

            v_rows.append({
                "سنة التعثر": str(v_yr),
                "عمر الدين (السنة)": str(v_yr),
                "عدد الحسابات": n_recs,
                "عدد العملاء": n_cust,
                "مبلغ المديونية": round(d_amt, 2),
                "متبقي سداد موثق": round(r_amt, 2),
                "إجمالي التحصيل": round(c_amt, 2),
                "نسبة التحصيل %": c_pct,
                "تم التوصل": yr_contacted,
                "نسبة التوصل %": f"{yr_contact_pct}%"
            })

        vintage_df = pd.DataFrame(v_rows)
        def _yr_sort(val):
            try: return int(val)
            except: return -1

        if not vintage_df.empty:
            vintage_df["sort_k"] = vintage_df["سنة التعثر"].apply(_yr_sort)
            vintage_df = vintage_df.sort_values("sort_k", ascending=False).drop(columns=["sort_k"]).reset_index(drop=True)

            v_tot = {
                "سنة التعثر": "الإجمالي",
                "عمر الدين (السنة)": "الإجمالي",
                "عدد الحسابات": total_records,
                "عدد العملاء": u_total,
                "مبلغ المديونية": round(tot_debt_all, 2),
                "متبقي سداد موثق": round(tot_remain_all, 2),
                "إجمالي التحصيل": round(total_pay_amt, 2),
                "نسبة التحصيل %": tot_coll_rate,
                "تم التوصل": cnt_contacted,
                "نسبة التوصل %": f"{pct_contacted}%"
            }
            vintage_table_display = pd.concat([vintage_df, pd.DataFrame([v_tot])], ignore_index=True)
        else:
            vintage_table_display = pd.DataFrame()

        # ── 11. تجميع التحصيل حسب الحالة الرئيسية ──
        st_rows = []
        for st_name, grp in df_p.groupby("_main_status_clean"):
            st_recs = len(grp)
            st_cust = grp["_cid_clean"].nunique()
            st_debt = grp["_debt_amt"].sum()
            st_rem  = grp["_remain_amt"].sum()
            st_coll = grp["_collected_amt"].sum()
            st_rate = round(st_coll / st_debt * 100, 2) if st_debt > 0 else 0.0
            st_share = round(st_coll / total_pay_amt * 100, 2) if total_pay_amt > 0 else 0.0

            st_rows.append({
                "الحالة الرئيسية": st_name,
                "عدد الحسابات": st_recs,
                "عدد العملاء": st_cust,
                "مبلغ المديونية": round(st_debt, 2),
                "متبقي السداد": round(st_rem, 2),
                "إجمالي التحصيل": round(st_coll, 2),
                "نسبة التحصيل %": st_rate,
                "نسبة المساهمة في التحصيل %": st_share
            })

        status_coll_df = pd.DataFrame(st_rows)
        if not status_coll_df.empty:
            status_coll_df = status_coll_df.sort_values("إجمالي التحصيل", ascending=False).reset_index(drop=True)
            top_status_row = status_coll_df.iloc[0]
            top_status_desc = f"{top_status_row['الحالة الرئيسية']} ({top_status_row['إجمالي التحصيل']:,.2f} ﷼ - {top_status_row['نسبة المساهمة في التحصيل %']}%)"
        else:
            top_status_desc = "غير محدد"


        # ── 12. تجميع وتحليل المحافظ (Portfolio Breakdown) ──
        port_rows = []
        if col_portfolio and col_portfolio in df_p.columns:
            for p_name, grp in df_p.groupby(col_portfolio, dropna=False):
                p_label = str(p_name).strip() if pd.notna(p_name) and str(p_name).strip() not in ("", "nan", "None") else "غير محدد"
                p_recs  = len(grp)
                p_cust  = grp["_cid_clean"].nunique()
                p_debt  = grp["_debt_amt"].sum()
                p_rem   = grp["_remain_amt"].sum()
                p_coll  = grp["_collected_amt"].sum()
                p_rate  = round(p_coll / p_debt * 100, 2) if p_debt > 0 else 0.0
                
                p_cont  = (grp["_contact_status"] == "تم التوصل").sum()
                p_cont_rate = round(p_cont / p_recs * 100, 1) if p_recs > 0 else 0.0
                p_no_ans = (grp["_contact_status"] == "لا يرد ومغلق").sum()
                p_no_cont = (grp["_contact_status"] == "عدم التوصل").sum()
                
                p_no_ans_rate = round(p_no_ans / p_recs * 100, 1) if p_recs > 0 else 0.0
                p_no_cont_rate = round(p_no_cont / p_recs * 100, 1) if p_recs > 0 else 0.0

                port_rows.append({
                    "المحفظة": p_label,
                    "عدد الحسابات": p_recs,
                    "عدد العملاء": p_cust,
                    "مبلغ المديونية": round(p_debt, 2),
                    "متبقي سداد موثق": round(p_rem, 2),
                    "إجمالي التحصيل": round(p_coll, 2),
                    "نسبة التحصيل %": p_rate,
                    "تم التوصل": p_cont,
                    "نسبة التوصل %": f"{p_cont_rate}%",
                    "لا يرد ومغلق": p_no_ans,
                    "نسبة لا يرد ومغلق %": f"{p_no_ans_rate}%",
                    "عدم التوصل": p_no_cont,
                    "نسبة عدم التوصل %": f"{p_no_cont_rate}%"
                })
        
        if port_rows:
            port_df = pd.DataFrame(port_rows).sort_values("مبلغ المديونية", ascending=False).reset_index(drop=True)
            p_tot = {
                "المحفظة": "الإجمالي",
                "عدد الحسابات": total_records,
                "عدد العملاء": u_total,
                "مبلغ المديونية": round(tot_debt_all, 2),
                "متبقي سداد موثق": round(tot_remain_all, 2),
                "إجمالي التحصيل": round(total_pay_amt, 2),
                "نسبة التحصيل %": tot_coll_rate,
                "تم التوصل": cnt_contacted,
                "نسبة التوصل %": f"{pct_contacted}%",
                "لا يرد ومغلق": cnt_closed,
                "نسبة لا يرد ومغلق %": f"{pct_closed}%",
                "عدم التوصل": cnt_not,
                "نسبة عدم التوصل %": f"{pct_not}%"
            }
            portfolio_summary_df = pd.concat([port_df, pd.DataFrame([p_tot])], ignore_index=True)
            top_port_debt = f"{port_df.iloc[0]['المحفظة']} ({port_df.iloc[0]['مبلغ المديونية']:,.2f} ﷼)"
            port_coll_sorted = port_df.sort_values("إجمالي التحصيل", ascending=False)
            top_port_coll = f"{port_coll_sorted.iloc[0]['المحفظة']} ({port_coll_sorted.iloc[0]['إجمالي التحصيل']:,.2f} ﷼)"
        else:
            portfolio_summary_df = pd.DataFrame()
            top_port_debt = "غير متوفر"
            top_port_coll = "غير متوفر"

        # ── 12. تجميع أداء ومقارنة المشرفين (التغطية والتحصيل وحالات التواصل) ──
        sup_rows = []
        for s_name, grp in df_p.groupby("_supervisor"):
            s_recs = len(grp)
            s_cust = grp["_cid_clean"].nunique()
            s_debt = grp["_debt_amt"].sum()
            s_rem  = grp["_remain_amt"].sum()

            s_cols = grp["_collector"].unique()
            s_coll = sum(collector_collection.get(c, 0.0) for c in s_cols)
            s_pct  = round(s_coll / s_debt * 100, 2) if s_debt > 0 else 0.0

            s_cov_cust = grp[grp["_is_covered"]]["_cid_clean"].nunique()
            s_cov_pct  = round(s_cov_cust / s_cust * 100, 1) if s_cust > 0 else 0.0

            s_contacted = int((grp["_contact_status"] == "تم التوصل").sum())
            s_contact_pct = round(s_contacted / s_recs * 100, 1) if s_recs > 0 else 0.0

            s_closed = int((grp["_contact_status"] == "لا يرد ومغلق").sum())
            s_closed_pct = round(s_closed / s_recs * 100, 1) if s_recs > 0 else 0.0

            s_not = int((grp["_contact_status"] == "عدم التوصل").sum())
            s_not_pct = round(s_not / s_recs * 100, 1) if s_recs > 0 else 0.0

            sup_rows.append({
                "المشرف": s_name,
                "عدد الحسابات": s_recs,
                "عدد العملاء": s_cust,
                "العملاء المغطين": s_cov_cust,
                "نسبة التغطية %": s_cov_pct,
                "مبلغ المديونية": round(s_debt, 2),
                "متبقي سداد موثق": round(s_rem, 2),
                "إجمالي التحصيل": round(s_coll, 2),
                "نسبة التحصيل %": s_pct,
                "تم التوصل": s_contacted,
                "نسبة التوصل %": f"{s_contact_pct}%",
                "لا يرد ومغلق": s_closed,
                "نسبة لا يرد ومغلق %": f"{s_closed_pct}%",
                "عدم التوصل": s_not,
                "نسبة عدم التوصل %": f"{s_not_pct}%"
            })

        sup_df = pd.DataFrame(sup_rows)
        if not sup_df.empty:
            sup_df = sup_df.sort_values("إجمالي التحصيل", ascending=False).reset_index(drop=True)

        # ── 13. تجميع أداء وكيرف المحصلين (Coverage vs Collection Quadrants) ──
        col_rows = []
        avg_coll_val = total_pay_amt / max(df_p["_collector"].nunique(), 1)

        for c_name, grp in df_p.groupby("_collector"):
            c_recs = len(grp)
            c_cust = grp["_cid_clean"].nunique()
            c_debt = grp["_debt_amt"].sum()
            c_rem  = grp["_remain_amt"].sum()
            c_coll = collector_collection.get(c_name, 0.0)
            c_pct  = round(c_coll / c_debt * 100, 2) if c_debt > 0 else 0.0

            c_cov_cust = grp[grp["_is_covered"]]["_cid_clean"].nunique()
            c_cov_pct  = round(c_cov_cust / c_cust * 100, 1) if c_cust > 0 else 0.0

            c_sup = grp["_supervisor"].iloc[0] if not grp["_supervisor"].empty else "غير محدد"
            c_contacted = int((grp["_contact_status"] == "تم التوصل").sum())
            c_contact_pct = round(c_contacted / c_recs * 100, 1) if c_recs > 0 else 0.0

            # مصفوفة الأداء (Quadrant)
            is_high_coll = (c_coll >= avg_coll_val)
            is_high_cov  = (c_cov_pct >= 50.0)

            if is_high_coll and is_high_cov:
                quadrant = "🌟 نجوم الأداء (High Cov + High Coll)"
            elif not is_high_coll and is_high_cov:
                quadrant = "🎯 مجهود عالي / فرص كامنة (High Cov + Low Coll)"
            elif is_high_coll and not is_high_cov:
                quadrant = "⚡ تحصيل نوعي / استجابة سريعة (Low Cov + High Coll)"
            else:
                quadrant = "⚠️ يحتاج متابعة وتدخل (Low Cov + Low Coll)"

            c_closed = int((grp["_contact_status"] == "لا يرد ومغلق").sum())
            c_closed_pct = round(c_closed / c_recs * 100, 1) if c_recs > 0 else 0.0

            c_not = int((grp["_contact_status"] == "عدم التوصل").sum())
            c_not_pct = round(c_not / c_recs * 100, 1) if c_recs > 0 else 0.0

            col_rows.append({
                "المحصل": c_name,
                "المشرف": c_sup,
                "عدد الحسابات": c_recs,
                "عدد العملاء": c_cust,
                "العملاء المغطين": c_cov_cust,
                "نسبة التغطية %": c_cov_pct,
                "مبلغ المديونية": round(c_debt, 2),
                "متبقي سداد موثق": round(c_rem, 2),
                "إجمالي التحصيل": round(c_coll, 2),
                "نسبة التحصيل %": c_pct,
                "تم التوصل": c_contacted,
                "نسبة التوصل %": f"{c_contact_pct}%",
                "لا يرد ومغلق": c_closed,
                "نسبة لا يرد ومغلق %": f"{c_closed_pct}%",
                "عدم التوصل": c_not,
                "نسبة عدم التوصل %": f"{c_not_pct}%",
                "تصنيف الربع": quadrant
            })

        col_df = pd.DataFrame(col_rows)
        if not col_df.empty:
            col_df_sorted = col_df.sort_values("إجمالي التحصيل", ascending=False).reset_index(drop=True)
            top_10    = col_df_sorted.head(10).copy()
            bottom_10 = col_df_sorted.tail(10).sort_values("إجمالي التحصيل", ascending=True).reset_index(drop=True)
        else:
            col_df_sorted = pd.DataFrame()
            top_10        = pd.DataFrame()
            bottom_10     = pd.DataFrame()

        # حساب التغطية الكلية على المحفظة
        total_covered_cust = df_p[df_p["_is_covered"]]["_cid_clean"].nunique()
        total_coverage_pct = round(total_covered_cust / u_total * 100, 1) if u_total > 0 else 0.0

        phone_valid_cnt = (df_p["_phone_status"] == "يوجد رقم رئيسي").sum()
        phone_valid_pct = round(phone_valid_cnt / total_records * 100, 1) if total_records > 0 else 0.0

        export_portfolio = df_p.drop(
            columns=["_debt_clean", "_cid_clean", "_debt_amt", "_remain_amt",
                     "_supervisor", "_collector", "_prio", "_contact_status", "_collected_amt",
                     "_is_covered", "_main_status_clean", "_phone_status"],
            errors="ignore"
        )

        stats = {
            "إجمالي الحسابات": f"{total_records:,} حساب",
            "إجمالي العملاء": f"{u_total:,} عميل",
            "إجمالي مبلغ المديونية": f"{tot_debt_all:,.2f} ﷼",
            "إجمالي متبقي سداد موثق": f"{tot_remain_all:,.2f} ﷼",
            "إجمالي التحصيل": f"{total_pay_amt:,.2f} ﷼",
            "نسبة التحصيل العامة %": f"{tot_coll_rate}%",
            "نسبة التغطية الكلية %": f"{total_coverage_pct}%",
            "العملاء المغطين": f"{total_covered_cust:,} عميل",
            "نسبة توفر الأرقام الرئيسية %": f"{phone_valid_pct}%",
            "تم التوصل": f"{cnt_contacted:,} ({pct_contacted}%)",
            "لا يرد ومغلق": f"{cnt_closed:,} ({pct_closed}%)",
            "عدم التوصل": f"{cnt_not:,} ({pct_not}%)",
            "نسبة التوصل العامة %": f"{pct_contacted}%",
            "أعلى حالة سداداً": top_status_desc,
            "عدد المحصلين": int(df_p["_collector"].nunique()),
            "عدد المشرفين": int(df_p["_supervisor"].nunique()),
            "عدد سنوات عمر الدين": len(vintage_df),
            "عمود عمر الدين المستخدم": vintage_desc,
            "عدد المحافظ": len(port_df) if 'port_df' in locals() and not port_df.empty else 0,
            "أعلى محفظة مديونية": top_port_debt if 'top_port_debt' in locals() else "غير محدد",
            "أعلى محفظة تحصيلاً": top_port_coll if 'top_port_coll' in locals() else "غير محدد"
        }


        # ── 14. استخراج وتحليل عملاء الخصومات (Discounted Clients) ──
        _DISCOUNT_COLS = ["نسبة الخصم", "نسبه الخصم", "الخصم", "مبلغ الخصم", "تاريخ الموافقه علي الخصم", "تاريخ الموافقة على الخصم", "discount", "Discount"]
        col_discount = _detect_col(df_p, _DISCOUNT_COLS)
        col_after_disc = _detect_col(df_p, ["المبلغ بعد الخصم", "المبلغ بعد التخفيض", "amount_after_discount"])
        col_cust_name_opt = _detect_col(df_p, ["اسم العميل", "العميل", "customer_name", "Client Name"])

        disc_rows = []
        if col_discount and col_discount in df_p.columns:
            raw_d = df_p[col_discount].fillna("").astype(str).str.strip().str.lower()
            valid_d_mask = (raw_d != "") & (~raw_d.isin(["nan", "none", "null", "0", "0.0", "0%", "-"]))
            df_disc = df_p[valid_d_mask].copy()

            if not df_disc.empty:
                for _, r_d in df_disc.iterrows():
                    c_name_val = str(r_d[col_cust_name_opt]) if col_cust_name_opt and pd.notna(r_d.get(col_cust_name_opt)) else str(r_d["_cid_clean"])
                    d_row = {
                        "رقم المديونية": r_d["_debt_clean"],
                        "اسم العميل": c_name_val,
                        "رقم الهوية": r_d["_cid_clean"],
                        "نسبة الخصم": str(r_d[col_discount]),
                        "مبلغ المديونية": round(r_d["_debt_amt"], 2),
                    }
                    if col_after_disc and col_after_disc in df_disc.columns:
                        d_row["المبلغ بعد الخصم"] = round(to_clean_num_scalar(r_d[col_after_disc]), 2)
                    d_row["مبلغ السداد المحقق"] = round(r_d["_collected_amt"], 2)
                    d_row["متبقي سداد موثق"] = round(r_d["_remain_amt"], 2)
                    d_row["الحالة الرئيسية"] = r_d["_main_status_clean"]
                    d_row["المشرف"] = r_d["_supervisor"]
                    d_row["المحصل"] = r_d["_collector"]
                    disc_rows.append(d_row)

        if disc_rows:
            discounts_df = pd.DataFrame(disc_rows).sort_values("مبلغ المديونية", ascending=False).reset_index(drop=True)
            d_tot = {
                "رقم المديونية": "الإجمالي",
                "اسم العميل": f"{len(discounts_df)} حساب",
                "رقم الهوية": f"{discounts_df['رقم الهوية'].nunique()} عميل",
                "نسبة الخصم": "-",
                "مبلغ المديونية": round(discounts_df["مبلغ المديونية"].sum(), 2)
            }
            if "المبلغ بعد الخصم" in discounts_df.columns:
                d_tot["المبلغ بعد الخصم"] = round(discounts_df["المبلغ بعد الخصم"].sum(), 2)
            d_tot["مبلغ السداد المحقق"] = round(discounts_df["مبلغ السداد المحقق"].sum(), 2)
            d_tot["متبقي سداد موثق"] = round(discounts_df["متبقي سداد موثق"].sum(), 2)
            d_tot["الحالة الرئيسية"] = "-"
            d_tot["المشرف"] = "-"
            d_tot["المحصل"] = "-"
            discounts_table_display = pd.concat([discounts_df, pd.DataFrame([d_tot])], ignore_index=True)
        else:
            discounts_table_display = pd.DataFrame()

        # ── 15. تحليل ومقارنة فرق العمل (Team Performance by Supervisor) ──
        teams_data = {}
        best_team_name = "غير محدد"
        max_team_coll = -1.0

        if not col_df.empty:
            for sup_name, grp_sup in col_df.groupby("المشرف"):
                if sup_name in ("غير محدد", "---", "-"):
                    continue
                t_cols_cnt = len(grp_sup)
                t_recs_cnt = grp_sup["عدد الحسابات"].sum()
                t_debt_sum = grp_sup["مبلغ المديونية"].sum()
                t_coll_sum = grp_sup["إجمالي التحصيل"].sum()
                t_coll_rate = round(t_coll_sum / t_debt_sum * 100, 2) if t_debt_sum > 0 else 0.0

                # متوسط نسب التوصل والتغطية للفريق
                grp_p_sup = df_p[df_p["_supervisor"] == sup_name]
                if not grp_p_sup.empty:
                    t_cov_rate = round(grp_p_sup["_is_covered"].sum() / len(grp_p_sup) * 100, 1)
                    t_cont_rate = round((grp_p_sup["_contact_status"] == "تم التوصل").sum() / len(grp_p_sup) * 100, 1)
                else:
                    t_cov_rate = 0.0
                    t_cont_rate = 0.0

                # أفضل 5 وأقل 5 محصلين في الفريق
                sorted_team = grp_sup.sort_values("إجمالي التحصيل", ascending=False)
                t_top5 = sorted_team.head(5).copy()
                t_bot5 = sorted_team.tail(5).sort_values("إجمالي التحصيل", ascending=True).copy()

                if t_coll_sum > max_team_coll:
                    max_team_coll = t_coll_sum
                    best_team_name = sup_name

                teams_data[sup_name] = {
                    "supervisor": sup_name,
                    "collectors_count": t_cols_cnt,
                    "accounts_count": t_recs_cnt,
                    "debt_sum": round(t_debt_sum, 2),
                    "coll_sum": round(t_coll_sum, 2),
                    "coll_rate": t_coll_rate,
                    "cont_rate": t_cont_rate,
                    "cov_rate": t_cov_rate,
                    "top_5": t_top5,
                    "bottom_5": t_bot5,
                    "all_members": sorted_team
                }


        # ── 16. تحليل المكالمات الفعلية (اختياري) ──
        calls_analysis = {}
        has_calls = False

        if calls is not None:
            # تحويل إلى pandas
            if isinstance(calls, pl.DataFrame):
                df_calls = calls.to_pandas()
            elif isinstance(calls, pd.DataFrame):
                df_calls = calls.copy()
            else:
                df_calls = pd.DataFrame(calls)

            if len(df_calls) > 0:
                has_calls = True
                _CALL_DEBT_COLS  = ["رقم المديونية", "رقم المديونيه", "رقم العقد", "رقم الحساب",
                                    "debt_id", "Debt ID", "AccountNo", "Account No"]
                _CALL_DUR_COLS   = ["مدة المكالمة", "المدة", "duration", "Duration", "call_duration",
                                    "مدة", "وقت المكالمة", "الوقت", "seconds", "Seconds", "ثانية"]
                _CALL_STATUS_COLS= ["حالة المكالمة", "الحالة", "status", "Status", "call_status",
                                    "نجاح", "نتيجة المكالمة", "result"]
                _CALL_AGENT_COLS = ["المحصل", "اسم المحصل", "الموظف", "المندوب", "agent", "Agent",
                                    "collector", "Collector", "user", "User"]
                _CALL_SUP_COLS   = ["المشرف", "اسم المشرف", "supervisor", "Supervisor"]

                # استخدام العمود الذي اختاره المستخدم أولاً، ثم Auto-detect
                call_debt_col  = _detect_col(df_calls, _CALL_DEBT_COLS)
                call_dur_col   = calls_duration_col if (calls_duration_col and calls_duration_col in df_calls.columns)                                  else _detect_col(df_calls, _CALL_DUR_COLS)
                call_status_col= _detect_col(df_calls, _CALL_STATUS_COLS)
                call_agent_col = _detect_col(df_calls, _CALL_AGENT_COLS)
                call_sup_col   = _detect_col(df_calls, _CALL_SUP_COLS)

                total_calls = len(df_calls)
                threshold   = int(calls_success_threshold) if calls_success_threshold else 30

                # ── دالة تحليل مدة المكالمة ──
                def parse_duration(val):
                    val = str(val).strip()
                    # HH:MM:SS
                    m = re.match(r"^(\d+):(\d+):(\d+)$", val)
                    if m:
                        return int(m.group(1))*3600 + int(m.group(2))*60 + int(m.group(3))
                    # MM:SS
                    m = re.match(r"^(\d+):(\d+)$", val)
                    if m:
                        return int(m.group(1))*60 + int(m.group(2))
                    # >30 or <30
                    m = re.match(r"^[><]?=?\s*(\d+)", val)
                    if m:
                        return int(m.group(1))
                    try:
                        return float(re.sub(r"[^\d.]", "", val))
                    except Exception:
                        return 0.0

                # ── تحديد المكالمات الناجحة ──
                successful_mask = pd.Series([False] * total_calls, index=df_calls.index)
                durations_series = pd.Series([0.0] * total_calls, index=df_calls.index)

                if call_dur_col:
                    raw_dur_str = df_calls[call_dur_col].astype(str).str.strip()
                    is_gt_notation = raw_dur_str.str.startswith(">")
                    is_lt_notation = raw_dur_str.str.startswith("<")
                    durations_parsed  = df_calls[call_dur_col].apply(parse_duration)
                    durations_series  = durations_parsed.astype(float)
                    successful_mask   = ((durations_series > threshold) | is_gt_notation) & ~is_lt_notation
                elif call_status_col:
                    success_keywords  = ["ناجح", "ناجحة", "نجاح", "completed", "success",
                                         "answered", "connected", "مكتملة", "تم التوصل", "1", "yes", "نعم"]
                    status_lower = df_calls[call_status_col].astype(str).str.lower().str.strip()
                    for kw in success_keywords:
                        successful_mask = successful_mask | status_lower.str.contains(kw, na=False)
                else:
                    successful_mask = pd.Series([True] * total_calls, index=df_calls.index)

                df_success = df_calls[successful_mask].copy()
                total_successful  = int(len(df_success))
                total_failed      = total_calls - total_successful
                success_rate      = round(total_successful / total_calls * 100, 2) if total_calls > 0 else 0.0

                # متوسطات المدة (كلية)
                avg_dur_all     = round(float(durations_series.mean()), 1) if call_dur_col else None
                avg_dur_success = None
                if call_dur_col and total_successful > 0:
                    avg_dur_success = round(float(durations_series[successful_mask].mean()), 1)

                # ── ربط المكالمات بالمحفظة ──
                call_debt_counts: Dict[str, int] = {}
                if call_debt_col:
                    for debt_val in df_success[call_debt_col].astype(str).str.strip():
                        d = re.sub(r"\.0$", "", debt_val)
                        if d and d not in ("nan", "none", "null", ""):
                            call_debt_counts[d] = call_debt_counts.get(d, 0) + 1

                # ── تجميع حسب المحصل والمشرف ──
                collector_call_stats: Dict[str, dict] = {}

                if call_agent_col:
                    for agent_val, grp in df_calls.groupby(df_calls[call_agent_col].astype(str).str.strip()):
                        if not agent_val or agent_val in ("nan", "none", ""):
                            continue
                        total_ag   = len(grp)
                        success_ag = int(successful_mask.loc[grp.index].sum())
                        # متوسط مدة المكالمات الناجحة لهذا المحصل
                        if call_dur_col and success_ag > 0:
                            succ_idx   = grp.index[successful_mask.loc[grp.index]]
                            avg_ag_dur = round(float(durations_series.loc[succ_idx].mean()), 1)
                            max_ag_dur = round(float(durations_series.loc[succ_idx].max()), 1)
                        else:
                            avg_ag_dur = None
                            max_ag_dur = None
                        sup_val = ""
                        if call_sup_col:
                            sup_vals = grp[call_sup_col].astype(str).str.strip().dropna()
                            sup_val  = sup_vals.mode()[0] if len(sup_vals) > 0 else ""
                        collector_call_stats[agent_val] = {
                            "المحصل":                agent_val,
                            "المشرف":                sup_val,
                            "إجمالي المكالمات":      total_ag,
                            "المكالمات الناجحة":     success_ag,
                            "المكالمات غير الناجحة": total_ag - success_ag,
                            "نسبة النجاح %":         round(success_ag / total_ag * 100, 2) if total_ag > 0 else 0.0,
                            "متوسط مدة الناجحة (ث)": avg_ag_dur,
                            "أطول مكالمة ناجحة (ث)": max_ag_dur,
                        }

                # لو مفيش عمود محصل، نربط برقم المديونية مع حساب متوسطات المدة
                if not collector_call_stats and call_debt_col:
                    # خريطة سريعة للمشرف حسب رقم المديونية
                    debt_to_sup = {}
                    if "_supervisor" in df_p.columns:
                        for did_p, sup_p in zip(df_p["_debt_clean"], df_p["_supervisor"]):
                            if did_p and did_p not in debt_to_sup:
                                debt_to_sup[did_p] = str(sup_p)

                    col_succ_durs: Dict[str, List[float]] = {}
                    raw_debts_list = df_calls[call_debt_col].astype(str).str.strip().tolist()
                    succ_list = successful_mask.tolist()
                    durs_list = durations_series.tolist()

                    for raw_d, is_succ, dur_v in zip(raw_debts_list, succ_list, durs_list):
                        did = re.sub(r"\.0$", "", raw_d)
                        if not did or did in ("nan", "none", "null", ""):
                            continue
                        col_name = debt_to_col.get(did, "")
                        if not col_name:
                            continue
                        sup_name = debt_to_sup.get(did, "")
                        if col_name not in collector_call_stats:
                            collector_call_stats[col_name] = {
                                "المحصل": col_name, "المشرف": sup_name,
                                "إجمالي المكالمات": 0, "المكالمات الناجحة": 0,
                                "المكالمات غير الناجحة": 0, "نسبة النجاح %": 0.0,
                                "متوسط مدة الناجحة (ث)": None, "أطول مكالمة ناجحة (ث)": None,
                            }
                            col_succ_durs[col_name] = []
                        collector_call_stats[col_name]["إجمالي المكالمات"] += 1
                        if is_succ:
                            collector_call_stats[col_name]["المكالمات الناجحة"] += 1
                            if call_dur_col:
                                col_succ_durs[col_name].append(float(dur_v))

                    for col_name in collector_call_stats:
                        t = collector_call_stats[col_name]["إجمالي المكالمات"]
                        s = collector_call_stats[col_name]["المكالمات الناجحة"]
                        collector_call_stats[col_name]["المكالمات غير الناجحة"] = t - s
                        collector_call_stats[col_name]["نسبة النجاح %"] = round(s/t*100, 2) if t > 0 else 0.0
                        d_list = col_succ_durs.get(col_name, [])
                        if call_dur_col and d_list:
                            collector_call_stats[col_name]["متوسط مدة الناجحة (ث)"] = round(sum(d_list) / len(d_list), 1)
                            collector_call_stats[col_name]["أطول مكالمة ناجحة (ث)"] = round(max(d_list), 1)

                # ── تجميع حسب المشرف ──
                supervisor_call_stats: Dict[str, dict] = {}
                for col_name, row_data in collector_call_stats.items():
                    sup = row_data.get("المشرف", "غير محدد") or "غير محدد"
                    if sup not in supervisor_call_stats:
                        supervisor_call_stats[sup] = {
                            "المشرف": sup, "إجمالي المكالمات": 0,
                            "المكالمات الناجحة": 0, "المكالمات غير الناجحة": 0,
                            "نسبة النجاح %": 0.0
                        }
                    supervisor_call_stats[sup]["إجمالي المكالمات"]      += row_data["إجمالي المكالمات"]
                    supervisor_call_stats[sup]["المكالمات الناجحة"]     += row_data["المكالمات الناجحة"]
                    supervisor_call_stats[sup]["المكالمات غير الناجحة"] += row_data["المكالمات غير الناجحة"]
                for sup in supervisor_call_stats:
                    t = supervisor_call_stats[sup]["إجمالي المكالمات"]
                    s = supervisor_call_stats[sup]["المكالمات الناجحة"]
                    supervisor_call_stats[sup]["نسبة النجاح %"] = round(s/t*100, 2) if t > 0 else 0.0

                # ── بناء DataFrames ──
                collector_calls_df = pd.DataFrame(list(collector_call_stats.values()))
                top_caller = bottom_caller = None
                if not collector_calls_df.empty:
                    collector_calls_df = collector_calls_df.sort_values(
                        "المكالمات الناجحة", ascending=False
                    ).reset_index(drop=True)
                    if call_dur_col and "متوسط مدة الناجحة (ث)" in collector_calls_df.columns:
                        dur_sorted = collector_calls_df[collector_calls_df["متوسط مدة الناجحة (ث)"].notna()].copy()
                        if not dur_sorted.empty:
                            top_caller    = dur_sorted.nlargest(1, "متوسط مدة الناجحة (ث)").iloc[0]["المحصل"]
                            bottom_caller = dur_sorted.nsmallest(1, "متوسط مدة الناجحة (ث)").iloc[0]["المحصل"]

                supervisor_calls_df = pd.DataFrame(list(supervisor_call_stats.values()))
                if not supervisor_calls_df.empty:
                    supervisor_calls_df = supervisor_calls_df.sort_values(
                        "المكالمات الناجحة", ascending=False
                    ).reset_index(drop=True)

                calls_analysis = {
                    "total_calls":          total_calls,
                    "total_successful":     total_successful,
                    "total_failed":         total_failed,
                    "success_rate":         success_rate,
                    "threshold_seconds":    threshold,
                    "avg_dur_all":          avg_dur_all,
                    "avg_dur_success":      avg_dur_success,
                    "top_caller_by_dur":    top_caller,
                    "bottom_caller_by_dur": bottom_caller,
                    "collector_calls_df":   collector_calls_df,
                    "supervisor_calls_df":  supervisor_calls_df,
                    "has_duration_col":     call_dur_col is not None,
                    "has_status_col":       call_status_col is not None,
                    "linked_to_portfolio":  call_debt_col is not None,
                }
                _log.info("📞 تحليل المكالمات: إجمالي=%d، ناجحة=%d، نسبة=%.1f%%",
                          total_calls, total_successful, success_rate)

        return {
            "portfolio_with_vintage": export_portfolio,
            "calls_analysis": calls_analysis,
            "has_calls": has_calls,
            "portfolio_summary": portfolio_summary_df,
            "discounts_table": discounts_table_display,
            "teams_data": teams_data,
            "best_team_name": best_team_name,
            "contact_summary": contact_summary_df,
            "phone_summary": phone_summary_df,
            "vintage_table": vintage_table_display,
            "status_collection_summary": status_coll_df,
            "supervisor_table": sup_df,
            "top_10_collectors": top_10,
            "bottom_10_collectors": bottom_10,
            "all_collectors": col_df_sorted,
            "supervisors_list": sorted([s for s in df_p["_supervisor"].unique() if s != "غير محدد"]),
            "stats": stats,
            "has_payments": has_payments
        }


def export_analytics_html(results: Dict[str, Any]) -> str:
    """
    توليد وتصدير لوحة القيادة التنفيذية كملف HTML مستقل وتفاعلي بنظام التبويبات المتعددة (Tabs)
    وتصميم متناسق (Justified RTL Layout) مع رسومات بيانية واضحة ومبسطة، وجاهزية تامة للطباعة وتصدير PDF.
    """
    stats = results.get("stats", {})

    contact_df = results.get("contact_summary", pd.DataFrame())
    phone_df = results.get("phone_summary", pd.DataFrame())
    port_df = results.get("portfolio_summary", pd.DataFrame())
    vintage_df = results.get("vintage_table", pd.DataFrame())
    status_df = results.get("status_collection_summary", pd.DataFrame())
    sup_df = results.get("supervisor_table", pd.DataFrame())
    top10_df = results.get("top_10_collectors", pd.DataFrame())
    bot10_df = results.get("bottom_10_collectors", pd.DataFrame())
    all_col_df = results.get("all_collectors", pd.DataFrame())
    disc_df = results.get("discounts_table", pd.DataFrame())
    teams_data = results.get("teams_data", {})
    best_team_name = results.get("best_team_name", "غير محدد")

    # ── 1. تجهيز بيانات الرسوم البيانية البسيطة والمباشرة ──
    c_pie_labels, c_pie_values = [], []
    if not contact_df.empty:
        c_sub = contact_df[~contact_df["حالة التواصل"].str.contains("إجمالي|الإجمالي", na=False)].copy()
        c_pie_labels = c_sub["حالة التواصل"].tolist()
        c_pie_values = c_sub["عدد الحسابات (السجلات)"].tolist()

    p_pie_labels, p_pie_values = [], []
    if not phone_df.empty:
        p_pie_labels = phone_df["حالة الرقم"].tolist()
        p_pie_values = phone_df["عدد الحسابات"].tolist()

    port_names, port_debts, port_colls = [], [], []
    if not port_df.empty:
        p_sub = port_df[port_df["المحفظة"] != "الإجمالي"].copy()
        port_names = p_sub["المحفظة"].tolist()
        port_debts = p_sub["مبلغ المديونية"].tolist()
        port_colls = p_sub["إجمالي التحصيل"].tolist()

    v_years, v_debts, v_colls = [], [], []
    if not vintage_df.empty:
        col_yr = "عمر الدين (السنة)" if "عمر الدين (السنة)" in vintage_df.columns else "سنة التعثر"
        v_sub = vintage_df[vintage_df[col_yr] != "الإجمالي"].copy()
        v_sub["_yr_n"] = pd.to_numeric(v_sub[col_yr], errors="coerce")
        v_sub_chart = v_sub[v_sub["_yr_n"].notna()].sort_values("_yr_n", ascending=True).tail(9)
        if v_sub_chart.empty:
            v_sub_chart = v_sub.head(9)
        v_years = v_sub_chart[col_yr].astype(str).tolist()
        v_debts = v_sub_chart["مبلغ المديونية"].tolist()
        v_colls = v_sub_chart["إجمالي التحصيل"].tolist()

    st_names, st_colls = [], []
    if not status_df.empty:
        st_sub = status_df.head(6)
        st_names = st_sub["الحالة الرئيسية"].tolist()
        st_colls = st_sub["إجمالي التحصيل"].tolist()

    top_col_names, top_col_colls = [], []
    if not top10_df.empty:
        t10_rev = top10_df.iloc[::-1]
        top_col_names = t10_rev["المحصل"].tolist()
        top_col_colls = t10_rev["إجمالي التحصيل"].tolist()

    def df_to_html_table(df: pd.DataFrame, max_rows: int = 60) -> str:
        if df is None or df.empty:
            return '<div style="padding:20px; text-align:center; color:#64748b; font-weight:600;">لا توجد بيانات متاحة لعرضها</div>'
        sub = df.head(max_rows)
        html = '<div class="table-responsive"><table><thead><tr>'
        for c in sub.columns:
            html += f'<th>{c}</th>'
        html += '</tr></thead><tbody>'
        for _, row in sub.iterrows():
            is_tot = any("إجمالي" in str(x) or "الإجمالي" in str(x) for x in row.values)
            row_cls = ' class="row-total"' if is_tot else ""
            html += f'<tr{row_cls}>'
            for idx_c, val in enumerate(row.values):
                if isinstance(val, (int, float)):
                    val_str = f"{val:,.2f}" if isinstance(val, float) else f"{val:,}"
                else:
                    val_str = str(val)

                if "🌟" in val_str:
                    cell_html = f'<span class="badge badge-star">{val_str}</span>'
                elif "🎯" in val_str:
                    cell_html = f'<span class="badge badge-target">{val_str}</span>'
                elif "⚡" in val_str:
                    cell_html = f'<span class="badge badge-flash">{val_str}</span>'
                elif "⚠️" in val_str:
                    cell_html = f'<span class="badge badge-alert">{val_str}</span>'
                elif val_str == "تم التوصل" or val_str == "يوجد رقم رئيسي":
                    cell_html = f'<span class="badge badge-success">{val_str}</span>'
                elif val_str == "لا يرد ومغلق":
                    cell_html = f'<span class="badge badge-warning">{val_str}</span>'
                elif val_str == "عدم التوصل" or val_str == "لا يوجد رقم مسجل":
                    cell_html = f'<span class="badge badge-danger">{val_str}</span>'
                else:
                    cell_html = val_str
                html += f'<td>{cell_html}</td>'
            html += '</tr>'
        html += '</tbody></table></div>'
        return html

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    # ── قراءة وحساب أرقام الـ KPIs بدقة 100% ودون أي احتمال لظهور 0 ──
    # 1. إجمالي مبلغ المديونية
    tot_debt_str = stats.get("إجمالي مبلغ المديونية") or ""
    if not tot_debt_str or tot_debt_str in ("0", "0.0", "0.00 ﷼"):
        if not port_df.empty and "مبلغ المديونية" in port_df.columns:
            d_val = port_df[port_df["المحفظة"] == "الإجمالي"]["مبلغ المديونية"].values
            tot_debt_str = f"{d_val[0]:,.2f} ﷼" if len(d_val) > 0 else "0 ﷼"
        else:
            tot_debt_str = "0 ﷼"

    # 2. إجمالي التحصيل الفعلي (حل مشكلة ظهور 0 ريال)
    tot_coll_str = stats.get("إجمالي التحصيل") or stats.get("إجمالي التحصيل الفعلي") or ""
    if not tot_coll_str or tot_coll_str in ("0", "0.0", "0.00 ﷼"):
        if not port_df.empty and "إجمالي التحصيل" in port_df.columns:
            c_val = port_df[port_df["المحفظة"] == "الإجمالي"]["إجمالي التحصيل"].values
            tot_coll_str = f"{c_val[0]:,.2f} ﷼" if len(c_val) > 0 else "0 ﷼"
        else:
            tot_coll_str = "0 ﷼"

    # 3. نسبة التحصيل العامة %
    coll_rate_str = stats.get("نسبة التحصيل العامة %") or stats.get("نسبة التحصيل الإجمالية %") or ""
    if not coll_rate_str or coll_rate_str in ("0", "0%", "0.0%"):
        if not port_df.empty and "نسبة التحصيل %" in port_df.columns:
            r_val = port_df[port_df["المحفظة"] == "الإجمالي"]["نسبة التحصيل %"].values
            coll_rate_str = f"{r_val[0]}%" if len(r_val) > 0 else "0%"
        else:
            coll_rate_str = "0%"

    # 4. نسبة التغطية الشاملة
    cov_rate_str = stats.get("نسبة التغطية الكلية %") or stats.get("نسبة التغطية الشاملة %") or "0%"
    if "%" not in str(cov_rate_str):
        cov_rate_str = f"{cov_rate_str}%"

    # 5. الحسابات والعملاء
    tot_recs_str = stats.get("إجمالي الحسابات") or stats.get("إجمالي الحسابات (السجلات)") or ""
    if not tot_recs_str:
        if not port_df.empty and "عدد الحسابات" in port_df.columns:
            tot_recs_str = f"{port_df[port_df['المحفظة'] == 'الإجمالي']['عدد الحسابات'].values[0]:,} حساب"
        else:
            tot_recs_str = "0"

    tot_cust_str = stats.get("إجمالي العملاء") or stats.get("عدد العملاء (هوية فريدة)") or ""
    if not tot_cust_str:
        if not port_df.empty and "عدد العملاء" in port_df.columns:
            tot_cust_str = f"{port_df[port_df['المحفظة'] == 'الإجمالي']['عدد العملاء'].values[0]:,} عميل"
        else:
            tot_cust_str = "0"

    top_stat_str = str(stats.get("أعلى حالة سداداً", "غير محدد"))
    top_port_str = str(stats.get("أعلى محفظة تحصيلاً", "غير محدد"))

    # تجهيز محتوى تبويب أداء فرق العمل (HTML)
    teams_html = ""
    if teams_data:
        teams_html += f"""
        <div style="background:linear-gradient(135deg, #10b981 0%, #059669 100%); color:#ffffff; border-radius:14px; padding:20px 24px; margin-bottom:24px; box-shadow:var(--shadow-md);">
            <div style="font-size:18px; font-weight:800; margin-bottom:4px;">🏆 أفضل فريق عمل تحصيلاً: فريق المشرف ({best_team_name})</div>
            <div style="font-size:13px; color:#d1fae5;">تم اختيار هذا الفريق بناءً على كفاءة التحصيل المحقق والتواصل الإيجابي مع العملاء المتعثرين.</div>
        </div>
        """
        for s_name, t_info in teams_data.items():
            t_cols_count = t_info["collectors_count"]
            t_debt = f"{t_info['debt_sum']:,.2f} ﷼"
            t_coll = f"{t_info['coll_sum']:,.2f} ﷼"
            t_rate = f"{t_info['coll_rate']}%"
            t_cont = f"{t_info['cont_rate']}%"
            t_cov  = f"{t_info['cov_rate']}%"

            teams_html += f"""
            <div class="section-card" style="margin-bottom:26px;">
                <div class="section-header">
                    <div class="sec-title">👥 فريق المشرف: <span style="color:var(--navy-primary); font-weight:900;">{s_name}</span></div>
                    <div style="display:flex; gap:8px; flex-wrap:wrap;">
                        <span class="sec-badge">عدد المحصلين: {t_cols_count}</span>
                        <span class="sec-badge" style="background:#ecfdf5; color:#059669;">تحصيل الفريق: {t_coll} ({t_rate})</span>
                        <span class="sec-badge" style="background:#eff6ff; color:#1e40af;">نسبة التوصل: {t_cont}</span>
                        <span class="sec-badge" style="background:#fffbeb; color:#d97706;">نسبة التغطية: {t_cov}</span>
                    </div>
                </div>

                <div class="grid-2col" style="margin-bottom:18px;">
                    <div>
                        <h4 style="color:#059669; font-weight:800; margin-bottom:8px;">🥇 أفضل 5 محصلين في الفريق (الأعلى تحصيلاً):</h4>
                        {df_to_html_table(t_info['top_5'], max_rows=5)}
                    </div>
                    <div>
                        <h4 style="color:#dc2626; font-weight:800; margin-bottom:8px;">⚠️ أقل 5 محصلين في الفريق (يحتاجون دعم ومتابعة):</h4>
                        {df_to_html_table(t_info['bottom_5'], max_rows=5)}
                    </div>
                </div>

                <div style="margin-top:14px;">
                    <details style="cursor:pointer;">
                        <summary style="font-weight:700; color:var(--navy-primary); padding:8px 0;">📋 عرض القائمة الكاملة لكافة أعضاء فريق {s_name} ({t_cols_count} محصل)</summary>
                        <div style="margin-top:10px;">
                            {df_to_html_table(t_info['all_members'], max_rows=100)}
                        </div>
                    </details>
                </div>
            </div>
            """
    else:
        teams_html = '<div class="section-card"><p style="text-align:center; color:var(--text-muted);">لا توجد بيانات كافية لفرق العمل</p></div>'

    has_calls_html = bool(results.get("has_calls"))
    calls_data_html = results.get("calls_analysis", {}) if has_calls_html else {}
    calls_tab_btn_html = '<button class="tab-btn" data-target="tab-calls" onclick="switchTab(\'tab-calls\')">📞 تحليل المكالمات</button>' if has_calls_html else ''
    if has_calls_html:
        _c_tot = calls_data_html.get("total_calls", 0)
        _c_suc = calls_data_html.get("total_successful", 0)
        _c_fail = calls_data_html.get("total_failed", 0)
        _c_rate = calls_data_html.get("success_rate", 0.0)
        _c_thresh = calls_data_html.get("threshold_seconds", 30)
        _c_avg_all = calls_data_html.get("avg_dur_all")
        _c_avg_suc = calls_data_html.get("avg_dur_success")
        _c_top_dur = calls_data_html.get("top_caller_by_dur") or "-"
        _c_bot_dur = calls_data_html.get("bottom_caller_by_dur") or "-"
        _c_sup_df = calls_data_html.get("supervisor_calls_df", pd.DataFrame())
        _c_col_df = calls_data_html.get("collector_calls_df", pd.DataFrame())
        _dur_kpis_html = ""
        if calls_data_html.get("has_duration_col"):
            _dur_kpis_html = f"""
            <div class="kpi-grid" style="margin-top:12px;">
                <div class="kpi-card"><div class="kpi-label">متوسط مدة كل المكالمات</div><div class="kpi-value">{f'{_c_avg_all:.1f} ث' if _c_avg_all is not None else '-'}</div></div>
                <div class="kpi-card"><div class="kpi-label">متوسط مدة الناجحة</div><div class="kpi-value">{f'{_c_avg_suc:.1f} ث' if _c_avg_suc is not None else '-'}</div></div>
                <div class="kpi-card"><div class="kpi-label">أعلى محصل (متوسط المدة)</div><div class="kpi-value" style="font-size:1.1rem;">{_c_top_dur}</div></div>
                <div class="kpi-card"><div class="kpi-label">أقل محصل (متوسط المدة)</div><div class="kpi-value" style="font-size:1.1rem;">{_c_bot_dur}</div></div>
            </div>"""
        calls_tab_pane_html = f"""
        <!-- ==================== TAB 9: CALLS ANALYSIS ==================== -->
        <div id="tab-calls" class="tab-pane">
            <div class="section-card">
                <div class="section-header">
                    <div class="sec-title">📞 تحليل المكالمات الفعلية (الحد: &gt;{_c_thresh} ثانية)</div>
                    <div class="sec-badge">شيت المكالمات</div>
                </div>
                <div class="kpi-grid">
                    <div class="kpi-card"><div class="kpi-label">إجمالي المكالمات</div><div class="kpi-value">{_c_tot:,}</div></div>
                    <div class="kpi-card"><div class="kpi-label">المكالمات الناجحة</div><div class="kpi-value">{_c_suc:,}</div></div>
                    <div class="kpi-card"><div class="kpi-label">المكالمات غير الناجحة</div><div class="kpi-value">{_c_fail:,}</div></div>
                    <div class="kpi-card"><div class="kpi-label">نسبة النجاح الكلية</div><div class="kpi-value">{_c_rate:.1f}%</div></div>
                </div>
                {_dur_kpis_html}
                <h3 style="margin-top:22px; color:#1E3A8A;">👥 أداء المشرفين في المكالمات</h3>
                <div style="margin-top:10px;">{df_to_html_table(_c_sup_df)}</div>
                <h3 style="margin-top:22px; color:#1E3A8A;">🧑‍💼 أداء المحصلين في المكالمات</h3>
                <div style="margin-top:10px;">{df_to_html_table(_c_col_df)}</div>
            </div>
        </div>"""
    else:
        calls_tab_pane_html = ""

    html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>تحليل شامل لشركة الاتصالات السعودية (STC) | شركة مهارة للتحصيل</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
    <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
    <style>
        :root {{
            --bg: #f8fafc;
            --surface: #ffffff;
            --border: #e2e8f0;
            --border-light: #f1f5f9;
            --navy-primary: #1e3a8a;
            --navy-dark: #0f172a;
            --blue-accent: #2563eb;
            --blue-soft: #eff6ff;
            --emerald: #10b981;
            --emerald-soft: #ecfdf5;
            --amber: #f59e0b;
            --amber-soft: #fffbeb;
            --rose: #ef4444;
            --rose-soft: #fef2f2;
            --text-dark: #0f172a;
            --text-body: #334155;
            --text-muted: #64748b;
            --shadow-sm: 0 1px 3px rgba(0,0,0,0.06);
            --shadow-md: 0 4px 14px rgba(0,0,0,0.05);
            --shadow-lg: 0 10px 25px rgba(0,0,0,0.07);
        }}
        * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: 'Cairo', 'Segoe UI', Tahoma, sans-serif; }}
        body {{
            background: var(--bg);
            color: var(--text-body);
            direction: rtl;
            font-size: 14px;
            line-height: 1.6;
            -webkit-font-smoothing: antialiased;
        }}
        .app-container {{
            max-width: 1440px;
            margin: 0 auto;
            padding: 24px 30px;
        }}

        /* Header */
        .header-bar {{
            background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 50%, #0f172a 100%);
            color: #ffffff;
            border-radius: 16px;
            padding: 26px 32px;
            margin-bottom: 22px;
            box-shadow: var(--shadow-lg);
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 16px;
        }}
        .header-title h1 {{
            font-size: 22px;
            font-weight: 800;
            margin-bottom: 6px;
            letter-spacing: -0.3px;
        }}
        .header-title p {{
            font-size: 13px;
            color: #bfdbfe;
        }}
        .header-meta {{
            text-align: left;
            display: flex;
            flex-direction: column;
            gap: 6px;
        }}
        .badge-live {{
            background: rgba(16, 185, 129, 0.2);
            color: #6ee7b7;
            border: 1px solid rgba(16, 185, 129, 0.4);
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 700;
            display: inline-block;
        }}

        /* Navigation Tabs */
        .tabs-nav {{
            position: sticky;
            top: 10px;
            z-index: 999;
            background: #ffffff;
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 8px 12px;
            margin-bottom: 24px;
            box-shadow: var(--shadow-md);
            display: flex;
            align-items: center;
            gap: 8px;
            overflow-x: auto;
            white-space: nowrap;
        }}
        .tab-btn {{
            background: transparent;
            border: none;
            outline: none;
            padding: 10px 18px;
            border-radius: 10px;
            font-size: 13.5px;
            font-weight: 700;
            color: var(--text-muted);
            cursor: pointer;
            transition: all 0.2s ease;
            display: flex;
            align-items: center;
            gap: 6px;
        }}
        .tab-btn:hover {{
            background: var(--blue-soft);
            color: var(--blue-accent);
        }}
        .tab-btn.active {{
            background: var(--navy-primary);
            color: #ffffff;
            box-shadow: 0 2px 8px rgba(30, 58, 138, 0.25);
        }}
        .btn-print {{
            margin-right: auto;
            background: #f1f5f9;
            color: var(--navy-dark);
            border: 1px solid var(--border);
        }}
        .btn-print:hover {{
            background: #e2e8f0;
            color: var(--navy-primary);
        }}

        /* Tab Panes */
        .tab-pane {{
            display: none;
            animation: fadeIn 0.3s ease;
        }}
        .tab-pane.active {{
            display: block;
        }}
        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(6px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        /* KPI Metric Cards */
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }}
        .kpi-card {{
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 20px 22px;
            box-shadow: var(--shadow-sm);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
            position: relative;
            overflow: hidden;
        }}
        .kpi-card:hover {{
            transform: translateY(-2px);
            box-shadow: var(--shadow-md);
        }}
        .kpi-card::before {{
            content: "";
            position: absolute;
            top: 0;
            right: 0;
            left: 0;
            height: 4px;
            background: var(--navy-primary);
        }}
        .kpi-card.emerald::before {{ background: var(--emerald); }}
        .kpi-card.blue::before {{ background: var(--blue-accent); }}
        .kpi-card.amber::before {{ background: var(--amber); }}
        .kpi-card.purple::before {{ background: #8b5cf6; }}
        .kpi-label {{
            font-size: 13px;
            font-weight: 700;
            color: var(--text-muted);
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            gap: 6px;
        }}
        .kpi-val {{
            font-size: 24px;
            font-weight: 900;
            color: var(--text-dark);
            margin-bottom: 4px;
            letter-spacing: -0.5px;
        }}
        .kpi-sub {{
            font-size: 12px;
            color: var(--text-muted);
            font-weight: 500;
        }}

        /* Section Cards */
        .section-card {{
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 24px;
            margin-bottom: 24px;
            box-shadow: var(--shadow-sm);
        }}
        .section-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid var(--border-light);
            padding-bottom: 14px;
            margin-bottom: 20px;
            flex-wrap: wrap;
            gap: 10px;
        }}
        .sec-title {{
            font-size: 17px;
            font-weight: 800;
            color: var(--navy-dark);
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .sec-badge {{
            background: var(--blue-soft);
            color: var(--blue-accent);
            padding: 4px 12px;
            border-radius: 8px;
            font-size: 12px;
            font-weight: 700;
        }}

        .chart-container {{
            width: 100%;
            border-radius: 10px;
            margin-bottom: 16px;
        }}

        /* Tables - Justified and Modern */
        .table-responsive {{
            width: 100%;
            overflow-x: auto;
            border: 1px solid var(--border);
            border-radius: 10px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            text-align: center;
            font-size: 13px;
        }}
        th {{
            background: var(--navy-primary);
            color: #ffffff;
            font-weight: 700;
            padding: 12px 14px;
            white-space: nowrap;
            border-left: 1px solid rgba(255, 255, 255, 0.1);
        }}
        td {{
            padding: 11px 14px;
            border-bottom: 1px solid var(--border-light);
            border-left: 1px solid var(--border-light);
            white-space: nowrap;
            color: var(--text-body);
        }}
        tr:nth-child(even) td {{
            background: #fafbfc;
        }}
        tr:hover td {{
            background: #f1f5f9;
        }}
        tr.row-total td {{
            background: var(--blue-soft) !important;
            color: var(--navy-primary) !important;
            font-weight: 800 !important;
            border-top: 2px solid var(--blue-accent) !important;
            font-size: 13.5px;
        }}

        .badge {{
            display: inline-block;
            padding: 3px 10px;
            border-radius: 6px;
            font-weight: 700;
            font-size: 11.5px;
        }}
        .badge-success {{ background: var(--emerald-soft); color: var(--emerald); }}
        .badge-warning {{ background: var(--amber-soft); color: var(--amber); }}
        .badge-danger  {{ background: var(--rose-soft); color: var(--rose); }}
        .badge-star    {{ background: #eff6ff; color: #1e3a8a; border: 1px solid #bfdbfe; }}
        .badge-target  {{ background: #f0fdf4; color: #15803d; border: 1px solid #bbf7d0; }}
        .badge-flash   {{ background: #fefce8; color: #a16207; border: 1px solid #fef08a; }}
        .badge-alert   {{ background: #fef2f2; color: #b91c1c; border: 1px solid #fecaca; }}

        .grid-2col {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
        }}
        @media (max-width: 900px) {{
            .grid-2col {{ grid-template-columns: 1fr; }}
        }}

        /* Print Styles */
        @media print {{
            body {{ background: #ffffff !important; padding: 0 !important; font-size: 11px !important; }}
            .app-container {{ max-width: 100% !important; padding: 0 !important; }}
            .tabs-nav, .btn-print {{ display: none !important; }}
            .tab-pane {{ display: block !important; page-break-after: always; }}
            .header-bar {{ box-shadow: none !important; border-radius: 0 !important; padding: 15px !important; }}
            .section-card {{ box-shadow: none !important; border: 1px solid #cbd5e1 !important; margin-bottom: 15px !important; padding: 15px !important; }}
            th {{ background: #1e3a8a !important; color: #ffffff !important; -webkit-print-color-adjust: exact; }}
            tr.row-total td {{ background: #eff6ff !important; -webkit-print-color-adjust: exact; }}
        }}
    </style>
</head>
<body>
    <div class="app-container">
        <!-- Top Header -->
        <header class="header-bar">
            <div class="header-title">
                <h1>🏛️ تحليل شامل لشركة الاتصالات السعودية (STC) — اللوحة التنفيذية لتحليل المحفظة وعمر الدين</h1>
                <p>تحليل موحد وموثق يربط المحفظة بمدفوعات التحصيل برقم المديونية وحساب عدد العملاء بالهوية الفريدة</p>
            </div>
            <div class="header-meta">
                <span class="badge-live">● تقرير تنفيذي رسمي جاهز</span>
                <span style="font-size:12px; color:#cbd5e1;">تاريخ التصدير: {now_str}</span>
            </div>
        </header>

        <!-- Sticky Navigation Tabs -->
        <nav class="tabs-nav">
            <button class="tab-btn active" data-target="tab-overview" onclick="switchTab('tab-overview')">📊 النظرة العامة والمؤشرات</button>
            <button class="tab-btn" data-target="tab-portfolios" onclick="switchTab('tab-portfolios')">📁 تحليل المحافظ</button>
            <button class="tab-btn" data-target="tab-vintage" onclick="switchTab('tab-vintage')">⏳ عمر الدين وسنة التعثر</button>
            <button class="tab-btn" data-target="tab-teams" onclick="switchTab('tab-teams')">🏆 أداء فرق العمل لكل مشرف</button>
            <button class="tab-btn" data-target="tab-team" onclick="switchTab('tab-team')">👥 أداء المشرفين والمحصلين</button>
            <button class="tab-btn" data-target="tab-discounts" onclick="switchTab('tab-discounts')">🎯 عملاء الخصومات</button>
            <button class="tab-btn" data-target="tab-contact" onclick="switchTab('tab-contact')">📞 التواصل وجودة الأرقام</button>
            <button class="tab-btn" data-target="tab-status" onclick="switchTab('tab-status')">📑 التحصيل حسب الحالة</button>
{calls_tab_btn_html}
            <button class="tab-btn btn-print" onclick="window.print()">🖨️ طباعة التقرير / PDF</button>
        </nav>

        <!-- ==================== TAB 1: OVERVIEW ==================== -->
        <div id="tab-overview" class="tab-pane active">
            <!-- KPIs -->
            <div class="kpi-grid">
                <div class="kpi-card blue">
                    <div class="kpi-label">💰 إجمالي مبلغ المديونية (المحفظة)</div>
                    <div class="kpi-val">{tot_debt_str}</div>
                    <div class="kpi-sub">عدد الحسابات: {tot_recs_str} | العملاء: {tot_cust_str}</div>
                </div>
                <div class="kpi-card emerald">
                    <div class="kpi-label">💵 إجمالي السدادات المحققة (التحصيل)</div>
                    <div class="kpi-val">{tot_coll_str}</div>
                    <div class="kpi-sub">نسبة التحصيل من المديونية: {coll_rate_str}</div>
                </div>
                <div class="kpi-card amber">
                    <div class="kpi-label">📞 نسبة التغطية الشاملة</div>
                    <div class="kpi-val">{cov_rate_str}</div>
                    <div class="kpi-sub">استناداً لعمود تاريخ المتابعة</div>
                </div>
                <div class="kpi-card purple">
                    <div class="kpi-label">📁 عدد المحافظ في الملف</div>
                    <div class="kpi-val">{len(port_names)} محافظ</div>
                    <div class="kpi-sub">الأعلى تحصيلاً: {top_port_str}</div>
                </div>
            </div>

            <!-- Overview Simple Charts Grid -->
            <div class="grid-2col">
                <div class="section-card">
                    <div class="section-header">
                        <div class="sec-title">📞 نسب حالات التواصل الرئيسية</div>
                        <div class="sec-badge">قواعد تصنيف الاتصال</div>
                    </div>
                    <div id="chart_ov_contact" class="chart-container" style="height:320px;"></div>
                </div>
                <div class="section-card">
                    <div class="section-header">
                        <div class="sec-title">📁 مقارنة تحصيل المحافظ</div>
                        <div class="sec-badge">سداد موثق</div>
                    </div>
                    <div id="chart_ov_port" class="chart-container" style="height:320px;"></div>
                </div>
            </div>

            <!-- Portfolio Summary Overview Table -->
            <div class="section-card">
                <div class="section-header">
                    <div class="sec-title">📋 ملخص مؤشرات المحافظ الأساسية</div>
                    <div class="sec-badge">{len(port_names)} محفظة</div>
                </div>
                {df_to_html_table(port_df)}
            </div>
        </div>

        <!-- ==================== TAB 2: PORTFOLIOS ==================== -->
        <div id="tab-portfolios" class="tab-pane">
            <div class="section-card">
                <div class="section-header">
                    <div class="sec-title">📁 التحليل المالي والتشغيلي للمحافظ</div>
                    <div class="sec-badge">المديونية مقابل التحصيل ونسب التوصل</div>
                </div>
                <p style="color:var(--text-muted); margin-bottom:16px;">
                    يقارن هذا الرسم مبالغ المديونية المسندة بكل محفظة مع حجم المبالغ المحصلة فعلياً من شيت التحصيل.
                </p>
                <div id="chart_portfolios_full" class="chart-container" style="height:380px;"></div>
                <div style="margin-top:20px;">
                    {df_to_html_table(port_df)}
                </div>
            </div>
        </div>

        <!-- ==================== TAB 3: VINTAGE ==================== -->
        <div id="tab-vintage" class="tab-pane">
            <div class="section-card">
                <div class="section-header">
                    <div class="sec-title">⏳ تحليل سنة التعثر وعمر الدين (مقارنة التحصيل بالمديونية)</div>
                    <div class="sec-badge">{stats.get('عمود عمر الدين المستخدم', 'تاريخ فصل الخدمة / تاريخ الغاء الخدمة')}</div>
                </div>
                <p style="color:var(--text-muted); margin-bottom:16px;">
                    تم استخراج السنوات حصراً من كولوم تاريخ فصل الخدمة وتاريخ إلغاء الخدمة الموجودين بملف المحفظة، ويعرض الرسم البياني أدناه أحدث وأهم السنوات مباشرة لتسهيل اتخاذ القرار.
                </p>
                <div id="chart_vintage_simple" class="chart-container" style="height:380px;"></div>
                <div style="margin-top:20px;">
                    <div class="sec-title" style="margin-bottom:12px; font-size:15px;">📊 جدول عمر الدين التفصيلي الكامل لكافة السنوات الموجودة بالمحفظة:</div>
                    {df_to_html_table(vintage_df)}
                </div>
            </div>
        </div>

        <!-- ==================== TAB 4: TEAMS BY SUPERVISOR ==================== -->
        <div id="tab-teams" class="tab-pane">
            <div class="section-card" style="margin-bottom:20px;">
                <div class="section-header">
                    <div class="sec-title">🏆 أداء ومقارنة فرق العمل تحت إشراف كل مشرف</div>
                    <div class="sec-badge">Best Teams Evaluation</div>
                </div>
                <p style="color:var(--text-muted);">
                    توزيع المحصلين بالكامل تحت إشراف كل مشرف، مع بيان أفضل وأقل 5 محصلين في كل فريق ونسب التوصل والإغلاق وعدم التوصل.
                </p>
            </div>
            {teams_html}
        </div>

        <!-- ==================== TAB 5: TEAM PERFORMANCE ==================== -->
        <div id="tab-team" class="tab-pane">
            <!-- Supervisors Section -->
            <div class="section-card">
                <div class="section-header">
                    <div class="sec-title">👥 تقييم ومقارنة أداء المشرفين الشامل (التغطية، التحصيل، وحالات التواصل)</div>
                    <div class="sec-badge">{stats.get('عدد المشرفين', '0')} مشرفين</div>
                </div>
                <div id="chart_sup_simple" class="chart-container" style="height:320px;"></div>
                <div style="margin-top:16px;">
                    {df_to_html_table(sup_df)}
                </div>
            </div>

            <!-- Collectors Section -->
            <div class="section-card">
                <div class="section-header">
                    <div class="sec-title">🥇 أفضل 10 محصلين تحصيلاً</div>
                    <div class="sec-badge">كفاءة التحصيل</div>
                </div>
                <div id="chart_top10_bar" class="chart-container" style="height:360px;"></div>
                <div class="grid-2col" style="margin-top:20px;">
                    <div>
                        <h4 style="color:var(--emerald); margin-bottom:10px; font-weight:800;">🥇 جدول أفضل 10 محصلين تحصيلاً</h4>
                        {df_to_html_table(top10_df)}
                    </div>
                    <div>
                        <h4 style="color:var(--rose); margin-bottom:10px; font-weight:800;">⚠️ جدول أقل 10 محصلين تحصيلاً</h4>
                        {df_to_html_table(bot10_df)}
                    </div>
                </div>
                <div style="margin-top:24px;">
                    <h4 style="color:var(--navy-primary); margin-bottom:10px; font-weight:800;">📋 القائمة الكاملة لكافة المحصلين مع نسب التوصل والإغلاق والربع (Quadrants)</h4>
                    {df_to_html_table(all_col_df, max_rows=150)}
                </div>
            </div>
        </div>

        <!-- ==================== TAB 6: DISCOUNTS ==================== -->
        <div id="tab-discounts" class="tab-pane">
            <div class="section-card">
                <div class="section-header">
                    <div class="sec-title">🎯 عملاء وحسابات الخصومات المعتمدة</div>
                    <div class="sec-badge">Discount Approved Accounts</div>
                </div>
                <p style="color:var(--text-muted); margin-bottom:16px;">
                    كشف الحسابات التي تمت الموافقة على خصومات لها، ومقارنة مبالغ المديونية بمبالغ السداد المحققة والمتبقي بعد السداد مع صف إجمالي في النهاية:
                </p>
                {df_to_html_table(disc_df, max_rows=100)}
            </div>
        </div>

        <!-- ==================== TAB 7: CONTACT & PHONE ==================== -->
        <div id="tab-contact" class="tab-pane">
            <div class="grid-2col">
                <div class="section-card">
                    <div class="section-header">
                        <div class="sec-title">📞 1. تحليل حالات التواصل</div>
                        <div class="sec-badge">الحالة الرئيسية والفرعية</div>
                    </div>
                    <div id="chart_contact_tab_pie" class="chart-container" style="height:300px;"></div>
                    {df_to_html_table(contact_df)}
                </div>
                <div class="section-card">
                    <div class="section-header">
                        <div class="sec-title">📱 2. كاشف توفر الأرقام الرئيسية</div>
                        <div class="sec-badge">الرقم الرئيسي</div>
                    </div>
                    <div id="chart_phone_tab_pie" class="chart-container" style="height:300px;"></div>
                    {df_to_html_table(phone_df)}
                </div>
            </div>
        </div>

        <!-- ==================== TAB 8: MAIN STATUS COLLECTIONS ==================== -->
        <div id="tab-status" class="tab-pane">
            <div class="section-card">
                <div class="section-header">
                    <div class="sec-title">📑 التحصيل الفعلي ومساهمة كل حالة رئيسية في السداد</div>
                    <div class="sec-badge">أعلى الحالات مساهمة</div>
                </div>
                <p style="color:var(--text-muted); margin-bottom:16px;">
                    أعلى حالة حققت تحصيلاً: <strong>{top_stat_str}</strong>
                </p>
                <div id="chart_status_tab_bar" class="chart-container" style="height:340px;"></div>
                <div style="margin-top:20px;">
                    {df_to_html_table(status_df)}
                </div>
            </div>
        </div>

{calls_tab_pane_html}

        <!-- Footer -->
        <footer style="text-align:center; padding:24px; color:var(--text-muted); font-size:12.5px; border-top:1px solid var(--border); margin-top:20px;">
            نظام التحليل الاستراتيجي وعمر الدين © شركة مهارة للتحصيل | تم التوليد آلياً بدقة 100% | {now_str}
        </footer>
    </div>

    <!-- JavaScript Interactive Scripts -->
    <script>
        function switchTab(tabId) {{
            document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));

            const target = document.getElementById(tabId);
            if (target) {{
                target.classList.add('active');
                window.dispatchEvent(new Event('resize'));
                setTimeout(() => {{
                    const plots = target.querySelectorAll('.js-plotly-plot');
                    plots.forEach(p => Plotly.Plots.resize(p));
                }}, 60);
            }}
            const activeBtn = document.querySelector(`[data-target="${{tabId}}"]`);
            if (activeBtn) activeBtn.classList.add('active');
        }}

        const cleanLayout = {{
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(0,0,0,0)',
            font: {{ family: 'Cairo, sans-serif', color: '#334155', size: 12 }},
            margin: {{ l: 40, r: 20, t: 30, b: 40 }},
            legend: {{ orientation: 'h', yanchor: 'bottom', y: 1.02, xanchor: 'right', x: 1 }}
        }};

        Plotly.newPlot('chart_ov_contact', [{{
            labels: {json.dumps(c_pie_labels, ensure_ascii=False)},
            values: {json.dumps(c_pie_values)},
            type: 'pie',
            hole: 0.55,
            marker: {{ colors: ['#10b981', '#f59e0b', '#ef4444', '#64748b'] }},
            textinfo: 'percent+label',
            textposition: 'inside',
            hoverinfo: 'label+value+percent'
        }}], {{
            ...cleanLayout,
            margin: {{ l: 10, r: 10, t: 10, b: 10 }},
            showlegend: false
        }}, {{ responsive: true }});

        Plotly.newPlot('chart_ov_port', [{{
            x: {json.dumps(port_names, ensure_ascii=False)},
            y: {json.dumps(port_colls)},
            type: 'bar',
            name: 'إجمالي التحصيل',
            marker: {{ color: '#10b981' }}
        }}], {{
            ...cleanLayout,
            xaxis: {{ tickangle: -15 }},
            yaxis: {{ title: 'التحصيل (ريال)' }}
        }}, {{ responsive: true }});

        Plotly.newPlot('chart_portfolios_full', [
            {{
                x: {json.dumps(port_names, ensure_ascii=False)},
                y: {json.dumps(port_debts)},
                type: 'bar',
                name: 'مبلغ المديونية',
                marker: {{ color: '#1e3a8a' }}
            }},
            {{
                x: {json.dumps(port_names, ensure_ascii=False)},
                y: {json.dumps(port_colls)},
                type: 'bar',
                name: 'إجمالي التحصيل',
                marker: {{ color: '#38bdf8' }}
            }}
        ], {{
            ...cleanLayout,
            barmode: 'group',
            xaxis: {{ tickangle: -15 }},
            yaxis: {{ title: 'المبلغ (ريال)' }}
        }}, {{ responsive: true }});

        Plotly.newPlot('chart_vintage_simple', [
            {{
                x: {json.dumps(v_years, ensure_ascii=False)},
                y: {json.dumps(v_debts)},
                type: 'bar',
                name: 'مبلغ المديونية',
                marker: {{ color: '#1e3a8a' }}
            }},
            {{
                x: {json.dumps(v_years, ensure_ascii=False)},
                y: {json.dumps(v_colls)},
                type: 'bar',
                name: 'إجمالي التحصيل',
                marker: {{ color: '#10b981' }}
            }}
        ], {{
            ...cleanLayout,
            barmode: 'group',
            xaxis: {{ title: 'سنة عمر الدين / التعثر' }},
            yaxis: {{ title: 'المبلغ (ريال)' }}
        }}, {{ responsive: true }});

        Plotly.newPlot('chart_sup_simple', [{{
            x: {json.dumps(sup_df['المشرف'].tolist() if not sup_df.empty else [], ensure_ascii=False)},
            y: {json.dumps(sup_df['إجمالي التحصيل'].tolist() if not sup_df.empty else [])},
            type: 'bar',
            name: 'تحصيل المشرف',
            marker: {{ color: '#2563eb' }}
        }}], {{
            ...cleanLayout,
            xaxis: {{ tickangle: -15 }},
            yaxis: {{ title: 'إجمالي التحصيل (ريال)' }}
        }}, {{ responsive: true }});

        Plotly.newPlot('chart_top10_bar', [{{
            x: {json.dumps(top_col_colls)},
            y: {json.dumps(top_col_names, ensure_ascii=False)},
            type: 'bar',
            orientation: 'h',
            marker: {{ color: '#10b981' }}
        }}], {{
            ...cleanLayout,
            margin: {{ l: 140, r: 20, t: 20, b: 30 }},
            xaxis: {{ title: 'إجمالي التحصيل (ريال)' }}
        }}, {{ responsive: true }});

        Plotly.newPlot('chart_contact_tab_pie', [{{
            labels: {json.dumps(c_pie_labels, ensure_ascii=False)},
            values: {json.dumps(c_pie_values)},
            type: 'pie',
            hole: 0.5,
            marker: {{ colors: ['#10b981', '#f59e0b', '#ef4444'] }}
        }}], {{ ...cleanLayout, margin: {{ l: 10, r: 10, t: 20, b: 20 }} }}, {{ responsive: true }});

        Plotly.newPlot('chart_phone_tab_pie', [{{
            labels: {json.dumps(p_pie_labels, ensure_ascii=False)},
            values: {json.dumps(p_pie_values)},
            type: 'pie',
            hole: 0.5,
            marker: {{ colors: ['#10b981', '#ef4444'] }}
        }}], {{ ...cleanLayout, margin: {{ l: 10, r: 10, t: 20, b: 20 }} }}, {{ responsive: true }});

        Plotly.newPlot('chart_status_tab_bar', [{{
            x: {json.dumps(st_names, ensure_ascii=False)},
            y: {json.dumps(st_colls)},
            type: 'bar',
            marker: {{ color: '#6366f1' }}
        }}], {{
            ...cleanLayout,
            xaxis: {{ tickangle: -20 }},
            yaxis: {{ title: 'إجمالي التحصيل (ريال)' }}
        }}, {{ responsive: true }});
    </script>
</body>
</html>"""
    return html


def export_analytics_excel_bytes(results: Dict[str, Any]) -> bytes:
    """تصدير التقرير الشامل في شيت واحد فقط بتنسيق RTL وأعمدة متناسقة"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        hdr_fmt = workbook.add_format({
            "bold": True, "bg_color": "#1E3A8A", "font_color": "#FFFFFF",
            "border": 1, "align": "center", "valign": "vcenter", "font_size": 11
        })
        section_fmt = workbook.add_format({
            "bold": True, "bg_color": "#2563EB", "font_color": "#FFFFFF",
            "border": 1, "align": "right", "valign": "vcenter", "font_size": 12
        })
        cell_fmt = workbook.add_format({
            "border": 1, "align": "center", "valign": "vcenter", "font_size": 10
        })
        total_fmt = workbook.add_format({
            "bold": True, "bg_color": "#DBEAFE", "border": 1,
            "align": "center", "valign": "vcenter", "font_size": 10
        })

        sheet_name = "التقرير الشامل"
        worksheet = workbook.add_worksheet(sheet_name)
        worksheet.right_to_left()
        worksheet.set_zoom(90)

        current_row = 0

        def write_section_header(title: str, num_cols: int):
            nonlocal current_row
            worksheet.merge_range(current_row, 0, current_row, max(num_cols - 1, 1), title, section_fmt)
            worksheet.set_row(current_row, 22)
            current_row += 1

        def write_dataframe(df: pd.DataFrame, is_total_last: bool = False):
            nonlocal current_row
            if df is None or df.empty:
                worksheet.write(current_row, 0, "لا توجد بيانات مسجلة", cell_fmt)
                current_row += 2
                return

            for col_idx, col_name in enumerate(df.columns):
                worksheet.write(current_row, col_idx, str(col_name), hdr_fmt)
                val_lens = [len(str(x)) for x in df[col_name].head(200).values]
                max_len = max([len(str(col_name))] + val_lens) if val_lens else len(str(col_name))
                worksheet.set_column(col_idx, col_idx, min(max(max_len + 2, 12), 40))
            current_row += 1

            for row_idx, (_, row) in enumerate(df.iterrows()):
                is_last = is_total_last and row_idx == len(df) - 1
                fmt = total_fmt if is_last else cell_fmt
                for col_idx, val in enumerate(row.values):
                    worksheet.write(current_row, col_idx, val, fmt)
                current_row += 1
            current_row += 1

        # 1. حالات التواصل
        c_df = results.get("contact_summary")
        write_section_header("1. ملخص حالات التواصل الشاملة", len(c_df.columns) if c_df is not None and not c_df.empty else 5)
        write_dataframe(c_df, is_total_last=True)

        # 2. كاشف الأرقام
        p_df = results.get("phone_summary")
        write_section_header("2. كاشف توفر الأرقام الرئيسية", len(p_df.columns) if p_df is not None and not p_df.empty else 9)
        write_dataframe(p_df)

        # 3. تحليل المحافظ
        port_s_df = results.get("portfolio_summary")
        if port_s_df is not None and not port_s_df.empty:
            write_section_header("3. تحليل المحافظ المالية (المبالغ، التحصيل، التغطية، ونسب التواصل)", len(port_s_df.columns))
            write_dataframe(port_s_df, is_total_last=True)

        # 4. عملاء الخصومات
        disc_s_df = results.get("discounts_table")
        if disc_s_df is not None and not disc_s_df.empty:
            write_section_header("4. كشف عملاء وحسابات الخصومات المعتمدة", len(disc_s_df.columns))
            write_dataframe(disc_s_df, is_total_last=True)

        # 5. سنة التعثر
        v_df = results.get("vintage_table")
        write_section_header("5. تحليل سنة التعثر وعمر الدين (مقارنة التحصيل بالمديونية)", len(v_df.columns) if v_df is not None and not v_df.empty else 9)
        write_dataframe(v_df, is_total_last=True)

        # 6. التحصيل حسب الحالة الرئيسية
        st_df = results.get("status_collection_summary")
        write_section_header("6. التحصيل الفعلي حسب الحالة الرئيسية", len(st_df.columns) if st_df is not None and not st_df.empty else 8)
        write_dataframe(st_df)

        # 7. المشرفين
        sup_df = results.get("supervisor_table")
        write_section_header("7. أداء وتقييم المشرفين (التغطية والتحصيل ونسب التواصل)", len(sup_df.columns) if sup_df is not None and not sup_df.empty else 15)
        write_dataframe(sup_df)

        # 8. أفضل 10
        t10 = results.get("top_10_collectors")
        write_section_header("8. أفضل 10 محصلين تحصيلاً", len(t10.columns) if t10 is not None and not t10.empty else 17)
        write_dataframe(t10)

        # 9. أقل 10
        b10 = results.get("bottom_10_collectors")
        write_section_header("9. أقل 10 محصلين تحصيلاً", len(b10.columns) if b10 is not None and not b10.empty else 17)
        write_dataframe(b10)

        # 10. جميع المحصلين
        all_c = results.get("all_collectors")
        write_section_header("10. القائمة الكاملة لأداء المحصلين وتصنيف الربع (Quadrants)", len(all_c.columns) if all_c is not None and not all_c.empty else 18)
        write_dataframe(all_c)

        # 11. تحليل المكالمات الفعلية (اختياري)
        if results.get("has_calls"):
            calls_d = results.get("calls_analysis", {})
            sup_calls = calls_d.get("supervisor_calls_df")
            if sup_calls is not None and not sup_calls.empty:
                write_section_header("11. تحليل المكالمات الفعلية — المشرفين", len(sup_calls.columns))
                write_dataframe(sup_calls)
            col_calls = calls_d.get("collector_calls_df")
            if col_calls is not None and not col_calls.empty:
                write_section_header("12. تحليل المكالمات الفعلية — المحصلين", len(col_calls.columns))
                write_dataframe(col_calls)

    output.seek(0)
    return output.getvalue()
