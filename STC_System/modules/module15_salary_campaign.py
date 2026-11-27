# -*- coding: utf-8 -*-
"""
modules/module15_salary_campaign.py
───────────────────────────────────
غرفة متابعة فترة الرواتب (من يوم 26 إلى يوم 5):
- متابعة أداء "أمس" + الأداء "التراكمي" منذ بداية الحملة حتى أمس
- مقارنة تغطية المحصلين للعملاء الإيجابيين + حساب الوتيرة اليومية المطلوبة
- ربط شيت التحصيل (تحصيل أمس + التحصيل التراكمي) برقم المديونية وتاريخ السداد
- ربط شيت المكالمات (اختياري) لكشف المكالمات الفعلية الناجحة ومتوسط المدة
- استخراج شيتات المتابعة العاجلة لكل مشرف (عملاء بدون متابعة + وعود سداد لم تُسدَّد)
- تصدير المحفظة المحللة للمدير + ملفات إكسل منفصلة لكل مشرف (ZIP)
"""

from __future__ import annotations
import io
import os
import re
import zipfile
import logging
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional

import pandas as pd
import polars as pl

from core.utils import normalize_date_series_to_iso, normalize_date_to_iso
from modules.module14_portfolio_analytics import (
    _DEBT_COLS, _ID_COLS, _DEBT_AMT_COLS, _REMAIN_AMT_COLS,
    _SUPERVISOR_COLS, _COLLECTOR_COLS, _PHONE_COLS,
    _FOLLOWUP_DATE_COLS, _STATUS_MAIN_COLS, _STATUS_SUB_COLS,
    _PAY_AMT_COLS, _PAY_DATE_COLS,
    _detect_col, clean_id, to_clean_num
)

_log = logging.getLogger("SALARY_CAMPAIGN")

_CUST_NAME_COLS = ["اسم العميل", "العميل", "customer_name", "Client Name", "Name"]


class SalaryCampaignModule:

    def run(
        self,
        portfolio: Any,
        payments: Optional[Any] = None,
        calls: Optional[Any] = None,
        report_date: Optional[Any] = None,
        campaign_start: Optional[Any] = None,
        campaign_end: Optional[Any] = None,
        collection_target: float = 50000.0,
        daily_coverage_target: float = 50.0,
        daily_collection_target: float = 5000.0,
        daily_calls_target: float = 25.0,
        supervisors: Optional[List[str]] = None,
    ) -> Dict[str, Any]:

        if isinstance(portfolio, pl.DataFrame):
            df_p = portfolio.to_pandas()
        elif isinstance(portfolio, pd.DataFrame):
            df_p = portfolio.copy()
        else:
            df_p = pd.DataFrame(portfolio)

        if len(df_p) == 0:
            raise ValueError("ملف محفظة العملاء الإيجابيين فارغ!")

        # ── 1. كشف الأعمدة الأساسية ──
        col_debt      = _detect_col(df_p, _DEBT_COLS)
        col_id        = _detect_col(df_p, _ID_COLS)
        col_cust      = _detect_col(df_p, _CUST_NAME_COLS)
        col_debt_a    = _detect_col(df_p, _DEBT_AMT_COLS)
        col_rem_a     = _detect_col(df_p, _REMAIN_AMT_COLS)
        col_sup       = _detect_col(df_p, _SUPERVISOR_COLS)
        col_col       = _detect_col(df_p, _COLLECTOR_COLS)
        col_phone     = _detect_col(df_p, _PHONE_COLS)
        col_followup  = _detect_col(df_p, _FOLLOWUP_DATE_COLS)
        col_stat_main = _detect_col(df_p, _STATUS_MAIN_COLS)
        col_stat_sub  = _detect_col(df_p, _STATUS_SUB_COLS)

        if not col_debt:
            raise ValueError("لم يتم العثور على عمود 'رقم المديونية' في ملف المحفظة!")

        # ── 2. التواريخ المرجعية للحملة (أمس + بداية ونهاية الحملة) ──
        yesterday_obj = date.today() - timedelta(days=1)
        rep_dt = normalize_date_to_iso(report_date) if report_date else yesterday_obj.strftime("%Y-%m-%d")
        rep_date_obj = datetime.strptime(rep_dt, "%Y-%m-%d").date()

        if campaign_start:
            camp_s = normalize_date_to_iso(campaign_start)
        else:
            # افتراضياً: يوم 26 من الشهر الحالي أو السابق
            if rep_date_obj.day >= 26:
                camp_s = rep_date_obj.replace(day=26).strftime("%Y-%m-%d")
            else:
                first_cur = rep_date_obj.replace(day=1)
                prev_last = first_cur - timedelta(days=1)
                camp_s = prev_last.replace(day=26).strftime("%Y-%m-%d")

        if campaign_end:
            camp_e = normalize_date_to_iso(campaign_end)
        else:
            s_obj = datetime.strptime(camp_s, "%Y-%m-%d").date()
            if s_obj.month == 12:
                camp_e = date(s_obj.year + 1, 1, 5).strftime("%Y-%m-%d")
            else:
                camp_e = date(s_obj.year, s_obj.month + 1, 5).strftime("%Y-%m-%d")

        end_obj = datetime.strptime(camp_e, "%Y-%m-%d").date()
        days_left = max(1, (end_obj - rep_date_obj).days)

        # ── 3. تنظيف الحقول الأساسية ──
        df_p["_debt_clean"] = df_p[col_debt].apply(clean_id)
        df_p["_cid_clean"]  = df_p[col_id].apply(clean_id) if col_id else df_p["_debt_clean"]
        df_p["_cid_clean"]  = df_p["_cid_clean"].where(df_p["_cid_clean"] != "", df_p["_debt_clean"])
        df_p["_cust_name"]  = df_p[col_cust].fillna("").astype(str).str.strip() if col_cust else df_p["_cid_clean"]
        df_p["_phone"]      = df_p[col_phone].fillna("").astype(str).str.strip().str.replace(r"\.0+$", "", regex=True) if col_phone else ""
        df_p["_debt_amt"]   = to_clean_num(df_p[col_debt_a]) if col_debt_a else 0.0
        df_p["_remain_amt"] = to_clean_num(df_p[col_rem_a]) if col_rem_a else df_p["_debt_amt"]
        df_p["_supervisor"] = df_p[col_sup].fillna("غير محدد").astype(str).str.strip().replace("", "غير محدد") if col_sup else "غير محدد"
        df_p["_collector"]  = df_p[col_col].fillna("غير محدد").astype(str).str.strip().replace("", "غير محدد") if col_col else "غير محدد"

        if supervisors:
            df_p = df_p[df_p["_supervisor"].isin(supervisors)].copy()
            if len(df_p) == 0:
                raise ValueError("لا توجد بيانات للمشرفين المختارين!")

        # توحيد تاريخ المتابعة
        if col_followup and col_followup in df_p.columns:
            norm_fup = normalize_date_series_to_iso(df_p[col_followup].tolist())
            df_p["_fup_iso"] = pd.Series(norm_fup, index=df_p.index)
        else:
            df_p["_fup_iso"] = ""

        is_cov_yesterday = (df_p["_fup_iso"] == rep_dt)
        is_cov_cum = (df_p["_fup_iso"] >= camp_s) & (df_p["_fup_iso"] <= rep_dt) & (df_p["_fup_iso"] != "")
        df_p["_cov_yesterday"] = is_cov_yesterday.astype(int)
        df_p["_cov_cum"] = is_cov_cum.astype(int)

        # حالات التواصل ووعد السداد
        main_c = df_p[col_stat_main].fillna("").astype(str).str.strip() if col_stat_main else pd.Series("", index=df_p.index)
        sub_c  = df_p[col_stat_sub].fillna("").astype(str).str.strip() if col_stat_sub else pd.Series("", index=df_p.index)
        combined_st = main_c + " " + sub_c

        is_no_contact = (
            combined_st.str.contains(r"عدم توصل|عدم تتوصل|الرقم لا يخص|لا يوجد ارقام|غير مستعمل|غير صحيح|لا يخص|لايخص", regex=True)
            | ((main_c == "") & (sub_c == ""))
            | main_c.isin(["---", "-", "null", "none"])
        )
        is_closed_no_ans = (~is_no_contact) & combined_st.str.contains(r"لايرد|لا يرد|مغلق", regex=True)
        is_contacted = (~is_no_contact) & (~is_closed_no_ans)

        df_p["_contact_status"] = "عدم التوصل"
        df_p.loc[is_closed_no_ans, "_contact_status"] = "لا يرد ومغلق"
        df_p.loc[is_contacted, "_contact_status"] = "تم التوصل"

        df_p["_is_promise"] = combined_st.str.contains(r"وعد|سداد|مهل|موعد|جدول", regex=True)
        df_p["_main_status_clean"] = main_c.replace("", "غير محدد")

        # ── 4. ربط شيت التحصيل (تحصيل أمس + التحصيل التراكمي من بداية الحملة) ──
        debt_to_col: Dict[str, str] = {}
        debt_to_sup: Dict[str, str] = {}
        for did, cname, sname in zip(df_p["_debt_clean"], df_p["_collector"], df_p["_supervisor"]):
            if did and did not in debt_to_col:
                debt_to_col[did] = cname
                debt_to_sup[did] = sname

        debt_pay_yest: Dict[str, float] = {}
        debt_pay_cum: Dict[str, float] = {}
        col_pay_yest: Dict[str, float] = {c: 0.0 for c in df_p["_collector"].unique()}
        col_pay_cum: Dict[str, float]  = {c: 0.0 for c in df_p["_collector"].unique()}
        has_payments = False

        if payments is not None:
            df_pay = payments.to_pandas() if isinstance(payments, pl.DataFrame) else (
                payments.copy() if isinstance(payments, pd.DataFrame) else pd.DataFrame(payments)
            )
            if not df_pay.empty:
                pay_debt_col = _detect_col(df_pay, _DEBT_COLS)
                pay_amt_col  = _detect_col(df_pay, _PAY_AMT_COLS)
                pay_date_col = _detect_col(df_pay, _PAY_DATE_COLS)
                if pay_amt_col:
                    has_payments = True
                    df_pay["_amt"] = to_clean_num(df_pay[pay_amt_col])
                    if pay_date_col:
                        df_pay["_p_iso"] = pd.Series(
                            normalize_date_series_to_iso(df_pay[pay_date_col].tolist()),
                            index=df_pay.index
                        )
                    else:
                        df_pay["_p_iso"] = rep_dt

                    p_debts = df_pay[pay_debt_col].apply(clean_id).tolist() if pay_debt_col else [""] * len(df_pay)
                    p_amts  = df_pay["_amt"].tolist()
                    p_dates = df_pay["_p_iso"].tolist()

                    for did, amt, pdt in zip(p_debts, p_amts, p_dates):
                        if amt <= 0.0 or not did or did not in debt_to_col:
                            continue
                        assigned_col = debt_to_col[did]
                        if pdt == rep_dt:
                            debt_pay_yest[did] = debt_pay_yest.get(did, 0.0) + amt
                            col_pay_yest[assigned_col] = col_pay_yest.get(assigned_col, 0.0) + amt
                        if camp_s <= pdt <= rep_dt and pdt != "":
                            debt_pay_cum[did] = debt_pay_cum.get(did, 0.0) + amt
                            col_pay_cum[assigned_col] = col_pay_cum.get(assigned_col, 0.0) + amt

        df_p["_pay_yesterday"] = df_p["_debt_clean"].map(debt_pay_yest).fillna(0.0)
        df_p["_pay_cum"]       = df_p["_debt_clean"].map(debt_pay_cum).fillna(0.0)

        # ── 5. ربط شيت المكالمات (اختياري) ──
        has_calls = False
        col_calls_map: Dict[str, dict] = {}
        debt_calls_map: Dict[str, dict] = {}
        if calls is not None:
            df_calls = calls.to_pandas() if isinstance(calls, pl.DataFrame) else (
                calls.copy() if isinstance(calls, pd.DataFrame) else pd.DataFrame(calls)
            )
            if not df_calls.empty:
                c_debt_col = _detect_col(df_calls, _DEBT_COLS)
                c_att_col  = _detect_col(df_calls, ["إجمالي المحاولات", "اجمالي المحاولات", "المحاولات", "attempts"])
                c_suc_col  = _detect_col(df_calls, [
                    "إجمالي المكالمات الناجحة أكبر من 30 ثانية", "المكالمات الناجحه اكبر من 30 ث",
                    "المكالمات الناجحة > 30", "المكالمات الناجحة", "successful_calls"
                ])
                c_fail_col = _detect_col(df_calls, ["failed", "Failed", "فاشلة", "غير ناجحة"])
                c_dur_col  = _detect_col(df_calls, ["مدة المكالمة بالثواني", "مده ثوانى المكالمه", "مدة المكالمة", "duration", "seconds"])

                if c_debt_col:
                    has_calls = True
                    n_c = len(df_calls)
                    att_s  = to_clean_num(df_calls[c_att_col])  if c_att_col  else pd.Series([1.0] * n_c)
                    suc_s  = to_clean_num(df_calls[c_suc_col])  if c_suc_col  else pd.Series([0.0] * n_c)
                    fail_s = to_clean_num(df_calls[c_fail_col]) if c_fail_col else (att_s - suc_s).clip(lower=0.0)
                    dur_s  = to_clean_num(df_calls[c_dur_col])  if c_dur_col  else pd.Series([0.0] * n_c)

                    for did_raw, a_v, s_v, f_v, d_v in zip(
                        df_calls[c_debt_col].apply(clean_id), att_s, suc_s, fail_s, dur_s
                    ):
                        if not did_raw or did_raw not in debt_to_col:
                            continue
                        cname = debt_to_col[did_raw]
                        dm = debt_calls_map.setdefault(did_raw, {"att": 0, "suc": 0, "fail": 0, "durs": []})
                        dm["att"]  += int(round(a_v))
                        dm["suc"]  += int(round(s_v))
                        dm["fail"] += int(round(f_v))
                        if c_dur_col and d_v > 0:
                            dm["durs"].append(float(d_v))

                        cm = col_calls_map.setdefault(cname, {"att": 0, "suc": 0, "fail": 0, "durs": []})
                        cm["att"]  += int(round(a_v))
                        cm["suc"]  += int(round(s_v))
                        cm["fail"] += int(round(f_v))
                        if c_dur_col and d_v > 0:
                            cm["durs"].append(float(d_v))

        # ── 6. بناء جداول الأداء للمحصلين والمشرفين ──
        tgt_val = float(collection_target if collection_target and collection_target > 0 else 50000.0)
        d_cov_tgt = float(daily_coverage_target if daily_coverage_target and daily_coverage_target > 0 else 50.0)
        d_pay_tgt = float(daily_collection_target if daily_collection_target and daily_collection_target > 0 else 5000.0)
        d_cal_tgt = float(daily_calls_target if daily_calls_target and daily_calls_target > 0 else 25.0)
        col_rows = []

        for (sup_name, col_name), grp in df_p.groupby(["_supervisor", "_collector"]):
            tot_acc  = int(grp["_cid_clean"].nunique())
            cov_y    = int(grp.loc[grp["_cov_yesterday"] == 1, "_cid_clean"].nunique())
            cov_c    = int(grp.loc[grp["_cov_cum"] == 1, "_cid_clean"].nunique())
            rem_cov  = max(0, tot_acc - cov_c)
            cov_pct  = round(cov_c / tot_acc * 100, 1) if tot_acc > 0 else 0.0
            cov_y_pct = round(cov_y / d_cov_tgt * 100, 1) if d_cov_tgt > 0 else 0.0
            req_daily_cov = round(rem_cov / days_left, 1)

            pay_y    = round(col_pay_yest.get(col_name, 0.0), 2)
            pay_y_pct = round(pay_y / d_pay_tgt * 100, 1) if d_pay_tgt > 0 else 0.0
            pay_c    = round(col_pay_cum.get(col_name, 0.0), 2)
            ach_pct  = round(pay_c / tgt_val * 100, 1) if tgt_val > 0 else 0.0
            rem_tgt  = max(0.0, tgt_val - pay_c)
            req_daily_pay = round(rem_tgt / days_left, 2)

            cm = col_calls_map.get(col_name, {"att": 0, "suc": 0, "fail": 0, "durs": []})
            calls_pct = round(cm["suc"] / d_cal_tgt * 100, 1) if (has_calls and d_cal_tgt > 0) else 0.0

            # تقييم التزام المحصل بناءً على المستهدفات اليومية المحددة من المستخدم
            calls_ok = (not has_calls) or (cm["suc"] >= d_cal_tgt * 0.8)
            if cov_y >= d_cov_tgt * 0.8 and calls_ok and (pay_y >= d_pay_tgt * 0.5 or pay_c > 0):
                eval_lbl = "🟢 على المسار"
            elif cov_y >= d_cov_tgt * 0.4 or cov_c > 0:
                eval_lbl = "🟡 يحتاج متابعة"
            else:
                eval_lbl = "🔴 تقصير / تدخل فوري"

            row_d = {
                "المشرف": sup_name,
                "المحصل": col_name,
                "إجمالي الإيجابيين": tot_acc,
                "غطى أمس": cov_y,
                "مستهدف التغطية اليومي": int(d_cov_tgt),
                "إنجاز التغطية اليومي %": cov_y_pct,
                "التغطية التراكمية": cov_c,
                "المتبقي بدون تغطية": rem_cov,
                "نسبة التغطية الكلية %": cov_pct,
                "المطلوب يومياً (عميل)": req_daily_cov,
                "تحصيل أمس (﷼)": pay_y,
                "مستهدف التحصيل اليومي (﷼)": round(d_pay_tgt, 2),
                "إنجاز التحصيل اليومي %": pay_y_pct,
                "التحصيل التراكمي (﷼)": pay_c,
                "مستهدف الحملة (﷼)": round(tgt_val, 2),
                "إنجاز الحملة %": ach_pct,
                "المطلوب يومياً (﷼)": req_daily_pay,
                "التقييم": eval_lbl,
            }
            if has_calls:
                row_d["محاولات الاتصال"] = cm["att"]
                row_d["مكالمات ناجحة (>30 ث)"] = cm["suc"]
                row_d["مستهدف المكالمات"] = int(d_cal_tgt)
                row_d["إنجاز المكالمات %"] = calls_pct
                row_d["مكالمات فاشلة"] = cm["fail"]
                row_d["متوسط المدة (ث)"] = round(sum(cm["durs"]) / len(cm["durs"]), 1) if cm["durs"] else 0.0
            col_rows.append(row_d)

        collector_df = pd.DataFrame(col_rows)
        if not collector_df.empty:
            collector_df = collector_df.sort_values(["التحصيل التراكمي (﷼)", "نسبة التغطية الكلية %"], ascending=[False, False]).reset_index(drop=True)

        # جدول المشرفين
        sup_rows = []
        if not collector_df.empty:
            for sup_name, sgrp in collector_df.groupby("المشرف"):
                s_tot = int(sgrp["إجمالي الإيجابيين"].sum())
                s_cy  = int(sgrp["غطى أمس"].sum())
                s_cc  = int(sgrp["التغطية التراكمية"].sum())
                s_rem = int(sgrp["المتبقي بدون تغطية"].sum())
                s_py  = round(float(sgrp["تحصيل أمس (﷼)"].sum()), 2)
                s_pc  = round(float(sgrp["التحصيل التراكمي (﷼)"].sum()), 2)
                s_tgt = round(float(sgrp["مستهدف الحملة (﷼)"].sum()), 2)
                s_row = {
                    "المشرف": sup_name,
                    "عدد المحصلين": len(sgrp),
                    "إجمالي الإيجابيين": s_tot,
                    "غطى أمس": s_cy,
                    "التغطية التراكمية": s_cc,
                    "المتبقي بدون تغطية": s_rem,
                    "نسبة التغطية %": round(s_cc / s_tot * 100, 1) if s_tot > 0 else 0.0,
                    "تحصيل أمس (﷼)": s_py,
                    "التحصيل التراكمي (﷼)": s_pc,
                    "المستهدف المالي (﷼)": s_tgt,
                    "نسبة الإنجاز %": round(s_pc / s_tgt * 100, 1) if s_tgt > 0 else 0.0,
                }
                if has_calls:
                    s_row["محاولات الاتصال"] = int(sgrp["محاولات الاتصال"].sum())
                    s_row["مكالمات ناجحة (>30 ث)"] = int(sgrp["مكالمات ناجحة (>30 ث)"].sum())
                sup_rows.append(s_row)

        supervisor_df = pd.DataFrame(sup_rows)
        if not supervisor_df.empty:
            supervisor_df = supervisor_df.sort_values("التحصيل التراكمي (﷼)", ascending=False).reset_index(drop=True)

        # ── 7. شيتات توجيه المشرفين (Action Sheets) ──
        def _action_label(r_row):
            if r_row["_pay_cum"] > 0:
                return "✅ تم السداد"
            if r_row["_is_promise"]:
                return "🔥 أولوية قصوى — وعد سداد لم يُسدَّد"
            if r_row["_cov_cum"] == 0:
                return "📞 اتصال فوري — عميل إيجابي لم يُغطَّ"
            if r_row["_contact_status"] == "لا يرد ومغلق":
                return "🔁 إعادة محاولة الاتصال"
            return "📋 متابعة مستمرة"

        df_p["_action_req"] = df_p.apply(_action_label, axis=1)

        action_cols_map = {
            "_supervisor": "المشرف",
            "_collector": "المحصل",
            "_debt_clean": "رقم المديونية",
            "_cid_clean": "رقم الهوية",
            "_cust_name": "اسم العميل",
            "_phone": "الرقم الرئيسي",
            "_debt_amt": "مبلغ المديونية",
            "_remain_amt": "متبقي السداد",
            "_main_status_clean": "الحالة الرئيسية",
            "_fup_iso": "آخر تاريخ متابعة",
            "_action_req": "الإجراء المطلوب",
        }

        uncovered_mask = (df_p["_cov_cum"] == 0) & (df_p["_pay_cum"] == 0)
        uncovered_df = df_p[uncovered_mask][list(action_cols_map.keys())].rename(columns=action_cols_map)
        uncovered_df = uncovered_df.sort_values("مبلغ المديونية", ascending=False).reset_index(drop=True)

        promises_mask = (df_p["_is_promise"]) & (df_p["_pay_cum"] == 0)
        unpaid_promises_df = df_p[promises_mask][list(action_cols_map.keys())].rename(columns=action_cols_map)
        unpaid_promises_df = unpaid_promises_df.sort_values("مبلغ المديونية", ascending=False).reset_index(drop=True)

        # ── 8. المحفظة المحللة الكاملة للمدير ──
        enriched_portfolio = df_p.copy()
        enriched_portfolio["تاريخ المتابعة (موحد)"] = df_p["_fup_iso"]
        enriched_portfolio["حالة التغطية بالحملة"] = df_p.apply(
            lambda x: "تمت المتابعة أمس" if x["_cov_yesterday"] == 1
            else ("تمت المتابعة سابقاً بالحملة" if x["_cov_cum"] == 1 else "لم تتم المتابعة بالحملة"),
            axis=1
        )
        enriched_portfolio["حالة التوصل (تحليل)"] = df_p["_contact_status"]
        enriched_portfolio["تحصيل أمس (﷼)"] = df_p["_pay_yesterday"].round(2)
        enriched_portfolio["تحصيل الحملة التراكمي (﷼)"] = df_p["_pay_cum"].round(2)
        enriched_portfolio["حالة السداد"] = df_p["_pay_cum"].apply(lambda v: "مسدد" if v > 0 else "غير مسدد")
        enriched_portfolio["الإجراء المطلوب للمحصل"] = df_p["_action_req"]

        if has_calls:
            _d_keys = df_p["_debt_clean"].tolist()
            enriched_portfolio["محاولات الاتصال"] = [debt_calls_map[k]["att"] if k in debt_calls_map else 0 for k in _d_keys]
            enriched_portfolio["مكالمات ناجحة (>30 ث)"] = [debt_calls_map[k]["suc"] if k in debt_calls_map else 0 for k in _d_keys]
            enriched_portfolio["مكالمات فاشلة"] = [debt_calls_map[k]["fail"] if k in debt_calls_map else 0 for k in _d_keys]
            enriched_portfolio["متوسط المدة (ث)"] = [
                round(sum(debt_calls_map[k]["durs"]) / len(debt_calls_map[k]["durs"]), 1)
                if (k in debt_calls_map and debt_calls_map[k]["durs"]) else 0.0
                for k in _d_keys
            ]

        drop_internal = [
            "_debt_clean", "_cid_clean", "_cust_name", "_phone", "_debt_amt", "_remain_amt",
            "_supervisor", "_collector", "_fup_iso", "_cov_yesterday", "_cov_cum",
            "_contact_status", "_is_promise", "_main_status_clean",
            "_pay_yesterday", "_pay_cum", "_action_req"
        ]
        enriched_portfolio = enriched_portfolio.drop(columns=drop_internal, errors="ignore")

        # ── 9. ملخص الإحصائيات (KPIs - بالهويات الفريدة) ──
        tot_p   = int(df_p["_cid_clean"].nunique())
        tot_cy  = int(df_p.loc[df_p["_cov_yesterday"] == 1, "_cid_clean"].nunique())
        tot_cc  = int(df_p.loc[df_p["_cov_cum"] == 1, "_cid_clean"].nunique())
        tot_rem = max(0, tot_p - tot_cc)
        tot_py  = round(sum(col_pay_yest.values()), 2)
        tot_pc  = round(sum(col_pay_cum.values()), 2)
        tot_tgt = round(float(collector_df["مستهدف الحملة (﷼)"].sum()), 2) if not collector_df.empty else tgt_val

        stats = {
            "تاريخ التقرير (أمس)": rep_dt,
            "فترة الحملة": f"{camp_s} إلى {camp_e} (متبقي {days_left} يوم)",
            "إجمالي العملاء الإيجابيين": f"{tot_p:,}",
            "تغطية أمس": f"{tot_cy:,}",
            "التغطية التراكمية": f"{tot_cc:,} ({round(tot_cc/tot_p*100,1) if tot_p else 0}%)",
            "المتبقي بدون تغطية": f"{tot_rem:,}",
            "تحصيل أمس": f"{tot_py:,.2f} ﷼",
            "التحصيل التراكمي": f"{tot_pc:,.2f} ﷼ ({round(tot_pc/tot_tgt*100,1) if tot_tgt else 0}%)",
            "وعود سداد لم تُسدَّد": f"{len(unpaid_promises_df):,}",
        }

        return {
            "report_date": rep_dt,
            "campaign_start": camp_s,
            "campaign_end": camp_e,
            "days_left": days_left,
            "stats": stats,
            "supervisor_table": supervisor_df,
            "collector_table": collector_df,
            "uncovered_df": uncovered_df,
            "unpaid_promises_df": unpaid_promises_df,
            "enriched_portfolio": enriched_portfolio,
            "has_payments": has_payments,
            "has_calls": has_calls,
        }


def _write_styled_df_sheet(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str):
    workbook = writer.book
    safe_name = sheet_name[:31].replace(":", "-").replace("/", "-").replace("\\", "-").replace("?", "").replace("*", "").replace("[", "").replace("]", "")
    ws = workbook.add_worksheet(safe_name)
    ws.right_to_left()

    hdr_fmt = workbook.add_format({
        "bold": True, "bg_color": "#262626", "font_color": "#FFFFFF",
        "border": 1, "align": "center", "valign": "vcenter", "font_size": 11
    })
    cell_fmt = workbook.add_format({
        "border": 1, "align": "center", "valign": "vcenter", "font_size": 10
    })

    if df is None or df.empty:
        ws.write(0, 0, "لا توجد بيانات", cell_fmt)
        return

    for c_idx, col in enumerate(df.columns):
        ws.write(0, c_idx, str(col), hdr_fmt)
        sample_lens = [len(str(x)) for x in df[col].head(100).values]
        max_len = max([len(str(col))] + sample_lens) if sample_lens else len(str(col))
        ws.set_column(c_idx, c_idx, min(max(max_len + 3, 12), 40))

    for r_idx, (_, row) in enumerate(df.iterrows(), start=1):
        for c_idx, val in enumerate(row.values):
            safe_val = "-" if (val is None or (isinstance(val, float) and pd.isna(val))) else val
            ws.write(r_idx, c_idx, safe_val, cell_fmt)


def export_salary_report_excel(res: Dict[str, Any]) -> bytes:
    """تقرير الإدارة الشامل متعدد الشيتات"""
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter", engine_kwargs={"options": {"nan_inf_to_errors": True}}) as writer:
        _write_styled_df_sheet(writer, res.get("supervisor_table"), "ملخص المشرفين")
        _write_styled_df_sheet(writer, res.get("collector_table"), "أداء المحصلين (أمس+تراكمي)")
        _write_styled_df_sheet(writer, res.get("unpaid_promises_df"), "وعود سداد لم تسدد")
        _write_styled_df_sheet(writer, res.get("uncovered_df"), "عملاء بدون تغطية")
    out.seek(0)
    return out.getvalue()


def export_enriched_portfolio_excel(res: Dict[str, Any]) -> bytes:
    """ملف المحفظة الإيجابية المحللة للمدير"""
    out = io.BytesIO()
    df_ep = res.get("enriched_portfolio", pd.DataFrame())
    with pd.ExcelWriter(out, engine="xlsxwriter", engine_kwargs={"options": {"nan_inf_to_errors": True}}) as writer:
        _write_styled_df_sheet(writer, df_ep, "المحفظة المحللة")
    out.seek(0)
    return out.getvalue()


def export_supervisor_action_zip(res: Dict[str, Any]) -> bytes:
    """توليد ملف ZIP يحتوي على شيت إكسل مستقل لكل مشرف لإرساله له مباشرة"""
    col_df = res.get("collector_table", pd.DataFrame())
    uncov_df = res.get("uncovered_df", pd.DataFrame())
    prom_df = res.get("unpaid_promises_df", pd.DataFrame())

    sup_names = sorted(col_df["المشرف"].unique().tolist()) if not col_df.empty else []
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for sup in sup_names:
            s_cols  = col_df[col_df["المشرف"] == sup] if not col_df.empty else pd.DataFrame()
            s_uncov = uncov_df[uncov_df["المشرف"] == sup] if not uncov_df.empty else pd.DataFrame()
            s_prom  = prom_df[prom_df["المشرف"] == sup] if not prom_df.empty else pd.DataFrame()

            sup_out = io.BytesIO()
            with pd.ExcelWriter(sup_out, engine="xlsxwriter", engine_kwargs={"options": {"nan_inf_to_errors": True}}) as sw:
                _write_styled_df_sheet(sw, s_cols, "ملخص أداء فريقك")
                _write_styled_df_sheet(sw, s_prom, "وعود سداد لم تسدد")
                _write_styled_df_sheet(sw, s_uncov, "عملاء بدون متابعة")
            safe_fname = re.sub(r'[\\/*?:"<>|]', "_", str(sup))
            zf.writestr(f"توجيهات_المشرف_{safe_fname}.xlsx", sup_out.getvalue())

    zip_buf.seek(0)
    return zip_buf.getvalue()
