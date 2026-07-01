from __future__ import annotations

import json
import re
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd


def _to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("%", "").replace(" ", "")
    if not text or text.lower() in {"nan", "none", "n/a"}:
        return None
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _infer_year_from_text(*values: Any) -> Optional[int]:
    """Infer a four-digit year from report metadata.

    The weekly reports have appeared in several formats, for example:
    "03 February 2026", "June 16, 2026", "13-Jun-26" and W15-26.
    The function deliberately returns a four-digit year so the app can
    display and sort history from 2024, 2025, 2026, etc. without mixing
    same-numbered weeks across different years.
    """
    text = " ".join(str(v or "") for v in values)
    m = re.search(r"\b(20\d{2})\b", text)
    if m:
        return int(m.group(1))
    # Date-like forms: 13-Jun-26, 13/06/26, 13.06.26
    m = re.search(r"\b\d{1,2}[\-/\.][A-Za-z]{3,9}[\-/\.](\d{2})\b", text)
    if not m:
        m = re.search(r"\b\d{1,2}[\-/\.]\d{1,2}[\-/\.](\d{2})\b", text)
    if not m:
        m = re.search(r"\bW\d{1,2}[-\s]?(\d{2})\b", text, flags=re.I)
    if m:
        yy = int(m.group(1))
        return 2000 + yy if yy < 80 else 1900 + yy
    return None


def _row_year(row: Dict[str, Any], fallback: Optional[int] = None) -> Optional[int]:
    year = row.get("year")
    if isinstance(year, int):
        return year
    if isinstance(year, str) and year.isdigit() and len(year) == 4:
        return int(year)
    inferred = _infer_year_from_text(
        row.get("issue_date"),
        row.get("cutoff"),
        row.get("file_name"),
        row.get("period"),
        row.get("week_label"),
    )
    return inferred if inferred is not None else fallback


def enrich_years(bundle: Dict[str, Any]) -> Dict[str, Any]:
    """Add a `year` field to rows when it can be derived.

    Rows such as watchlist items do not always contain issue dates. They do,
    however, carry the source report file name, so the year is propagated from
    HSE / schedule rows belonging to the same file.
    """
    bundle = deepcopy(bundle)
    year_by_file: Dict[str, int] = {}

    for section in ["hse", "schedule", "risks_summary", "source_files"]:
        for row in bundle.get(section, []) or []:
            year = _row_year(row)
            if year and row.get("file_name"):
                year_by_file[str(row.get("file_name"))] = year

    # Sometimes only the HSE row has an issue date and the schedule/watchlist
    # rows only have week + file. Propagate year from file first, then infer.
    for section, rows in bundle.items():
        if not isinstance(rows, list):
            continue
        for row in rows:
            if not isinstance(row, dict):
                continue
            fallback = year_by_file.get(str(row.get("file_name", "")))
            year = _row_year(row, fallback=fallback)
            if year:
                row["year"] = int(year)
                if row.get("file_name"):
                    year_by_file[str(row.get("file_name"))] = int(year)
    return bundle


def period_key(row: Dict[str, Any]) -> Tuple[int, int, str]:
    week = row.get("week")
    try:
        week_int = int(week)
    except (TypeError, ValueError):
        week_int = -1
    year = _row_year(row) or -1
    return (int(year), week_int, str(row.get("file_name", "")))


def _week_sort(row: Dict[str, Any]) -> Tuple[int, int, str]:
    return period_key(row)


def period_label(row: Dict[str, Any] | pd.Series | None) -> str:
    if row is None:
        return ""
    if isinstance(row, pd.Series):
        data = row.to_dict()
    else:
        data = row
    week = data.get("week")
    year = data.get("year") or _row_year(data)
    try:
        week_label = f"W{int(week):02d}"
    except (TypeError, ValueError):
        week_label = str(week or "")
    return f"{int(year)} {week_label}" if year else week_label


def add_period_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    if "year" not in df.columns:
        df["year"] = df.apply(lambda r: _row_year(r.to_dict()), axis=1)
    df["period_label"] = df.apply(period_label, axis=1)
    df["period_sort"] = df.apply(lambda r: ((r.get("year") if pd.notna(r.get("year")) else -1) * 100 + (r.get("week") if pd.notna(r.get("week")) else -1)), axis=1)
    return df


def load_json(path: Path | str, default: Any = None) -> Any:
    path = Path(path)
    if not path.exists():
        return deepcopy(default)
    return json.loads(path.read_text(encoding="utf-8"))


def load_bundle(path: Path | str) -> Dict[str, Any]:
    bundle = load_json(path, default={
        "hse": [],
        "schedule": [],
        "waypoints": [],
        "risks_summary": [],
        "risks": [],
        "watchlist": [],
        "engineering_scope_history": [],
        "source_files": [],
    })
    return enrich_years(bundle)


def save_bundle(bundle: Dict[str, Any], path: Path | str) -> None:
    Path(path).write_text(json.dumps(enrich_years(bundle), ensure_ascii=False, indent=2), encoding="utf-8")


def load_profiles(path: Path | str) -> Dict[str, Any]:
    return load_json(path, default={})


def hse_df(bundle: Dict[str, Any]) -> pd.DataFrame:
    df = pd.DataFrame(enrich_years(bundle).get("hse", []))
    if df.empty:
        return df
    df = df.sort_values(["year", "week"], na_position="first")
    numeric_cols = [
        "site_walks", "observations", "bbs", "target_inspections", "near_miss",
        "permits_to_work", "rams_approved", "rewards", "sending_offs",
        "ltifr", "trifr", "positive_pct", "negative_pct", "lti_free_days",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    leading_cols = [c for c in ["site_walks", "observations", "bbs", "target_inspections"] if c in df.columns]
    df["leading_index"] = df[leading_cols].sum(axis=1, skipna=True) if leading_cols else 0
    return add_period_columns(df)


def schedule_df(bundle: Dict[str, Any]) -> pd.DataFrame:
    df = pd.DataFrame(enrich_years(bundle).get("schedule", []))
    if df.empty:
        return df
    df = df.sort_values(["year", "week"], na_position="first")
    numeric_cols = [
        "actual_pct", "baseline_pct", "forecast_pct", "deviation_pct", "overall_week_actual",
        "engineering_cumm_actual", "engineering_cumm_plan", "engineering_cumm_forecast", "engineering_deviation_pct",
        "procurement_cumm_actual", "procurement_cumm_plan", "procurement_cumm_forecast", "procurement_deviation_pct",
        "construction_cumm_actual", "construction_cumm_plan", "construction_cumm_forecast", "construction_deviation_pct",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return add_period_columns(df)


def risks_summary_df(bundle: Dict[str, Any]) -> pd.DataFrame:
    df = pd.DataFrame(enrich_years(bundle).get("risks_summary", []))
    if df.empty:
        return df
    df = df.sort_values(["year", "week"], na_position="first")
    return add_period_columns(df)


def latest_row(rows: Iterable[Dict[str, Any]]) -> Dict[str, Any]:
    rows = list(rows)
    if not rows:
        return {}
    return sorted(rows, key=period_key)[-1]


def latest_period(bundle: Dict[str, Any]) -> Tuple[Optional[int], Optional[int]]:
    rows: List[Dict[str, Any]] = []
    enriched = enrich_years(bundle)
    for section in ["hse", "schedule", "risks_summary", "watchlist"]:
        rows.extend([r for r in enriched.get(section, []) if isinstance(r, dict) and r.get("week") is not None])
    if not rows:
        return (None, None)
    row = sorted(rows, key=period_key)[-1]
    return (_row_year(row), int(row["week"]) if row.get("week") is not None else None)


def latest_week(bundle: Dict[str, Any]) -> Optional[int]:
    return latest_period(bundle)[1]


def latest_period_label(bundle: Dict[str, Any]) -> str:
    year, week = latest_period(bundle)
    if week is None:
        return "No reports loaded"
    return f"{year} W{week:02d}" if year else f"W{week:02d}"


def source_count(bundle: Dict[str, Any]) -> int:
    return len(bundle.get("source_files", []))


def metric_delta(series: pd.Series) -> Optional[float]:
    clean = series.dropna()
    if len(clean) < 2:
        return None
    return float(clean.iloc[-1] - clean.iloc[-2])


_BAD_ISSUES = [
    r"^$", r"^progress$", r"^progress status$", r"^status:?$", r"^red activity$", r"^mitigating action$",
    r"^planned date$", r"^forecast date$", r"^key waypoints", r"^discipline$", r"^behind / major concern$",
    r"^on track / no issue", r"^risk / concern$", r"^ach\.?$", r"^act$", r"^tba$", r"^week\s+[0-9,.% ]+$",
    r"^w\d{2}-\d{2}", r"^\d{2}/\d{2}-\d{2}/\d{2}$", r"^1\. achievement",
    r"^cumm\.?$", r"^vprogress status$", r"^buildings & others progress$",
]


def is_meaningful_issue(text: str) -> bool:
    text = (text or "").strip()
    if len(text) < 8:
        return False
    if len(re.sub(r"[^A-Za-z]", "", text)) < 5:
        return False
    lower = text.lower().strip()
    return not any(re.search(pattern, lower) for pattern in _BAD_ISSUES)


def scope_row_filter(row: Dict[str, Any], scope_id: str, profile: Dict[str, Any]) -> bool:
    area = row.get("area")
    text = f"{row.get('issue', '')} {row.get('action', '')}".lower()
    if scope_id == "":
        ugp_terms = [
            "", "", "", "", "underground networks", "",
            "flange management", "hdpe", "golden joint", "test pack bottleneck", "ug cable line", "avenue 9",
            "road 6.1", "road 6.2", "road 5.1", "road 5.2", "road 7 ave 9.5", "route 1",
        ]
        if not any(term in text for term in ugp_terms):
            return False
        return area in {"Roads", "Rail", "Ponds", "NPB/Buildings", "General", "BOP", "PDOP", ""}
    if scope_id in {"rail", "ponds", "roads"}:
        return area == profile.get("source_area")
    if area != "NPB/Buildings":
        return False
    text = f"{row.get('issue', '')} {row.get('action', '')}".lower()
    if scope_id == "npb100":
        return any(k in text for k in ["npb 10", "npb102", "npb 102", "npb103", "npb 103", "employees changing", "safety & security"])
    if scope_id == "npb200":
        return any(k in text for k in ["npb 20", "npb202", "npb 202", "npb203", "npb 203", "npb205", "npb 205", "npb206", "npb 206", "npb207", "npb 207", "warehouse", "refractory", "by-products", "slag", "crane maintenance", "chemical storage"])
    return False


def scope_watchlist(bundle: Dict[str, Any], scope_id: str, profile: Dict[str, Any], week: Optional[int] = None, year: Optional[int] = None) -> List[Dict[str, Any]]:
    rows = [r for r in enrich_years(bundle).get("watchlist", []) if scope_row_filter(r, scope_id, profile)]
    if week is not None:
        rows = [r for r in rows if r.get("week") == week]
    if year is not None:
        rows = [r for r in rows if _row_year(r) == year]
    out: List[Dict[str, Any]] = []
    seen = set()
    for row in sorted(rows, key=lambda r: (period_key(r), 0 if r.get("severity") == "High" else 1, str(r.get("category", ""))), reverse=True):
        issue = str(row.get("issue", "")).strip()
        action = str(row.get("action", "")).strip()
        if not is_meaningful_issue(issue):
            continue
        key = (issue[:120].lower(), action[:80].lower(), row.get("week"), _row_year(row))
        if key in seen:
            continue
        seen.add(key)
        out.append(row)
    return out


def parse_progress_from_issue(issue: str) -> Dict[str, Optional[float]]:
    out = {"actual": None, "forecast": None, "deviation": None}
    issue = issue or ""
    for key, pat in [
        ("actual", r"actual\s+([0-9]+[,.]?[0-9]*)%"),
        ("forecast", r"forecast\s+([0-9]+[,.]?[0-9]*)%"),
        ("deviation", r"deviation\s+([-]?[0-9]+[,.]?[0-9]*)%"),
    ]:
        m = re.search(pat, issue, re.I)
        if m:
            out[key] = _to_float(m.group(1))
    return out


def scope_progress_series(bundle: Dict[str, Any], scope_id: str, profile: Dict[str, Any]) -> pd.DataFrame:
    rows = []
    enriched = enrich_years(bundle)
    for row in enriched.get("watchlist", []):
        if row.get("category") != "Progress":
            continue
        if not scope_row_filter(row, scope_id, profile):
            continue
        progress = parse_progress_from_issue(row.get("issue", ""))
        if progress["actual"] is not None:
            rows.append({
                "year": _row_year(row),
                "week": row.get("week"),
                "actual": progress["actual"],
                "forecast": progress["forecast"],
                "deviation": progress["deviation"],
            })
    if not rows:
        # profile-only fallback
        p = profile.get("progress", {})
        if p.get("actual") is not None:
            year, week = latest_period(enriched)
            rows = [{"year": profile.get("latest_year") or year, "week": profile.get("latest_week") or week, "actual": p.get("actual"), "forecast": p.get("forecast"), "deviation": p.get("deviation")}]
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df = df.sort_values(["year", "week"], na_position="first")
    for col in ["actual", "forecast", "deviation"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return add_period_columns(df)


def _extract_first_percent(action: str) -> Optional[float]:
    m = re.search(r"([0-9]+[,.][0-9]+)%", action or "")
    return _to_float(m.group(1)) if m else None


def dynamic_profile_from_bundle(bundle: Dict[str, Any], profile: Dict[str, Any], scope_id: str) -> Dict[str, Any]:
    profile = deepcopy(profile)
    enriched = enrich_years(bundle)
    progress_df = scope_progress_series(enriched, scope_id, profile)
    if not progress_df.empty:
        row = progress_df.iloc[-1]
        progress = profile.setdefault("progress", {})
        if pd.notna(row.get("actual")):
            progress["actual"] = float(row["actual"])
        if pd.notna(row.get("forecast")):
            progress["forecast"] = float(row["forecast"])
        if pd.notna(row.get("deviation")):
            progress["deviation"] = float(row["deviation"])
        if pd.notna(row.get("week")):
            profile["latest_week"] = int(row["week"])
        if pd.notna(row.get("year")):
            profile["latest_year"] = int(row["year"])
    # NPB 200 progress is currently published as a Buildings & Others table row rather than a clean Progress item.
    if scope_id == "npb200":
        rows = [r for r in enriched.get("watchlist", []) if r.get("area") == "NPB/Buildings" and r.get("issue") == "Cumm."]
        if rows:
            row = sorted(rows, key=_week_sort)[-1]
            val = _extract_first_percent(row.get("action", ""))
            if val is not None:
                profile.setdefault("progress", {})["actual"] = val
                profile.setdefault("progress", {})["forecast"] = val
                profile["latest_week"] = int(row.get("week") or profile.get("latest_week") or 0)
                y = _row_year(row)
                if y:
                    profile["latest_year"] = y
    return profile


def get_scope_summary(bundle: Dict[str, Any], profiles: Dict[str, Any], scope_id: str) -> Dict[str, Any]:
    enriched = enrich_years(bundle)
    base = profiles[scope_id]
    profile = dynamic_profile_from_bundle(enriched, base, scope_id)
    year, week = latest_period(enriched)
    current_rows = scope_watchlist(enriched, scope_id, profile, week=week, year=year)
    all_rows = scope_watchlist(enriched, scope_id, profile)
    high_rows = [r for r in current_rows if r.get("severity") == "High"]
    lookahead = [r for r in current_rows if r.get("category") == "Look-ahead"]
    blockers = profile.get("blockers", [])[:]
    # Enrich blockers with high issues, but keep clean and short.
    for row in high_rows:
        issue = row.get("issue", "")
        if issue and all(issue[:70] not in b for b in blockers):
            blockers.append(issue)
        if len(blockers) >= 6:
            break
    return {
        "profile": profile,
        "current_rows": current_rows,
        "all_rows": all_rows,
        "high_count": len(high_rows),
        "lookahead": lookahead,
        "blockers": blockers[:6],
        "quality": profile.get("quality", []),
        "progress_series": scope_progress_series(enriched, scope_id, profile),
    }


def merge_bundles(base: Dict[str, Any], incoming: Dict[str, Any]) -> Dict[str, Any]:
    """Merge a parsed upload into the current app data by stable keys."""
    merged = enrich_years(deepcopy(base))
    incoming = enrich_years(incoming)
    for section in ["hse", "schedule", "risks_summary"]:
        current = {(r.get("year"), r.get("week"), r.get("file_name")): r for r in merged.get(section, [])}
        for row in incoming.get(section, []):
            key = (row.get("year"), row.get("week"), row.get("file_name"))
            current[key] = row
        merged[section] = sorted(current.values(), key=_week_sort)
    for section in ["waypoints"]:
        current = {(r.get("year"), r.get("week"), r.get("file_name"), r.get("area"), r.get("waypoint")): r for r in merged.get(section, [])}
        for row in incoming.get(section, []):
            key = (row.get("year"), row.get("week"), row.get("file_name"), row.get("area"), row.get("waypoint"))
            current[key] = row
        merged[section] = sorted(current.values(), key=lambda r: (period_key(r), str(r.get("area", "")), str(r.get("waypoint", ""))))
    for section in ["risks"]:
        current = {(r.get("year"), r.get("week"), r.get("file_name"), r.get("risk_id")): r for r in merged.get(section, [])}
        for row in incoming.get(section, []):
            key = (row.get("year"), row.get("week"), row.get("file_name"), row.get("risk_id"))
            current[key] = row
        merged[section] = sorted(current.values(), key=lambda r: (period_key(r), -(r.get("current_score") or 0)))
    for section in ["watchlist"]:
        current = {(r.get("year"), r.get("week"), r.get("file_name"), r.get("area"), r.get("category"), str(r.get("issue", ""))[:160]): r for r in merged.get(section, [])}
        for row in incoming.get(section, []):
            key = (row.get("year"), row.get("week"), row.get("file_name"), row.get("area"), row.get("category"), str(row.get("issue", ""))[:160])
            current[key] = row
        merged[section] = sorted(current.values(), key=lambda r: (period_key(r), str(r.get("area", "")), str(r.get("category", ""))))

    for section in ["engineering_scope_history"]:
        current = {(r.get("year"), r.get("week"), r.get("file_name"), r.get("scope_id"), r.get("area")): r for r in merged.get(section, [])}
        for row in incoming.get(section, []):
            key = (row.get("year"), row.get("week"), row.get("file_name"), row.get("scope_id"), row.get("area"))
            current[key] = row
        merged[section] = sorted(current.values(), key=lambda r: (period_key(r), str(r.get("scope_id", "")), str(r.get("area", ""))))
    current_sources = {r.get("file_name"): r for r in merged.get("source_files", [])}
    for row in incoming.get("source_files", []):
        current_sources[row.get("file_name")] = row
    merged["source_files"] = sorted(current_sources.values(), key=lambda r: str(r.get("file_name", "")))
    return enrich_years(merged)


def compact_text(text: str, max_len: int = 220) -> str:
    text = re.sub(r"\s+", " ", text or "").strip()
    if len(text) <= max_len:
        return text
    return text[: max_len - 1].rstrip() + "…"


def risk_posture_label(high_count: int) -> str:
    if high_count >= 8:
        return "High exposure"
    if high_count >= 3:
        return "Watch closely"
    if high_count >= 1:
        return "Controlled watch"
    return "No high blockers logged"


def format_pct(value: Optional[float]) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "N/A"
    return f"{float(value):.2f}%"
