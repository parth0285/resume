"""Helpers to turn extracted faculty/publication JSON records into clean
tabular data for on-screen tables and the two-sheet Excel export."""

import io
import pandas as pd

MASTER_COLUMN_ORDER = [
    "source_file", "name", "emails", "phones", "nationality", "country", "address",
    "current_designation", "current_department", "current_organization",
    # education_*_institution is a single merged field for each degree level
    # containing the combined institute/college and university name when both are present.
    "education_ug_degree", "education_ug_branch", "education_ug_institution", "education_ug_year",
    "education_pg_degree", "education_pg_branch", "education_pg_institution", "education_pg_year",
    "education_phd_degree", "education_phd_branch", "education_phd_institution", "education_phd_year",
    "experience_academic_years", "experience_industry_years", "experience_research_years",
    "experience_administrative_years", "experience_total_years", "experience_stated_total_years",
    "counts_journal", "counts_international_conference", "counts_national_conference",
    "counts_books", "counts_book_chapters", "counts_patents", "counts_projects",
    "counts_sponsored_projects", "counts_consultancy_projects", "counts_trainings", "counts_projects_general",
    "counts_awards", "counts_memberships", "counts_fdp_attended", "counts_sttp_attended",
    "counts_workshops_attended", "counts_online_courses",
    "profiles_orcid", "profiles_google_scholar", "profiles_researchgate", "profiles_linkedin", "profiles_scopus",
    "designation_history", "administrative_roles", "sponsored_projects", "consultancy_projects",
    "trainings_conducted", "projects_list", "awards_list", "memberships_list", "fdp_list", "sttp_list",
    "workshops_attended_list", "online_courses_list", "flags"
]

PUBLICATION_COLUMN_ORDER = [
    "source_file", "faculty_name", "sr_no", "title", "authors", "publication_type",
    "journal_or_conference", "publisher", "volume", "issue", "pages", "year", "month",
    "doi", "issn", "isbn", "indexed_in", "quartile", "impact_factor", "citation_count", "url",
]


def _join_list(value):
    if isinstance(value, list):
        return "; ".join(str(v) for v in value if v not in (None, ""))
    return value


def _join_records(value):
    if not isinstance(value, list):
        return value
    if not value:
        return ""
    if isinstance(value[0], dict):
        parts = []
        for d in value:
            title = d.get("title", "")
            date_range = d.get("date_range", "")
            if title and date_range:
                parts.append(f"{title} ({date_range})")
            elif title:
                parts.append(title)
        return "; ".join(parts)
    return "; ".join(str(v) for v in value if v not in (None, ""))


def _flatten_pg_education(rec: dict) -> dict:
    """education.pg is now extracted as an ARRAY (to support multiple
    postgraduate-level degrees, e.g. an MCA earned before an M.Tech). For the
    flat Excel/table export, join multiple entries' degree/branch/institution
    /year into semicolon-separated strings within the SAME existing columns
    (education_pg_degree, etc.) — this keeps the export column structure
    backward-compatible while still surfacing every degree, rather than
    silently dropping all but one."""
    education = rec.get("education") or {}
    pg = education.get("pg")
    if isinstance(pg, list):
        degrees = "; ".join(p.get("degree", "") for p in pg if isinstance(p, dict) and p.get("degree"))
        branches = "; ".join(p.get("branch", "") for p in pg if isinstance(p, dict) and p.get("branch"))
        institutions = "; ".join(p.get("institution", "") for p in pg if isinstance(p, dict) and p.get("institution"))
        years = "; ".join(str(p.get("year", "")) for p in pg if isinstance(p, dict) and p.get("year"))
        education = dict(education)
        education["pg"] = {"degree": degrees, "branch": branches, "institution": institutions, "year": years}
        rec = dict(rec)
        rec["education"] = education
    return rec


def master_records_to_dataframe(records: list) -> pd.DataFrame:
    """Flatten nested master_faculty_database records into a flat DataFrame."""
    if not records:
        return pd.DataFrame(columns=MASTER_COLUMN_ORDER)

    # Join list fields before normalizing so they render as plain strings, not
    # python list reprs, in both the on-screen table and the Excel sheet.
    prepped = []
    for rec in records:
        rec = _flatten_pg_education(rec)
        rec = dict(rec)
        rec["emails"] = _join_records(rec.get("emails"))
        rec["phones"] = _join_records(rec.get("phones"))
        rec["designation_history"] = _join_records(rec.get("designation_history"))
        rec["administrative_roles"] = _join_records(rec.get("administrative_roles"))
        rec["sponsored_projects"] = _join_records(rec.get("sponsored_projects"))
        rec["consultancy_projects"] = _join_records(rec.get("consultancy_projects"))
        rec["trainings_conducted"] = _join_records(rec.get("trainings_conducted"))
        rec["projects_list"] = _join_records(rec.get("projects_list"))
        rec["awards_list"] = _join_records(rec.get("awards_list"))
        rec["memberships_list"] = _join_records(rec.get("memberships_list"))
        rec["fdp_list"] = _join_records(rec.get("fdp_list"))
        rec["sttp_list"] = _join_records(rec.get("sttp_list"))
        rec["workshops_attended_list"] = _join_records(rec.get("workshops_attended_list"))
        rec["online_courses_list"] = _join_records(rec.get("online_courses_list"))
        rec["flags"] = _join_records(rec.get("flags"))
        prepped.append(rec)

    df = pd.json_normalize(prepped, sep="_")

    # Ensure every expected column exists, in a stable, readable order.
    for col in MASTER_COLUMN_ORDER:
        if col not in df.columns:
            df[col] = None
    extra_cols = [c for c in df.columns if c not in MASTER_COLUMN_ORDER]
    return df[MASTER_COLUMN_ORDER + extra_cols]


def publication_records_to_dataframe(records: list) -> pd.DataFrame:
    """Flatten publication_details records into a flat DataFrame."""
    if not records:
        return pd.DataFrame(columns=PUBLICATION_COLUMN_ORDER)

    prepped = []
    for rec in records:
        rec = dict(rec)
        rec["authors"] = _join_list(rec.get("authors"))
        rec["indexed_in"] = _join_list(rec.get("indexed_in"))
        prepped.append(rec)

    df = pd.json_normalize(prepped, sep="_")

    for col in PUBLICATION_COLUMN_ORDER:
        if col not in df.columns:
            df[col] = None
    extra_cols = [c for c in df.columns if c not in PUBLICATION_COLUMN_ORDER]
    return df[PUBLICATION_COLUMN_ORDER + extra_cols]


def build_excel_workbook(master_df: pd.DataFrame, publication_df: pd.DataFrame) -> bytes:
    """Build a two-sheet .xlsx workbook in memory and return its bytes."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        master_df.to_excel(writer, sheet_name="Master Faculty Database", index=False)
        publication_df.to_excel(writer, sheet_name="Publication Details", index=False)

        # Light auto-fit so the exported sheets are usable without manual resizing.
        for sheet_name, df in (
            ("Master Faculty Database", master_df),
            ("Publication Details", publication_df),
        ):
            worksheet = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns, start=1):
                max_len = max(
                    [len(str(col))] + [len(str(v)) for v in df[col].astype(str).tolist()]
                )
                worksheet.column_dimensions[worksheet.cell(row=1, column=i).column_letter].width = min(
                    max(12, max_len + 2), 60
                )

    buffer.seek(0)
    return buffer.getvalue()


def build_single_sheet_workbook(df: pd.DataFrame, sheet_name: str) -> bytes:
    """Build a single-sheet .xlsx workbook in memory and return its bytes."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        for i, col in enumerate(df.columns, start=1):
            max_len = max([len(str(col))] + [len(str(v)) for v in df[col].astype(str).tolist()])
            worksheet.column_dimensions[worksheet.cell(row=1, column=i).column_letter].width = min(max(12, max_len + 2), 60)
    buffer.seek(0)
    return buffer.getvalue()
