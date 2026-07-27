import logging
import re
import os
import json
from typing import Optional
from date_utils import parse_year_range

logger = logging.getLogger(__name__)

def _classify_pub_type(ptype: str):
    """Map a model's free-text publication_type to our fixed count buckets.
    Different models phrase this differently ("International Conference Paper",
    "Conf. (National)", "Journal Article", etc.), so match on keywords rather
    than requiring an exact string — an exact-match dict silently drops any
    phrasing variant to 0 with no visibility into why."""
    t = (ptype or "").strip().lower()
    if not t:
        return None
    if "patent" in t:
        return "patents"
    if "book" in t and "chapter" in t:
        return "book_chapters"
    if "book" in t:
        return "books"
    if "conference" in t or "symposium" in t or "seminar" in t:
        if "international" in t or "intl" in t:
            return "international_conference"
        if "national" in t:
            return "national_conference"
        # Conference type stated but scope (national/international) ambiguous —
        # don't silently drop it; fall through to journal-style default only
        # if nothing else matches, otherwise leave unclassified.
        return None
    if "journal" in t:
        return "journal"
    return None


def compute_all_counts(record: dict, publications: list) -> dict:
    counts = record.get("counts") or {}

    tally = {
        "journal": 0, "international_conference": 0, "national_conference": 0,
        "books": 0, "book_chapters": 0, "patents": 0,
    }
    unclassified_pub_types = []
    for pub in publications:
        ptype_raw = pub.get("publication_type") or ""
        key = _classify_pub_type(ptype_raw)
        if key:
            tally[key] += 1
        elif ptype_raw.strip():
            unclassified_pub_types.append(ptype_raw.strip())

    flags = record.get("flags") or []

    # CRITICAL: do NOT blindly overwrite the model's counts with the itemized
    # tally. If the model already provided a count for a field (e.g. because
    # it read an aggregate summary table per the "counts_from_summary_table"
    # rule, which often states a HIGHER, more authoritative total than what
    # could be itemized into individual publication_details rows), silently
    # replacing it with a possibly-lower itemized tally throws away real
    # data. Instead, take whichever value is larger, and flag it when the
    # two sources disagree so a reviewer can see the discrepancy rather than
    # having it silently resolved one way.
    mismatches = []
    for key, itemized_value in tally.items():
        model_value = counts.get(key)
        if isinstance(model_value, (int, float)) and model_value > 0:
            model_value = int(model_value)
            if model_value != itemized_value:
                mismatches.append(
                    f"{key}: summary/model count={model_value} vs itemized publication_details count={itemized_value}"
                )
            counts[key] = max(model_value, itemized_value)
        else:
            counts[key] = itemized_value

    if mismatches:
        flags.append(f"count_mismatch_model_vs_itemized: {'; '.join(mismatches)}")

    if unclassified_pub_types:
        flags.append(f"unclassified_publication_type: {'; '.join(unclassified_pub_types)}")

    sponsored = record.get("sponsored_projects") or []
    consultancy = record.get("consultancy_projects") or []
    trainings = record.get("trainings_conducted") or []
    generic = record.get("projects_list") or []

    counts["sponsored_projects"] = len(sponsored)
    counts["consultancy_projects"] = len(consultancy)
    counts["trainings"] = len(trainings)
    counts["projects_general"] = len(generic)
    
    # Generic projects bucket used as total fallback, or we can just count it
    # total projects in the schema might just be sum of all, or we leave it separate
    # The prompt will ask for projects_list if undifferentiated.
    # NOTE: trainings/FDPs are already counted separately in counts["trainings"];
    # they must NOT be folded into the projects total or every FDP entry gets
    # double-counted as a "project" as well.
    counts["projects"] = len(sponsored) + len(consultancy) + len(generic)

    # Dedupe near-identical award entries (e.g. the same medal described
    # twice with slightly different wording) before counting/exporting.
    deduped_awards = dedupe_awards(record.get("awards_list") or [])
    record["awards_list"] = deduped_awards
    counts["awards"] = len(deduped_awards)

    counts["memberships"] = len(record.get("memberships_list") or [])
    counts["fdp_attended"] = len(record.get("fdp_list") or [])
    counts["sttp_attended"] = len(record.get("sttp_list") or [])
    counts["workshops_attended"] = len(record.get("workshops_attended_list") or [])
    counts["online_courses"] = len(record.get("online_courses_list") or [])

    record["counts"] = counts
    record["flags"] = flags
    return record


_RESEARCH_KEYWORDS = (
    "research associate", "research scientist", "research fellow", "research scholar",
    "postdoc", "post-doc", "post doctoral", "postdoctoral", "post-doctoral",
    "postdoctoral research fellow", "postdoctoral fellow", "postdoctoral researcher",
    "doctoral research", "doctoral research fellow", "doctoral fellow", "doctoral",
    "phd scholar", "ph.d.", "phd candidate", "postgraduate", "postgraduate scholar",
    "post grad", "post-grad", "research assistant", "scientist",
)
_ACADEMIC_KEYWORDS = (
    "professor", "lecturer", "instructor", "reader", "dean", "principal",
    "head of department", "hod", "assistant prof", "associate prof", "faculty",
)
_INDUSTRY_KEYWORDS = (
    "engineer", "developer", "manager", "consultant", "analyst", "executive",
    "officer", "specialist", "architect", "industry",
)

def _resolve_category(desig: dict) -> Optional[str]:
    """Trust the model's category if it's one of the three valid values.
    Otherwise fall back to keyword matching on title/organization so a
    designation isn't silently dropped from experience calculation just
    because the model left `category` blank (which small/free models do
    often, despite the prompt asking for it)."""
    given = (desig.get("category") or "").strip().lower()
    if given in ("academic", "industry", "research"):
        return given

    text = f"{desig.get('title') or ''} {desig.get('organization') or ''}".lower()
    for kw in _RESEARCH_KEYWORDS:
        if kw in text:
            return "research"
    for kw in _ACADEMIC_KEYWORDS:
        if kw in text:
            return "academic"
    for kw in _INDUSTRY_KEYWORDS:
        if kw in text:
            return "industry"
    return None


def _merge_intervals(intervals):
    if not intervals:
        return 0
    intervals.sort(key=lambda x: x[0])
    merged = [intervals[0]]
    for current in intervals[1:]:
        previous = merged[-1]
        if current[0] <= previous[1]:
            merged[-1] = (previous[0], max(previous[1], current[1]))
        else:
            merged.append(current)
    return sum(end - start for start, end in merged)

def compute_experience(record: dict) -> dict:
    designations = record.get("designation_history") or []
    flags = record.get("flags") or []

    academic_intervals = []
    industry_intervals = []
    research_intervals = []
    admin_intervals = []
    unclassified_intervals = []

    if not designations:
        # This is the #1 cause of experience showing as 0: the model didn't
        # return any designation_history entries at all, even though the
        # resume may clearly list employment history. Surface this loudly
        # in the exported "flags" column instead of silently reporting 0.
        flags.append("designation_history_empty_from_model")

    unparseable_dates = []
    current_role_titles = []
    for desig in designations:
        date_range = desig.get("date_range") or ""
        cat = _resolve_category(desig)
        start, end, is_current = parse_year_range(date_range)
        if is_current:
            current_role_titles.append(
                f"{desig.get('title') or '?'} @ {desig.get('organization') or '?'}"
            )
        if start and end and start <= end:
            interval = (start, end)
            if cat == "academic":
                academic_intervals.append(interval)
            elif cat == "industry":
                industry_intervals.append(interval)
            elif cat == "research":
                research_intervals.append(interval)
            else:
                # Couldn't classify even with keyword fallback — still real
                # experience, so don't let it vanish from total_years.
                unclassified_intervals.append(interval)
        else:
            unparseable_dates.append(f"{desig.get('title') or '?'} [{date_range or 'no date_range given'}]")
            logger.warning(
                f"Could not parse date_range '{date_range}' for designation "
                f"'{desig.get('title')}' (source: {desig.get('organization')})."
            )

    if designations and unparseable_dates:
        flags.append(f"designation_dates_unparseable: {'; '.join(unparseable_dates)}")

    # Surface (not silently resolve) the case where more than one role is
    # open-ended/"Present" at once — e.g. two concurrent appointments at
    # different institutions. Downstream code still picks one as "current
    # organization" (rule 16/17 in the prompt: latest designation), but a
    # reviewer should be able to see this was ambiguous.
    if len(current_role_titles) > 1:
        flags.append(f"multiple_concurrent_current_roles: {'; '.join(current_role_titles)}")

    admin_roles = record.get("administrative_roles") or []
    for role in admin_roles:
        date_range = role.get("date_range") if isinstance(role, dict) else str(role)
        start, end, _ = parse_year_range(date_range)
        if start and end and start <= end:
            admin_intervals.append((start, end))

    exp = record.get("experience") or {}
    exp["academic_years"] = _merge_intervals(academic_intervals)
    exp["industry_years"] = _merge_intervals(industry_intervals)
    exp["research_years"] = _merge_intervals(research_intervals)
    exp["administrative_years"] = _merge_intervals(admin_intervals)
    
    total_intervals = (
        academic_intervals + industry_intervals + research_intervals + unclassified_intervals
    )
    exp["total_years"] = _merge_intervals(total_intervals)
    # stated_total_years is extracted verbatim by the model from the resume's
    # own summary (e.g. "12 years 5 months") — preserved as-is here (it's
    # already in `exp` since `exp` is the same dict as record["experience"]),
    # purely for comparison against the computed total_years above. Never
    # recomputed or overwritten by this function.
    exp.setdefault("stated_total_years", "")

    record["experience"] = exp
    record["flags"] = flags
    return record


def _normalize_title(title: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (title or "").lower())

def dedupe_publications(pubs: list) -> list:
    """Dedupe by (normalized title, publication_type) rather than title alone.
    A paper legitimately conference-published and later journal-published
    (or vice versa) is two distinct entries in most CVs' own publication
    lists and must not be collapsed into one just because the title matches."""
    seen = set()
    result = []
    for pub in pubs:
        title_key = _normalize_title(pub.get("title", ""))
        type_key = (pub.get("publication_type") or "").strip().lower()
        key = (title_key, type_key)
        if not title_key or key in seen:
            continue
        seen.add(key)
        result.append(pub)
    return result


def _normalize_award_text(text: str) -> str:
    """Loosely normalize an award description for near-duplicate detection:
    lowercase, strip punctuation, and collapse whitespace. This intentionally
    only catches close wording matches (same words, different order/casing/
    punctuation) — it will not catch two genuinely different awards that
    happen to share a keyword like 'Gold Medal'."""
    t = (text or "").lower()
    t = re.sub(r"[^a-z0-9\s]", " ", t)
    words = sorted(t.split())
    return " ".join(words)


def dedupe_awards(awards: list) -> list:
    """Remove near-duplicate award entries that describe the same underlying
    award with slightly different wording (e.g. two lines both describing an
    'Inter IIT Gold Medal' with different levels of detail). Keeps the
    longer/more descriptive entry when a near-duplicate is found."""
    if not awards:
        return awards

    def _text_of(a):
        return a if isinstance(a, str) else (a.get("title") or a.get("name") or str(a))

    kept = []
    kept_keys = []
    for award in awards:
        text = _text_of(award)
        key_words = set(_normalize_award_text(text).split())
        is_dup = False
        for i, existing_words in enumerate(kept_keys):
            if not key_words or not existing_words:
                continue
            overlap = len(key_words & existing_words) / max(len(key_words), len(existing_words))
            if overlap >= 0.5:
                # Near-duplicate — keep whichever text is longer/more descriptive.
                if len(text) > len(_text_of(kept[i])):
                    kept[i] = award
                    kept_keys[i] = key_words
                is_dup = True
                break
        if not is_dup:
            kept.append(award)
            kept_keys.append(key_words)
    return kept


_PG_LEVEL_DEGREE_ABBREVIATIONS = {
    "msc", "ma", "mcom", "mtech", "mba", "mphil", "me", "ms", "mca",
}
_PG_LEVEL_DEGREE_FULLNAMES = (
    "master of science", "master of arts", "master of technology",
    "master of business administration", "master of philosophy",
    "master of engineering", "master of computer applications",
    "master of commerce",
)

def _normalize_degree_str(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())

def normalize_education(record: dict) -> dict:
    """Two responsibilities:
    1. Ensure education.pg is always a list internally (the model may still
       occasionally return a single dict despite the schema/prompt asking
       for an array) so downstream export code can rely on a consistent shape.
    2. Detect a common misclassification: a postgraduate-level degree
       (M.Sc, M.Tech, MBA, etc.) placed in the UG slot. Matching is done on
       a punctuation/whitespace-normalized form (so "M. Sc", "M.Sc.", and
       "MSc" all match the same way) plus a full-name check, rather than a
       naive substring check on the raw string. Rather than silently leaving
       this wrong or guessing how to move it, flag it clearly for manual
       review — moving it automatically risks losing branch/year pairing if
       the extraction is inconsistent in other ways too."""
    flags = record.get("flags") or []
    education = record.get("education") or {}

    pg = education.get("pg")
    if isinstance(pg, dict):
        education["pg"] = [pg] if pg else []
    elif pg is None:
        education["pg"] = []

    ug = education.get("ug") or {}
    ug_degree_raw = (ug.get("degree") or "").strip()
    if ug_degree_raw:
        normalized = _normalize_degree_str(ug_degree_raw)
        lower = ug_degree_raw.lower()
        is_pg_level = (
            normalized in _PG_LEVEL_DEGREE_ABBREVIATIONS
            or any(full in lower for full in _PG_LEVEL_DEGREE_FULLNAMES)
        )
        if is_pg_level:
            flags.append(
                f"possible_ug_pg_misclassification: '{ug_degree_raw}' appears to be a "
                "postgraduate-level degree but was placed in the UG slot — please verify manually."
            )

    record["education"] = education
    record["flags"] = flags
    return record


def safe_process_record(record: dict, publications: list) -> dict:
    try:
        record = normalize_education(record)
    except Exception as e:
        logger.error(f"Education normalization failed for {record.get('source_file')}: {e}")
    try:
        # Debugging aid: log raw designation_history before experience calc.
        # This is intentionally lightweight and does NOT make any external calls.
        logger.info(f"RAW designation_history for {record.get('source_file')}: {record.get('designation_history')}")
        # Also append to a local debug file for offline inspection.
        try:
            dump_dir = os.path.join(os.getcwd(), "debug_logs")
            os.makedirs(dump_dir, exist_ok=True)
            dump_path = os.path.join(dump_dir, "designation_history_dump.jsonl")
            with open(dump_path, "a", encoding="utf-8") as f:
                json.dump({
                    "source_file": record.get("source_file"),
                    "designation_history": record.get("designation_history"),
                }, f, ensure_ascii=False)
                f.write("\n")
        except Exception:
            # Never let debug dumping break processing
            logger.debug("Failed to append designation_history to debug file.")
        record = compute_experience(record)
    except Exception as e:
        logger.error(f"Experience calc failed for {record.get('source_file')}: {e}")
    try:
        record = compute_all_counts(record, publications)
    except Exception as e:
        logger.error(f"Count calc failed for {record.get('source_file')}: {e}")
    return record


def renumber_publications(pubs: list) -> list:
    """Assign clean, gapless sr_no values in Python — never trust the model's count."""
    for i, pub in enumerate(pubs, start=1):
        pub["sr_no"] = i
    return pubs
