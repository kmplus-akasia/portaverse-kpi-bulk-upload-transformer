"""Resolve pre-restructure worksheets against historical Q1 assignment evidence.

This module is deliberately DB-free.  Its input is the read-only JSON payload
written by ``export_historical_q1_position_reference.mjs``.
"""

from __future__ import annotations

from collections import Counter
from difflib import SequenceMatcher
from typing import Any, Iterable

from position_mapping import normalize_position_lookup


STRUCTURAL = "structural"
NON_STRUCTURAL = "non_structural"
SCOPE_UNCERTAIN = "scope_uncertain"

HIGH_CONFIDENCE = "high_confidence"
LOW_CONFIDENCE = "low_confidence"
MAPPING_CONFLICT = "mapping_conflict"
NO_CANDIDATE = "no_candidate"
HEAD_OFFICE_COMPANY_ID = "1"

REPORT_COLUMNS = [
    "Source Workbook",
    "Worksheet",
    "Worksheet Position",
    "Worksheet Group",
    "Historical Employee Numbers",
    "Historical Employee Names",
    "Assignment Types",
    "Historical End Date",
    "Missing Historical Organization Evidence",
    "Inferred Scope",
    "Candidate PMID",
    "Candidate PNID",
    "Candidate Position Title",
    "Candidate Group",
    "Candidate Company",
    "Confidence Label",
    "Confidence Reason",
    "Existing Config PMID",
    "Existing Config PNID",
    "Reviewer Confirm Mapping",
    "Reviewer Actual PMID",
    "Reviewer Actual PNID",
    "Reviewer Notes",
]


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _identifier(value: Any) -> str:
    text = _text(value)
    return "" if text in {"", "0", "None"} else text


def _truthy(value: Any) -> bool:
    return value in (True, 1, "1", "true", "TRUE", "yes", "YES", "Y", "y")


def _unique(values: Iterable[Any]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        text = _text(value)
        if text and text not in seen:
            seen.add(text)
            result.append(text)
    return result


def historical_assignment_type(row: dict[str, Any]) -> str:
    """Classify a retained TEPMS row without filtering away secondary evidence."""
    if row.get("lakhar_id") not in (None, ""):
        return "LAKHAR"
    if row.get("job_sharing_id") not in (None, ""):
        return "JOB_SHARING"
    return "PRIMARY"


def _scope(row: dict[str, Any]) -> str:
    return STRUCTURAL if _text(row.get("position_master_type_id")) == "5" else NON_STRUCTURAL


def _worksheet_titles(position: dict[str, Any]) -> list[str]:
    titles: list[str] = []
    raw_lookup_names = position.get("position_lookup_names")
    if isinstance(raw_lookup_names, (list, tuple)):
        titles.extend(_text(value) for value in raw_lookup_names)
    titles.extend(
        _text(position.get(key))
        for key in ("position_name", "portaverse_position_title", "sheet_name")
    )
    return _unique(titles)


def _worksheet_groups(position: dict[str, Any]) -> list[str]:
    return _unique(position.get(key) for key in ("group_name", "portaverse_group_name"))


def _normalized_values(values: Iterable[str]) -> list[str]:
    return _unique(normalize_position_lookup(value) for value in values)


def _context_score(expected_values: list[str], candidate_values: Iterable[Any]) -> tuple[float, bool]:
    expected = _normalized_values(expected_values)
    candidates = _normalized_values(_text(value) for value in candidate_values)
    if not expected:
        return 1.0, False
    if not candidates:
        return 0.0, False
    score = 0.0
    exact = False
    for source in expected:
        for candidate in candidates:
            if source == candidate:
                exact = True
                score = max(score, 1.0)
            elif source in candidate or candidate in source:
                score = max(score, 0.85)
            else:
                source_tokens = set(source.split())
                candidate_tokens = set(candidate.split())
                overlap = len(source_tokens & candidate_tokens) / max(len(source_tokens), 1)
                score = max(score, min(overlap, 0.75))
    return score, exact


def _title_score(position: dict[str, Any], candidate_titles: Iterable[Any]) -> tuple[float, bool]:
    titles = _normalized_values(_worksheet_titles(position))
    candidates = _normalized_values(_text(value) for value in candidate_titles)
    if not titles or not candidates:
        return 0.0, False
    score = 0.0
    exact = False
    for source in titles:
        source_tokens = set(source.split())
        for candidate in candidates:
            if source == candidate:
                exact = True
                score = max(score, 1.0)
                continue
            if source in candidate or candidate in source:
                short_length = min(len(source), len(candidate))
                score = max(score, 0.88 if short_length >= 10 else 0.78)
                continue
            candidate_tokens = set(candidate.split())
            overlap = len(source_tokens & candidate_tokens) / max(len(source_tokens), 1)
            ratio = SequenceMatcher(None, source, candidate).ratio()
            score = max(score, max(overlap * 0.82, ratio * 0.8))
    return score, exact


def _matching_raw_rows(position: dict[str, Any], assignments: list[dict[str, Any]]) -> list[dict[str, Any]]:
    matched: list[dict[str, Any]] = []
    for assignment in assignments:
        title_score, _ = _title_score(position, [assignment.get("position_title")])
        group_score, group_exact = _context_score(
            _worksheet_groups(position), [assignment.get("group_name")]
        )
        if title_score >= 0.65 and (group_score >= 0.5 or group_exact or not _worksheet_groups(position)):
            matched.append(assignment)
    return matched


def _eligible_assignment(row: dict[str, Any], company_id: str) -> bool:
    return (
        _text(row.get("company_id")) == str(company_id)
        and not _truthy(row.get("missing_historical_organization"))
        and bool(_identifier(row.get("position_master_id")))
        and bool(_text(row.get("position_title")))
    )


def _group_assignments(assignments: list[dict[str, Any]]) -> dict[tuple[str, str], list[dict[str, Any]]]:
    grouped: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for assignment in assignments:
        key = (_scope(assignment), _identifier(assignment.get("position_master_id")))
        if not key[1]:
            continue
        grouped.setdefault(key, []).append(assignment)
    return grouped


def _nomenclature_by_pmid(rows: list[dict[str, Any]], company_id: str) -> dict[str, list[dict[str, Any]]]:
    by_pmid: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        if _text(row.get("company_id")) != str(company_id):
            continue
        pmid = _identifier(row.get("position_master_id"))
        pnid = _identifier(row.get("cluster_id"))
        if pmid and pnid:
            by_pmid.setdefault(pmid, []).append(row)
    return by_pmid


def _candidate_records(historical_payload: dict[str, Any], company_id: str) -> list[dict[str, Any]]:
    assignments = [
        row
        for row in historical_payload.get("historical_assignment_rows", [])
        if isinstance(row, dict) and _eligible_assignment(row, company_id)
    ]
    grouped = _group_assignments(assignments)
    nomenclature = _nomenclature_by_pmid(
        [row for row in historical_payload.get("nomenclature_rows", []) if isinstance(row, dict)],
        company_id,
    )
    candidates: list[dict[str, Any]] = []
    for (scope, pmid), evidence in grouped.items():
        if scope == STRUCTURAL:
            candidates.append(
                {
                    "identity": (STRUCTURAL, pmid),
                    "scope": STRUCTURAL,
                    "pmid": pmid,
                    "pnid": "",
                    "title_values": _unique(row.get("position_title") for row in evidence),
                    "group_values": _unique(row.get("group_name") for row in evidence),
                    "company_values": _unique(row.get("company_name") for row in evidence),
                    "evidence": evidence,
                }
            )
            continue
        for mapping in nomenclature.get(pmid, []):
            pnid = _identifier(mapping.get("cluster_id"))
            candidates.append(
                {
                    "identity": (NON_STRUCTURAL, pnid),
                    "scope": NON_STRUCTURAL,
                    "pmid": "",
                    "pnid": pnid,
                    "title_values": _unique(
                        [
                            *(row.get("position_title") for row in evidence),
                            mapping.get("cluster_label"),
                            mapping.get("position_name"),
                        ]
                    ),
                    "group_values": _unique(
                        [*(row.get("group_name") for row in evidence), mapping.get("group_name")]
                    ),
                    "company_values": _unique(
                        [*(row.get("company_name") for row in evidence), mapping.get("company_name")]
                    ),
                    "evidence": evidence,
                }
            )
    return _merge_candidate_records(candidates)


def _merge_candidate_records(candidates: list[dict[str, Any]]) -> list[dict[str, Any]]:
    merged: dict[tuple[str, str], dict[str, Any]] = {}
    for candidate in candidates:
        identity = candidate["identity"]
        existing = merged.get(identity)
        if existing is None:
            merged[identity] = {**candidate, "evidence": list(candidate["evidence"])}
            continue
        for key in ("title_values", "group_values", "company_values"):
            existing[key] = _unique([*existing[key], *candidate[key]])
        existing["evidence"].extend(candidate["evidence"])
    return list(merged.values())


def _score_candidate(position: dict[str, Any], candidate: dict[str, Any]) -> dict[str, Any]:
    title_score, title_exact = _title_score(position, candidate["title_values"])
    group_score, group_exact = _context_score(_worksheet_groups(position), candidate["group_values"])
    if title_exact and group_exact:
        rank = 3
    elif title_exact:
        rank = 2
    elif title_score >= 0.75 and group_score >= 0.75:
        rank = 1
    else:
        rank = 0
    score = round((title_score * 0.76) + (group_score * 0.24), 4)
    return {
        **candidate,
        "title_score": title_score,
        "group_score": group_score,
        "rank": rank,
        "score": score,
        "reason": f"title={title_score:.2f}; group={group_score:.2f}",
    }


def _comparison_index(existing_config: dict[str, Any] | list[dict[str, Any]]) -> dict[tuple[str, str], dict[str, Any]]:
    positions = existing_config.get("positions", []) if isinstance(existing_config, dict) else existing_config
    index: dict[tuple[str, str], dict[str, Any]] = {}
    for position in positions or []:
        if isinstance(position, dict):
            index[(_text(position.get("source_workbook")), _text(position.get("sheet_name")))] = position
    return index


def _candidate_value(candidate: dict[str, Any] | None, key: str) -> str:
    if not candidate:
        return ""
    values = candidate.get(key, [])
    return _text(values[0]) if values else ""


def _evidence_fields(rows: list[dict[str, Any]]) -> dict[str, str]:
    return {
        "Historical Employee Numbers": "; ".join(_unique(row.get("employee_number") for row in rows)),
        "Historical Employee Names": "; ".join(_unique(row.get("employee_name") for row in rows)),
        "Assignment Types": "; ".join(
            _unique(row.get("assignment_type") or historical_assignment_type(row) for row in rows)
        ),
        "Historical End Date": "; ".join(
            _unique(row.get("assignment_end_date") or row.get("end_date") for row in rows)
        ),
        "Missing Historical Organization Evidence": "YES"
        if any(_truthy(row.get("missing_historical_organization")) for row in rows)
        else "NO",
    }


def _has_primary_secondary_identity_conflict(candidates: list[dict[str, Any]]) -> bool:
    primary_identities_by_employee: dict[str, set[tuple[str, str]]] = {}
    secondary_identities_by_employee: dict[str, set[tuple[str, str]]] = {}
    for candidate in candidates:
        for evidence in candidate["evidence"]:
            employee_number = _text(evidence.get("employee_number"))
            if not employee_number:
                continue
            identities = (
                primary_identities_by_employee
                if historical_assignment_type(evidence) == "PRIMARY"
                else secondary_identities_by_employee
            )
            identities.setdefault(employee_number, set()).add(candidate["identity"])
    return any(
        primary_identities_by_employee.get(employee_number, set())
        - secondary_identities
        for employee_number, secondary_identities in secondary_identities_by_employee.items()
    )


def _validate_head_office_scope(historical_payload: dict[str, Any], company_id: str) -> None:
    if _text(company_id) != HEAD_OFFICE_COMPANY_ID:
        raise ValueError("Historical Q1 mapping supports only Head Office company ID '1'.")
    source = historical_payload.get("source")
    if isinstance(source, dict) and "company_id" in source:
        if _text(source.get("company_id")) != HEAD_OFFICE_COMPANY_ID:
            raise ValueError("Historical payload source company_id must be '1'.")


def _fallback_scope(position: dict[str, Any]) -> str:
    configured = _text(position.get("position_scope"))
    if configured in {STRUCTURAL, NON_STRUCTURAL}:
        return configured
    title = " ".join(_worksheet_titles(position))
    normalized = normalize_position_lookup(title)
    if any(role in normalized for role in ("group head", "department head", "division head", "manager", "team lead", "deputy", "pimpinan proyek")):
        return STRUCTURAL
    if any(role in normalized for role in ("officer", "auditor", "analyst", "specialist", "staff", "expert")):
        return NON_STRUCTURAL
    return SCOPE_UNCERTAIN


def build_mapping_rows(
    positions: list[dict[str, Any]],
    historical_payload: dict[str, Any],
    existing_config: dict[str, Any] | list[dict[str, Any]],
    company_id: str = "1",
) -> list[dict[str, Any]]:
    """Build one unapproved, evidence-backed report row for every worksheet key."""
    _validate_head_office_scope(historical_payload, str(company_id))
    comparison = _comparison_index(existing_config)
    raw_assignments = [
        row
        for row in historical_payload.get("historical_assignment_rows", [])
        if isinstance(row, dict)
    ]
    candidates = _candidate_records(historical_payload, str(company_id))
    rows: list[dict[str, Any]] = []
    for position in positions:
        source_workbook = _text(position.get("source_workbook"))
        worksheet = _text(position.get("sheet_name"))
        matching_raw = _matching_raw_rows(position, raw_assignments)
        scored = sorted(
            (_score_candidate(position, candidate) for candidate in candidates),
            key=lambda candidate: (candidate["rank"], candidate["score"], candidate["scope"], candidate["pmid"], candidate["pnid"]),
            reverse=True,
        )
        eligible = [candidate for candidate in scored if candidate["rank"] > 0]
        best = eligible[0] if eligible else None
        strong_competitors = (
            [candidate for candidate in eligible if candidate["rank"] == best["rank"]]
            if best
            else []
        )
        primary_secondary_conflict = _has_primary_secondary_identity_conflict(eligible)
        conflict = len(strong_competitors) > 1 or primary_secondary_conflict
        if not best:
            label = NO_CANDIDATE
            reason = "No historical company-1 candidate matched the worksheet title and group."
            proposed = None
            inferred_scope = _fallback_scope(position)
        elif conflict:
            label = MAPPING_CONFLICT
            reason = "Multiple equally ranked historical identities require reviewer selection."
            if primary_secondary_conflict:
                reason = (
                    "PRIMARY assignment evidence takes precedence, but conflicting secondary "
                    "identity remains visible for reviewer selection."
                )
            proposed = None
            scopes = {candidate["scope"] for candidate in strong_competitors}
            inferred_scope = best["scope"] if len(scopes) == 1 else SCOPE_UNCERTAIN
        else:
            proposed = best
            inferred_scope = best["scope"]
            if best["rank"] >= 2 or (best["rank"] == 1 and best["score"] >= 0.86):
                label = HIGH_CONFIDENCE
                reason = "Unique historical candidate has exact or strong title/group evidence."
            else:
                label = LOW_CONFIDENCE
                reason = "Historical candidate is plausible but title or group evidence is weak."

        existing = comparison.get((source_workbook, worksheet), {})
        row = {
            "Source Workbook": source_workbook,
            "Worksheet": worksheet,
            "Worksheet Position": _text(position.get("position_name")) or worksheet,
            "Worksheet Group": _text(position.get("group_name")),
            **_evidence_fields(matching_raw),
            "Inferred Scope": inferred_scope,
            "Candidate PMID": proposed["pmid"] if proposed else "",
            "Candidate PNID": proposed["pnid"] if proposed else "",
            "Candidate Position Title": _candidate_value(proposed, "title_values"),
            "Candidate Group": _candidate_value(proposed, "group_values"),
            "Candidate Company": _candidate_value(proposed, "company_values"),
            "Confidence Label": label,
            "Confidence Reason": f"{reason} {proposed['reason']}" if proposed else reason,
            "Existing Config PMID": _identifier(existing.get("position_master_id")),
            "Existing Config PNID": _identifier(existing.get("position_nomenclature_id")),
            "Reviewer Confirm Mapping": "",
            "Reviewer Actual PMID": "",
            "Reviewer Actual PNID": "",
            "Reviewer Notes": "",
        }
        errors = validate_mapping_row(row)
        if errors:
            row["Confidence Label"] = MAPPING_CONFLICT
            row["Confidence Reason"] = f"{' '.join(errors)} Reviewer selection required."
            row["Candidate PMID"] = ""
            row["Candidate PNID"] = ""
        rows.append({column: row.get(column, "") for column in REPORT_COLUMNS})
    return rows


def validate_mapping_row(row: dict[str, Any]) -> list[str]:
    """Return namespace violations that would block an automatic proposal."""
    pmid = _identifier(row.get("Candidate PMID"))
    pnid = _identifier(row.get("Candidate PNID"))
    scope = _text(row.get("Inferred Scope"))
    errors: list[str] = []
    if pmid and pnid:
        errors.append("Both Candidate PMID and Candidate PNID are populated.")
    if scope == STRUCTURAL and pnid:
        errors.append("Structural mappings must not propose a PNID.")
    if scope == NON_STRUCTURAL and pmid:
        errors.append("Non-structural mappings must not propose a PMID.")
    return errors


def mapping_summary(rows: list[dict[str, Any]], historical_payload: dict[str, Any]) -> dict[str, Any]:
    confidence = Counter(_text(row.get("Confidence Label")) for row in rows)
    return {
        "source": historical_payload.get("source", {}),
        "mapping_rows": len(rows),
        "source_workbooks": len({_text(row.get("Source Workbook")) for row in rows}),
        "historical_assignment_rows": len(historical_payload.get("historical_assignment_rows", [])),
        "confidence": dict(sorted(confidence.items())),
        "reviewer_approved_rows": 0,
        "all_rows_unapproved": all(not _text(row.get("Reviewer Confirm Mapping")) for row in rows),
    }
