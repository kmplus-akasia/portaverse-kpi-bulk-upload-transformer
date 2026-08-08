"""Strict worksheet-to-position mapping for KPI converter."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from dataclasses import replace
from difflib import SequenceMatcher
from typing import Any


STRUCTURAL = "structural"
NON_STRUCTURAL = "non_structural"
SCOPE_UNCERTAIN = "scope_uncertain"

HIGH_CONFIDENCE = "high_confidence"
LOW_CONFIDENCE = "low_confidence"
NO_CANDIDATE = "no_candidate"
MAPPING_CONFLICT = "mapping_conflict"

BLOCKED_CONFIDENCE_LABELS = {
    LOW_CONFIDENCE,
    SCOPE_UNCERTAIN,
    NO_CANDIDATE,
    MAPPING_CONFLICT,
}

STRUCTURAL_TYPE_IDS = {"5"}

REPORT_COLUMNS = [
    "Source Workbook",
    "Worksheet",
    "Raw Worksheet Title",
    "Normalized Worksheet Title",
    "Inferred Scope",
    "Confidence Label",
    "Confidence Reason",
    "Candidate PMID",
    "Candidate PNID",
    "Candidate Title",
    "Candidate Group",
    "Candidate Company",
    "Candidate Company Code",
    "Candidate Score",
    "Runner-up Scope",
    "Runner-up PMID",
    "Runner-up PNID",
    "Runner-up Title",
    "Runner-up Score",
    "Active Variant Count",
    "Active Employee Count",
    "Active Employee Name",
    "Active Employee NIPP",
    "Recommended Action",
]


@dataclass(frozen=True)
class ScopeInference:
    scope: str
    normalized_title: str
    reason: str


@dataclass(frozen=True)
class LookupCandidate:
    scope: str
    position_master_id: str | None
    position_nomenclature_id: str | None
    title: str | None
    group_name: str | None
    company_id: str | None
    company_name: str | None
    company_code: str | None
    active_variant_count: int
    active_employee_count: int
    definitive_employee_count: int
    secondary_employee_count: int
    active_employee_names: tuple[str, ...] = ()
    active_employee_nipps: tuple[str, ...] = ()
    group_ancestor_names: tuple[str, ...] = ()
    lookup_keys: tuple[str, ...] = ()
    source_row: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class ScoredCandidate:
    candidate: LookupCandidate
    score: float
    reason: str


@dataclass(frozen=True)
class LookupIndexes:
    structural: list[LookupCandidate]
    non_structural: list[LookupCandidate]
    structural_by_pmid: dict[str, LookupCandidate]
    non_structural_by_pnid: dict[str, LookupCandidate]
    source: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class MappingResolution:
    source_workbook: str | None
    worksheet: str
    raw_worksheet_title: str
    normalized_worksheet_title: str
    inferred_scope: str
    confidence_label: str
    confidence_reason: str
    position_master_id: str | None = None
    position_nomenclature_id: str | None = None
    candidate_title: str | None = None
    candidate_group: str | None = None
    candidate_company: str | None = None
    candidate_company_code: str | None = None
    candidate_score: float | None = None
    runner_up_scope: str | None = None
    runner_up_position_master_id: str | None = None
    runner_up_position_nomenclature_id: str | None = None
    runner_up_title: str | None = None
    runner_up_score: float | None = None
    active_variant_count: int | None = None
    active_employee_count: int | None = None
    active_employee_name: str | None = None
    active_employee_nipp: str | None = None
    upload_allowed: bool = False


@dataclass(frozen=True)
class OverrideValidation:
    allowed: bool
    reason: str
    candidate: LookupCandidate | None = None


def norm_text(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def normalize_title(value: Any) -> str:
    text = (norm_text(value) or "").lower().replace("&", " dan ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_position_lookup(value: Any) -> str:
    text = normalize_title(value)
    text = re.sub(r"\b([a-z]+)([0-9]+)\b", r"\1 \2", text)
    text = re.sub(r"\bp\s+([0-9]+)\b", r"proyek \1", text)
    replacements = [
        (r"\bdivhead\b", "division head"),
        (r"\bdepthead\b", "department head"),
        (r"\bspv\b", "supervisor"),
        (r"\btl\b", "team lead"),
        (r"\bcorsec\b", "corporate secretary"),
        (r"\bmanrisk\b", "manajemen risiko"),
        (r"\bmonev\b", "monitoring evaluasi"),
        (r"\bfaspel\b", "fasilitas pelabuhan"),
        (r"\bmanagr\b", "manager"),
        (r"\bmanajer\b", "manager"),
        (r"\bmgr\b", "manager"),
        (r"\basst\b", "assistant"),
        (r"\basisten\b", "assistant"),
        (r"\bops\b", "operasi"),
        (r"\bsupt\b", "superintendent"),
        (r"\boficer\b", "officer"),
        (r"\boffice\b", "officer"),
        (r"\bprinciple\b", "principal"),
        (r"\bpimpro\b", "pimpinan proyek"),
        (r"\bdept\b", "department"),
        (r"\bdiv\b", "division"),
        (r"\bdh\b", "department head"),
        (r"\bgh\b", "group head"),
    ]
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text)
    text = re.sub(
        r"\b(group head|department head|division head|regional manager|manager|senior officer|officer|superintendent|assistant)\s+(i|ii|iii|iv|v)\b",
        r"\1",
        text,
    )
    return re.sub(r"\s+", " ", text).strip()


def _has_phrase(text: str, phrase: str) -> bool:
    return bool(re.search(rf"(^|\s){re.escape(phrase)}($|\s)", text))


def infer_worksheet_scope(worksheet_title: str | None) -> ScopeInference:
    normalized = normalize_position_lookup(worksheet_title)
    structural_phrases = {
        "group head",
        "department head",
        "division head",
        "regional manager",
        "manager",
        "team lead",
        "pimpinan proyek",
        "deputy",
    }
    non_structural_phrases = {
        "officer",
        "auditor",
        "analyst",
        "specialist",
        "staff",
        "expert",
    }
    has_structural = any(_has_phrase(normalized, phrase) for phrase in structural_phrases)
    has_non_structural = any(_has_phrase(normalized, phrase) for phrase in non_structural_phrases)
    if has_structural and not has_non_structural:
        return ScopeInference(STRUCTURAL, normalized, "Worksheet title contains structural role signal only.")
    if has_non_structural and not has_structural:
        return ScopeInference(NON_STRUCTURAL, normalized, "Worksheet title contains non-structural role signal only.")
    if has_structural and has_non_structural:
        return ScopeInference(SCOPE_UNCERTAIN, normalized, "Worksheet title contains both structural and non-structural signals.")
    return ScopeInference(SCOPE_UNCERTAIN, normalized, "Worksheet title has no clear structural/non-structural role signal.")


def _truthy(value: Any) -> bool:
    return value in (True, 1, "1", "true", "TRUE", "Y", "y")


def _int_value(value: Any) -> int:
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


def _list_value(row: dict[str, Any], *keys: str) -> tuple[str, ...]:
    for key in keys:
        value = row.get(key)
        if isinstance(value, list):
            return tuple(str(item).strip() for item in value if str(item).strip())
        if isinstance(value, tuple):
            return tuple(str(item).strip() for item in value if str(item).strip())
        if isinstance(value, str) and value.strip():
            return tuple(part.strip() for part in re.split(r"\s*;\s*|\s*,\s*", value) if part.strip())
    return ()


def _is_active_row(row: dict[str, Any]) -> bool:
    for key in [
        "is_company_active",
        "is_group_active",
        "is_position_active",
        "is_position_organization_active",
    ]:
        if row.get(key) not in (None, "") and not _truthy(row.get(key)):
            return False
    if "active_variant_count" in row and _int_value(row.get("active_variant_count")) < 1:
        return False
    if "active_employee_count" in row and _int_value(row.get("active_employee_count")) < 1:
        return False
    return True


def _target_company_matches(row: dict[str, Any], target_company_id: str | None) -> bool:
    return not target_company_id or str(row.get("company_id") or "") == str(target_company_id)


def _candidate_keys(row: dict[str, Any], *values: Any) -> tuple[str, ...]:
    keys: list[str] = []
    for value in values:
        key = normalize_position_lookup(value)
        if key and key not in keys:
            keys.append(key)
    for key in row.get("normalized_lookup_keys") or row.get("lookup_keys") or []:
        normalized = normalize_position_lookup(key)
        if normalized and normalized not in keys:
            keys.append(normalized)
    return tuple(keys)


def _candidate_from_structural(row: dict[str, Any], group_ancestor_names: tuple[str, ...] = ()) -> LookupCandidate | None:
    pmid = row.get("position_master_id")
    title = norm_text(row.get("position_name") or row.get("portaverse_position_title"))
    if pmid in (None, "", 0, "0") or not title:
        return None
    if str(row.get("position_master_type_id") or "") not in STRUCTURAL_TYPE_IDS:
        return None
    if not _is_active_row(row):
        return None
    return LookupCandidate(
        scope=STRUCTURAL,
        position_master_id=str(pmid),
        position_nomenclature_id=None,
        title=title,
        group_name=norm_text(row.get("group_name") or row.get("active_group_name")),
        company_id=norm_text(row.get("company_id")),
        company_name=norm_text(row.get("company_name") or row.get("active_company_name")),
        company_code=norm_text(row.get("company_code") or row.get("active_company_code")),
        active_variant_count=_int_value(row.get("active_variant_count")) or 1,
        active_employee_count=_int_value(row.get("active_employee_count")) or 1,
        definitive_employee_count=_int_value(row.get("definitive_employee_count")),
        secondary_employee_count=_int_value(row.get("secondary_employee_count")),
        active_employee_names=_list_value(row, "active_employee_names", "active_employee_name", "employee_names"),
        active_employee_nipps=_list_value(row, "active_employee_nipps", "active_employee_nipp", "employee_nipps"),
        group_ancestor_names=group_ancestor_names,
        lookup_keys=_candidate_keys(row, title),
        source_row=row,
    )


def _candidate_from_non_structural(row: dict[str, Any], group_ancestor_names: tuple[str, ...] = ()) -> LookupCandidate | None:
    pnid = row.get("cluster_id") or row.get("position_nomenclature_id")
    title = norm_text(row.get("cluster_label") or row.get("position_name") or row.get("portaverse_position_title"))
    if pnid in (None, "", 0, "0") or not title:
        return None
    if str(row.get("position_master_type_id") or "") in STRUCTURAL_TYPE_IDS:
        return None
    if not _is_active_row(row):
        return None
    return LookupCandidate(
        scope=NON_STRUCTURAL,
        position_master_id=None,
        position_nomenclature_id=str(pnid),
        title=title,
        group_name=norm_text(row.get("group_name") or row.get("active_group_name")),
        company_id=norm_text(row.get("company_id")),
        company_name=norm_text(row.get("company_name") or row.get("active_company_name")),
        company_code=norm_text(row.get("company_code") or row.get("active_company_code")),
        active_variant_count=_int_value(row.get("active_variant_count")) or 1,
        active_employee_count=_int_value(row.get("active_employee_count")) or 1,
        definitive_employee_count=_int_value(row.get("definitive_employee_count")),
        secondary_employee_count=_int_value(row.get("secondary_employee_count")),
        active_employee_names=_list_value(row, "active_employee_names", "active_employee_name", "employee_names"),
        active_employee_nipps=_list_value(row, "active_employee_nipps", "active_employee_nipp", "employee_nipps"),
        group_ancestor_names=group_ancestor_names,
        lookup_keys=_candidate_keys(row, title, row.get("position_name")),
        source_row=row,
    )


def _organization_ancestor_names(data: dict[str, Any]) -> dict[str, tuple[str, ...]]:
    organization_rows = data.get("organization_rows", [])
    by_id = {
        str(row.get("group_master_id")): row
        for row in organization_rows
        if isinstance(row, dict) and row.get("group_master_id") not in (None, "")
    }
    cache: dict[str, tuple[str, ...]] = {}

    def ancestors(group_master_id: Any) -> tuple[str, ...]:
        key = str(group_master_id or "")
        if not key:
            return ()
        if key in cache:
            return cache[key]
        names: list[str] = []
        seen: set[str] = set()
        current = by_id.get(key)
        while current:
            parent_id = current.get("parent_id")
            parent_key = str(parent_id or "")
            if not parent_key or parent_key in seen:
                break
            seen.add(parent_key)
            parent = by_id.get(parent_key)
            if not parent:
                break
            name = norm_text(parent.get("group_name"))
            if name and name not in names:
                names.append(name)
            current = parent
        cache[key] = tuple(names)
        return cache[key]

    return {key: ancestors(key) for key in by_id}


def build_lookup_indexes(payload: dict[str, Any] | list[dict[str, Any]], target_company_id: str | None = None) -> LookupIndexes:
    if isinstance(payload, list):
        data: dict[str, Any] = {"rows": payload}
    else:
        data = payload
    structural_rows = data.get("structural_lookup_rows")
    non_structural_rows = data.get("non_structural_lookup_rows")
    if structural_rows is None:
        structural_rows = data.get("position_master_rows", [])
    if non_structural_rows is None:
        non_structural_rows = data.get("rows", [])
    ancestor_names_by_group = _organization_ancestor_names(data)

    structural: list[LookupCandidate] = []
    for row in structural_rows:
        if not isinstance(row, dict) or not _target_company_matches(row, target_company_id):
            continue
        candidate = _candidate_from_structural(
            row,
            ancestor_names_by_group.get(str(row.get("group_master_id") or ""), ()),
        )
        if candidate:
            structural.append(candidate)

    non_structural: list[LookupCandidate] = []
    for row in non_structural_rows:
        if not isinstance(row, dict) or not _target_company_matches(row, target_company_id):
            continue
        candidate = _candidate_from_non_structural(
            row,
            ancestor_names_by_group.get(str(row.get("group_master_id") or ""), ()),
        )
        if candidate:
            non_structural.append(candidate)

    return LookupIndexes(
        structural=structural,
        non_structural=non_structural,
        structural_by_pmid={candidate.position_master_id: candidate for candidate in structural if candidate.position_master_id},
        non_structural_by_pnid={
            candidate.position_nomenclature_id: candidate
            for candidate in non_structural
            if candidate.position_nomenclature_id
        },
        source=data.get("source", {}) if isinstance(data.get("source"), dict) else {},
    )


def _context_score(source: str | None, candidate: str | None) -> float:
    source_key = normalize_position_lookup(source)
    candidate_key = normalize_position_lookup(candidate)
    if not source_key:
        return 1.0
    if not candidate_key:
        return 0.0
    if source_key == candidate_key:
        return 1.0
    if source_key in candidate_key or candidate_key in source_key:
        return 0.85
    source_tokens = set(source_key.split())
    candidate_tokens = set(candidate_key.split())
    overlap = len(source_tokens & candidate_tokens) / max(len(source_tokens), 1)
    return min(overlap, 0.75)


def _title_score(source_title: str, candidate: LookupCandidate) -> float:
    source_key = normalize_position_lookup(source_title)
    if not source_key:
        return 0.0
    source_numbers = set(re.findall(r"\b\d+\b", source_key))
    scores: list[float] = []
    for key in candidate.lookup_keys or (normalize_position_lookup(candidate.title),):
        if source_key == key:
            score = 1.0
        elif source_key in key or key in source_key:
            shorter = min(len(source_key), len(key))
            score = 0.88 if shorter >= 12 else 0.80 if shorter >= 8 else 0.74
        else:
            ratio = SequenceMatcher(None, source_key, key).ratio()
            token_overlap = len(set(source_key.split()) & set(key.split())) / max(len(set(source_key.split())), 1)
            score = max(ratio * 0.8, token_overlap * 0.82)
        if source_numbers:
            key_numbers = set(re.findall(r"\b\d+\b", key))
            if source_numbers & key_numbers:
                score = min(1.0, score + 0.08)
            else:
                score *= 0.75
        scores.append(score)
    return max(scores or [0.0])


def _score_candidate(
    worksheet_title: str,
    group_name: str | None,
    company_hint: str | None,
    candidate: LookupCandidate,
) -> ScoredCandidate:
    title = _title_score(worksheet_title, candidate)
    group_scores = [_context_score(group_name, candidate.group_name)]
    group_scores.extend(_context_score(group_name, ancestor) for ancestor in candidate.group_ancestor_names)
    group = max(group_scores or [0.0])
    company = _context_score(company_hint, candidate.company_name or candidate.company_code)
    score = round((title * 0.72) + (group * 0.20) + (company * 0.08), 4)
    return ScoredCandidate(candidate=candidate, score=score, reason=f"title={title:.2f}; group={group:.2f}; company={company:.2f}")


def _candidate_list(indexes: LookupIndexes, scope: str) -> list[LookupCandidate]:
    if scope == STRUCTURAL:
        return indexes.structural
    if scope == NON_STRUCTURAL:
        return indexes.non_structural
    return []


def _identity(candidate: LookupCandidate) -> tuple[str | None, str | None, str]:
    return (candidate.position_master_id, candidate.position_nomenclature_id, candidate.scope)


def _unique_in_order(values: list[str]) -> tuple[str, ...]:
    seen: set[str] = set()
    unique: list[str] = []
    for value in values:
        text = str(value).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        unique.append(text)
    return tuple(unique)


def _merge_candidate_identity(best: ScoredCandidate, duplicates: list[ScoredCandidate]) -> ScoredCandidate:
    candidates = [best.candidate, *[item.candidate for item in duplicates]]
    employee_names = _unique_in_order(
        [name for candidate in candidates for name in candidate.active_employee_names]
    )
    employee_nipps = _unique_in_order(
        [nipp for candidate in candidates for nipp in candidate.active_employee_nipps]
    )
    variant_keys = {
        str(candidate.source_row.get("position_master_id") or candidate.position_master_id or candidate.position_nomenclature_id)
        for candidate in candidates
    }
    merged = replace(
        best.candidate,
        active_variant_count=max(len(variant_keys), best.candidate.active_variant_count),
        active_employee_count=max(len(employee_nipps), best.candidate.active_employee_count),
        definitive_employee_count=sum(candidate.definitive_employee_count for candidate in candidates),
        secondary_employee_count=sum(candidate.secondary_employee_count for candidate in candidates),
        active_employee_names=employee_names,
        active_employee_nipps=employee_nipps,
    )
    return ScoredCandidate(candidate=merged, score=best.score, reason=best.reason)


def resolve_mapping(
    worksheet: str,
    worksheet_title: str,
    group_name: str | None,
    source_workbook: str | None,
    indexes: LookupIndexes,
    company_hint: str | None = None,
) -> MappingResolution:
    scope = infer_worksheet_scope(worksheet_title)
    if scope.scope == SCOPE_UNCERTAIN:
        return MappingResolution(
            source_workbook=source_workbook,
            worksheet=worksheet,
            raw_worksheet_title=worksheet_title,
            normalized_worksheet_title=scope.normalized_title,
            inferred_scope=SCOPE_UNCERTAIN,
            confidence_label=SCOPE_UNCERTAIN,
            confidence_reason=scope.reason,
        )

    raw_scored = [
        _score_candidate(worksheet_title, group_name, company_hint, candidate)
        for candidate in _candidate_list(indexes, scope.scope)
    ]
    scored_by_identity: dict[tuple[str | None, str | None, str], ScoredCandidate] = {}
    duplicate_scores_by_identity: dict[tuple[str | None, str | None, str], list[ScoredCandidate]] = {}
    for item in sorted(raw_scored, key=lambda candidate_score: candidate_score.score, reverse=True):
        identity = _identity(item.candidate)
        if identity in scored_by_identity:
            duplicate_scores_by_identity.setdefault(identity, []).append(item)
            continue
        scored_by_identity[identity] = item
    for identity, best_item in list(scored_by_identity.items()):
        duplicates = duplicate_scores_by_identity.get(identity, [])
        if duplicates:
            scored_by_identity[identity] = _merge_candidate_identity(best_item, duplicates)
    scored = list(scored_by_identity.values())
    best = scored[0] if scored else None
    runner_up = scored[1] if len(scored) > 1 else None
    if best is None or best.score < 0.65:
        return MappingResolution(
            source_workbook=source_workbook,
            worksheet=worksheet,
            raw_worksheet_title=worksheet_title,
            normalized_worksheet_title=scope.normalized_title,
            inferred_scope=scope.scope,
            confidence_label=NO_CANDIDATE,
            confidence_reason="No active candidate scored above the minimum threshold in the matching lookup.",
        )

    duplicate_strong = bool(
        runner_up
        and runner_up.score >= 0.80
        and best.score - runner_up.score < 0.15
        and _identity(best.candidate) != _identity(runner_up.candidate)
    )
    if duplicate_strong:
        label = MAPPING_CONFLICT
        reason = "Duplicate strong candidates found in the matching active lookup."
    else:
        high = (
            best.score >= 0.90
            and (runner_up is None or best.score - runner_up.score >= 0.15)
            and best.candidate.active_variant_count >= 1
            and best.candidate.active_employee_count >= 1
        )
        label = HIGH_CONFIDENCE if high else LOW_CONFIDENCE
        reason = "Strict high-confidence checks passed." if high else "Candidate exists, but at least one strict condition is weak."

    candidate = best.candidate
    return MappingResolution(
        source_workbook=source_workbook,
        worksheet=worksheet,
        raw_worksheet_title=worksheet_title,
        normalized_worksheet_title=scope.normalized_title,
        inferred_scope=scope.scope,
        confidence_label=label,
        confidence_reason=f"{reason} {best.reason}",
        position_master_id=candidate.position_master_id if scope.scope == STRUCTURAL else None,
        position_nomenclature_id=candidate.position_nomenclature_id if scope.scope == NON_STRUCTURAL else None,
        candidate_title=candidate.title,
        candidate_group=candidate.group_name,
        candidate_company=candidate.company_name,
        candidate_company_code=candidate.company_code,
        candidate_score=best.score,
        runner_up_scope=runner_up.candidate.scope if runner_up else None,
        runner_up_position_master_id=runner_up.candidate.position_master_id if runner_up else None,
        runner_up_position_nomenclature_id=runner_up.candidate.position_nomenclature_id if runner_up else None,
        runner_up_title=runner_up.candidate.title if runner_up else None,
        runner_up_score=runner_up.score if runner_up else None,
        active_variant_count=candidate.active_variant_count,
        active_employee_count=candidate.active_employee_count,
        active_employee_name="; ".join(candidate.active_employee_names) or None,
        active_employee_nipp="; ".join(candidate.active_employee_nipps) or None,
        upload_allowed=label == HIGH_CONFIDENCE,
    )


def recommended_action(label: str) -> str:
    return {
        HIGH_CONFIDENCE: "No action required; auto-mapped.",
        LOW_CONFIDENCE: "Review candidate before allowing upload rows.",
        SCOPE_UNCERTAIN: "Decide whether worksheet is structural or non-structural.",
        NO_CANDIDATE: "Check active reference or source worksheet title.",
        MAPPING_CONFLICT: "Choose one candidate or create manual override.",
    }.get(label, "Review mapping before allowing upload rows.")


def mapping_report_row(result: MappingResolution) -> dict[str, Any]:
    row = {
        "Source Workbook": result.source_workbook,
        "Worksheet": result.worksheet,
        "Raw Worksheet Title": result.raw_worksheet_title,
        "Normalized Worksheet Title": result.normalized_worksheet_title,
        "Inferred Scope": result.inferred_scope,
        "Confidence Label": result.confidence_label,
        "Confidence Reason": result.confidence_reason,
        "Candidate PMID": result.position_master_id,
        "Candidate PNID": result.position_nomenclature_id,
        "Candidate Title": result.candidate_title,
        "Candidate Group": result.candidate_group,
        "Candidate Company": result.candidate_company,
        "Candidate Company Code": result.candidate_company_code,
        "Candidate Score": result.candidate_score,
        "Runner-up Scope": result.runner_up_scope,
        "Runner-up PMID": result.runner_up_position_master_id,
        "Runner-up PNID": result.runner_up_position_nomenclature_id,
        "Runner-up Title": result.runner_up_title,
        "Runner-up Score": result.runner_up_score,
        "Active Variant Count": result.active_variant_count,
        "Active Employee Count": result.active_employee_count,
        "Active Employee Name": result.active_employee_name,
        "Active Employee NIPP": result.active_employee_nipp,
        "Recommended Action": recommended_action(result.confidence_label),
    }
    return {column: row.get(column) for column in REPORT_COLUMNS}


def validate_manual_override(
    inferred_scope: str,
    position_master_id: str | None,
    position_nomenclature_id: str | None,
    indexes: LookupIndexes,
) -> OverrideValidation:
    pmid = str(position_master_id) if position_master_id not in (None, "", 0, "0") else None
    pnid = str(position_nomenclature_id) if position_nomenclature_id not in (None, "", 0, "0") else None
    if inferred_scope == STRUCTURAL:
        if not pmid or pnid:
            return OverrideValidation(False, "Manual structural override must provide PMID only.")
        candidate = indexes.structural_by_pmid.get(pmid)
        if not candidate:
            return OverrideValidation(False, "Manual structural PMID is not present in active structural lookup.")
        return OverrideValidation(True, "Manual structural PMID validated against active structural lookup.", candidate)
    if inferred_scope == NON_STRUCTURAL:
        if not pnid or pmid:
            return OverrideValidation(False, "Manual non-structural override must provide PNID only.")
        candidate = indexes.non_structural_by_pnid.get(pnid)
        if not candidate:
            return OverrideValidation(False, "Manual non-structural PNID is not present in active non-structural lookup.")
        return OverrideValidation(True, "Manual non-structural PNID validated against active non-structural lookup.", candidate)
    return OverrideValidation(False, "Manual override requires reviewer-selected structural or non-structural scope.")
