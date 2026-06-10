# Portaverse KPI Converter Hardening Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make `scripts/kpi_bulk_transform.py` conform to `docs/superpowers/specs/2026-06-10-portaverse-kpi-converter-design.md` so Group 2/3 conversions are enum-safe, PMID/PNID-safe, and auditable.

**Architecture:** Keep the current single converter entrypoint, but add small internal units: enum normalization, position-scope validation, final upload gate, and richer reporting. Do not rewrite parser/output architecture.

**Tech Stack:** Python 3.11+, `openpyxl`, `unittest`, existing CSV report flow.

---

## File Map

- Modify `scripts/kpi_bulk_transform.py`: enum normalizers, report issue fields, mapping validator, output validation gate.
- Modify `tests/test_kpi_bulk_transform.py`: unit tests for enum pollution, nature logic, PMID/PNID scope correction, final validation.
- Modify `scripts/build_conversion_recap.py`: add enum/mapping issue counters and readiness fields.
- Modify `README.md`: add non-Codex runbook commands for zip batch, single workbook, and recap.

## Task 1: Enum Normalizer Tests

**Files:**
- Modify: `tests/test_kpi_bulk_transform.py`

- [ ] **Step 1: Add imports for new normalizer functions**

```python
from kpi_bulk_transform import (  # noqa: E402
    NormalizationStatus,
    normalize_cascading,
    normalize_kai_nature,
    normalize_ownership_type,
    normalize_period,
    normalize_polarity,
)
```

- [ ] **Step 2: Add failing enum tests**

Append these test methods to `KpiBulkTransformTest`:

```python
def test_period_normalizes_raw_variants_and_flags_ambiguous_combo(self):
    self.assertEqual(normalize_period("per tahun").value, "TAHUNAN")
    self.assertEqual(normalize_period("Semesteran").value, "SEMESTER")
    combo = normalize_period("Triwulanan/Tahunan")
    self.assertEqual(combo.status, NormalizationStatus.AMBIGUOUS)
    self.assertIsNone(combo.value)

def test_polarity_pollution_defaults_positive(self):
    result = normalize_polarity("INDIRECT")
    self.assertEqual(result.value, "POSITIVE")
    self.assertEqual(result.status, NormalizationStatus.CROSS_COLUMN)
    self.assertEqual(normalize_polarity("Negatif").value, "NEGATIVE")

def test_cascading_pollution_defaults_indirect(self):
    result = normalize_cascading("SPECIFIC")
    self.assertEqual(result.value, "INDIRECT")
    self.assertEqual(result.status, NormalizationStatus.CROSS_COLUMN)
    self.assertEqual(normalize_cascading("Indirect").value, "INDIRECT")

def test_ownership_pollution_defaults_specific(self):
    result = normalize_ownership_type("Non Routine")
    self.assertEqual(result.value, "SPECIFIC")
    self.assertEqual(result.status, NormalizationStatus.CROSS_COLUMN)
    self.assertEqual(normalize_ownership_type("SPESIFIC").value, "SPECIFIC")

def test_kai_nature_infers_from_period(self):
    self.assertEqual(normalize_kai_nature(None, "TAHUNAN").value, "Non Routine")
    self.assertEqual(normalize_kai_nature(None, "TRIWULANAN").value, "Routine")
    self.assertEqual(normalize_kai_nature("INDIRECT", "BULANAN").value, "Routine")
    self.assertEqual(normalize_kai_nature("Non-Rotine", "BULANAN").value, "Non Routine")
```

- [ ] **Step 3: Run tests and confirm failure**

Run:

```bash
python3 -m unittest tests/test_kpi_bulk_transform.py
```

Expected: FAIL with missing imports/functions.

## Task 2: Implement Enum Normalizers

**Files:**
- Modify: `scripts/kpi_bulk_transform.py`

- [ ] **Step 1: Add normalization types and constants near existing enum helpers**

Insert after `ALLOWED_UPLOAD_POLARITIES`:

```python
from enum import Enum

class NormalizationStatus(str, Enum):
    OK = "ok"
    NORMALIZED = "normalized"
    DEFAULTED = "defaulted"
    CROSS_COLUMN = "cross_column"
    AMBIGUOUS = "ambiguous"
    INVALID = "invalid"

@dataclass
class NormalizedEnum:
    value: str | None
    status: NormalizationStatus
    raw_value: str | None
    message: str

ALLOWED_UPLOAD_PERIODS = {"BULANAN", "TRIWULANAN", "TAHUNAN", "SEMESTER", "MONTHLY", "QUARTERLY", "WEEKLY"}
ALLOWED_UPLOAD_CASCADING = {"DIRECT", "INDIRECT", "DUPLICATE"}
ALLOWED_UPLOAD_OWNERSHIP = {"SPECIFIC", "SHARED", "COMMON"}
CROSS_COLUMN_ENUM_VALUES = ALLOWED_UPLOAD_POLARITIES | ALLOWED_UPLOAD_PERIODS | ALLOWED_UPLOAD_CASCADING | ALLOWED_UPLOAD_OWNERSHIP | {"Routine", "Non Routine"}
```

- [ ] **Step 2: Add exact normalizer functions**

Place before legacy `uploader_period`:

```python
def normalized_key(value: str | None) -> str:
    return normalize_title(value).replace(" ", "_")

def enum_result(value: str | None, status: NormalizationStatus, raw: str | None, message: str) -> NormalizedEnum:
    return NormalizedEnum(value=value, status=status, raw_value=raw, message=message)

def normalize_period(value: str | None, fallback: str | None = None) -> NormalizedEnum:
    raw = norm_text(value)
    if not raw or is_placeholder(raw):
        return enum_result(fallback, NormalizationStatus.DEFAULTED if fallback else NormalizationStatus.INVALID, raw, "Period missing.")
    key = normalized_key(raw)
    mapping = {
        "triwulan": "TRIWULANAN",
        "triwulanan": "TRIWULANAN",
        "tahunan": "TAHUNAN",
        "tahun": "TAHUNAN",
        "per_tahun": "TAHUNAN",
        "per_tahunan": "TAHUNAN",
        "semester": "SEMESTER",
        "semesteran": "SEMESTER",
        "per_semester": "SEMESTER",
        "bulanan": "BULANAN",
        "monthly": "MONTHLY",
        "quarterly": "QUARTERLY",
        "weekly": "WEEKLY",
    }
    if "/" in raw and fallback:
        return enum_result(fallback, NormalizationStatus.AMBIGUOUS, raw, f"Ambiguous period defaulted to parent period {fallback}.")
    if "/" in raw:
        return enum_result(None, NormalizationStatus.AMBIGUOUS, raw, "Ambiguous period requires review.")
    if key in mapping:
        status = NormalizationStatus.OK if raw == mapping[key] else NormalizationStatus.NORMALIZED
        return enum_result(mapping[key], status, raw, f"Period normalized to {mapping[key]}.")
    upper = raw.upper()
    if upper in ALLOWED_UPLOAD_PERIODS:
        return enum_result(upper, NormalizationStatus.OK, raw, "Period already valid.")
    return enum_result(fallback, NormalizationStatus.INVALID, raw, "Invalid period.")

def normalize_polarity(value: str | None) -> NormalizedEnum:
    raw = norm_text(value)
    if not raw or is_placeholder(raw):
        return enum_result("POSITIVE", NormalizationStatus.DEFAULTED, raw, "Polarity defaulted to POSITIVE.")
    key = normalized_key(raw)
    mapping = {"positif": "POSITIVE", "positive": "POSITIVE", "pos": "POSITIVE", "negatif": "NEGATIVE", "negative": "NEGATIVE", "neg": "NEGATIVE", "netral": "NEUTRAL", "neutral": "NEUTRAL"}
    if key in mapping:
        return enum_result(mapping[key], NormalizationStatus.NORMALIZED, raw, f"Polarity normalized to {mapping[key]}.")
    if raw.upper() in ALLOWED_UPLOAD_POLARITIES:
        return enum_result(raw.upper(), NormalizationStatus.OK, raw, "Polarity already valid.")
    status = NormalizationStatus.CROSS_COLUMN if raw.upper() in CROSS_COLUMN_ENUM_VALUES or raw in CROSS_COLUMN_ENUM_VALUES else NormalizationStatus.DEFAULTED
    return enum_result("POSITIVE", status, raw, "Invalid polarity defaulted to POSITIVE.")

def normalize_cascading(value: str | None) -> NormalizedEnum:
    raw = norm_text(value)
    if not raw or is_placeholder(raw):
        return enum_result(None, NormalizationStatus.OK, raw, "Cascading blank.")
    upper = raw.upper()
    if upper in ALLOWED_UPLOAD_CASCADING:
        return enum_result(upper, NormalizationStatus.NORMALIZED if raw != upper else NormalizationStatus.OK, raw, f"Cascading normalized to {upper}.")
    status = NormalizationStatus.CROSS_COLUMN if upper in CROSS_COLUMN_ENUM_VALUES or raw in CROSS_COLUMN_ENUM_VALUES else NormalizationStatus.DEFAULTED
    return enum_result("INDIRECT", status, raw, "Invalid cascading defaulted to INDIRECT.")

def normalize_ownership_type(value: str | None) -> NormalizedEnum:
    raw = norm_text(value)
    if not raw or is_placeholder(raw):
        return enum_result("SPECIFIC", NormalizationStatus.DEFAULTED, raw, "Ownership Type defaulted to SPECIFIC.")
    key = normalized_key(raw)
    mapping = {"specific": "SPECIFIC", "spesific": "SPECIFIC", "shared": "SHARED", "common": "COMMON"}
    if key in mapping:
        return enum_result(mapping[key], NormalizationStatus.NORMALIZED if raw != mapping[key] else NormalizationStatus.OK, raw, f"Ownership Type normalized to {mapping[key]}.")
    status = NormalizationStatus.CROSS_COLUMN if raw.upper() in CROSS_COLUMN_ENUM_VALUES or raw in CROSS_COLUMN_ENUM_VALUES else NormalizationStatus.DEFAULTED
    return enum_result("SPECIFIC", status, raw, "Invalid Ownership Type defaulted to SPECIFIC.")

def normalize_kai_nature(value: str | None, period: str | None = None) -> NormalizedEnum:
    raw = norm_text(value)
    normalized_period = normalize_period(period).value
    inferred = "Non Routine" if normalized_period == "TAHUNAN" else "Routine"
    if not raw or is_placeholder(raw):
        return enum_result(inferred, NormalizationStatus.DEFAULTED, raw, f"KAI Nature inferred as {inferred}.")
    key = normalized_key(raw)
    if key in {"routine", "rutin"}:
        return enum_result("Routine", NormalizationStatus.NORMALIZED, raw, "KAI Nature normalized to Routine.")
    if key in {"non_routine", "non_rutin", "non_rutine", "non_rotine", "non_rotin", "non_routinee"}:
        return enum_result("Non Routine", NormalizationStatus.NORMALIZED, raw, "KAI Nature normalized to Non Routine.")
    status = NormalizationStatus.CROSS_COLUMN if raw.upper() in CROSS_COLUMN_ENUM_VALUES or raw in CROSS_COLUMN_ENUM_VALUES or raw.lower().startswith("http") else NormalizationStatus.DEFAULTED
    return enum_result(inferred, status, raw, f"Invalid KAI Nature inferred as {inferred}.")
```

- [ ] **Step 3: Keep legacy wrapper names working**

Replace `uploader_period`, `uploader_polarity`, and `uploader_kai_nature` bodies:

```python
def uploader_period(value: str | None) -> str | None:
    return normalize_period(value).value

def uploader_polarity(value: str | None) -> str | None:
    return normalize_polarity(value).value

def uploader_kai_nature(value: str | None, period: str | None = None) -> str:
    return normalize_kai_nature(value, period).value or "Routine"
```

- [ ] **Step 4: Run enum tests**

Run:

```bash
python3 -m unittest tests.test_kpi_bulk_transform.KpiBulkTransformTest.test_period_normalizes_raw_variants_and_flags_ambiguous_combo tests.test_kpi_bulk_transform.KpiBulkTransformTest.test_cascading_pollution_defaults_indirect
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add scripts/kpi_bulk_transform.py tests/test_kpi_bulk_transform.py
git commit -m "Add KPI upload enum normalizers"
```

## Task 3: Integrate Normalizers Into Output Rows and Reports

**Files:**
- Modify: `scripts/kpi_bulk_transform.py`
- Test: `tests/test_kpi_bulk_transform.py`

- [ ] **Step 1: Add report helper**

Add near `ValidationIssue`:

```python
def append_enum_issue(
    issues: list[ValidationIssue] | None,
    config: PositionConfig,
    source_row: int | None,
    record_type: str,
    title: str | None,
    field_name: str,
    result: NormalizedEnum,
) -> None:
    if issues is None or result.status in {NormalizationStatus.OK, NormalizationStatus.NORMALIZED}:
        return
    severity = "error" if result.status in {NormalizationStatus.AMBIGUOUS, NormalizationStatus.INVALID} and not result.value else "warning"
    issues.append(
        ValidationIssue(
            severity=severity,
            sheet_name=config.sheet_name,
            source_row=source_row,
            record_type=record_type,
            title=title,
            message=(
                f"enum_issue category={result.status.value}; field={field_name}; "
                f"raw={result.raw_value}; normalized={result.value}; {result.message}"
            ),
        )
    )
```

- [ ] **Step 2: Replace direct uploader calls in `build_upload_rows`**

For each IMPACT/OUTPUT/KAI block, compute `period_result`, `polarity_result`, `cascading_result`, `ownership_result`, and `nature_result`, append issues, then write `.value`.

OUTPUT example:

```python
output_period_result = normalize_period(output_period, normalize_period(impact.period).value)
output_polarity_result = normalize_polarity(output.get("polarity"))
output_cascading_result = normalize_cascading(output.get("cascading"))
output_ownership_result = normalize_ownership_type(output.get("ownership_type"))
append_enum_issue(issues, config, output.get("source_row"), "OUTPUT", output.get("title"), "Period", output_period_result)
append_enum_issue(issues, config, output.get("source_row"), "OUTPUT", output.get("title"), "Polarity", output_polarity_result)
append_enum_issue(issues, config, output.get("source_row"), "OUTPUT", output.get("title"), "Cascading", output_cascading_result)
append_enum_issue(issues, config, output.get("source_row"), "OUTPUT", output.get("title"), "Ownership Type", output_ownership_result)
```

Use these values in the row:

```python
output_polarity_result.value
output_period_result.value
output_cascading_result.value
output_ownership_result.value
```

- [ ] **Step 3: Run full unit tests**

```bash
python3 -m unittest tests/test_kpi_bulk_transform.py
```

Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add scripts/kpi_bulk_transform.py tests/test_kpi_bulk_transform.py
git commit -m "Report enum corrections during KPI conversion"
```

## Task 4: Position Scope Validation

**Files:**
- Modify: `scripts/kpi_bulk_transform.py`
- Test: `tests/test_kpi_bulk_transform.py`

- [ ] **Step 1: Add tests for `515` and `67` scope correction**

Add tests that create fake reference payloads where IDs exist in both `rows` and `position_master_rows`, then assert structural title wins:

```python
def test_structural_title_wins_when_same_number_exists_as_pmid_and_pnid(self):
    mapping = {
        "manager rekrutmen karir": {
            "position_master_id": "515",
            "position_nomenclature_id": "515",
            "position_scope": "non_structural",
            "portaverse_position_title": "Manager Rekrutmen dan Karir",
            "position_master_type_id": "5",
        }
    }
    config = PositionConfig(
        sheet_name="Manager Rekrutmen-Karir",
        position_name="Manager Rekrutmen-Karir",
        group_name="Group Pengelolaan SDM",
        directorate_name="Direktorat SDM & Umum",
        position_nomenclature_id="515",
        position_scope="non_structural",
    )
    refresh_configs_from_mapping([config], mapping)
    self.assertEqual(config.position_master_id, "515")
    self.assertIsNone(config.position_nomenclature_id)
    self.assertEqual(config.position_scope, "structural")
```

- [ ] **Step 2: Extend mapping rows with type metadata**

In `load_nomenclature_mapping`, include `position_master_type_id` when creating mapping dicts from `position_master_rows` and `rows`.

- [ ] **Step 3: Update `refresh_config_from_mapping`**

Make scope resolution title-first:

```python
def resolved_scope_from_lookup(lookup: dict[str, str | None]) -> str | None:
    if str(lookup.get("position_master_type_id") or "") == "5":
        return "structural"
    return lookup.get("position_scope")
```

When resolved scope is structural, set `position_master_id` and clear PNID. When non-structural, set PNID and clear PMID.

- [ ] **Step 4: Run targeted mapping tests**

```bash
python3 -m unittest tests.test_kpi_bulk_transform.KpiBulkTransformTest.test_structural_title_wins_when_same_number_exists_as_pmid_and_pnid
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add scripts/kpi_bulk_transform.py tests/test_kpi_bulk_transform.py
git commit -m "Validate KPI position scope before ID output"
```

## Task 5: Final Workbook Gate

**Files:**
- Modify: `scripts/kpi_bulk_transform.py`
- Test: `tests/test_kpi_bulk_transform.py`

- [ ] **Step 1: Extend `validate_output_rows`**

Add allowlist checks for Period, Cascading, Ownership Type, and Nature:

```python
period = norm_text(row_map.get("Period"))
if period and period not in ALLOWED_UPLOAD_PERIODS:
    issues.append(ValidationIssue("error", config.sheet_name, None, record_type or "row", title, f"Invalid Period enum: {period}"))
cascading = norm_text(row_map.get("Cascading"))
if cascading and cascading not in ALLOWED_UPLOAD_CASCADING:
    issues.append(ValidationIssue("error", config.sheet_name, None, record_type or "row", title, f"Invalid Cascading enum: {cascading}"))
ownership = norm_text(row_map.get("Ownership Type"))
if ownership and ownership not in ALLOWED_UPLOAD_OWNERSHIP:
    issues.append(ValidationIssue("error", config.sheet_name, None, record_type or "row", title, f"Invalid Ownership Type enum: {ownership}"))
nature = norm_text(row_map.get("Nature Of Work (KAI Only)"))
if record_type == "KAI" and nature not in {"Routine", "Non Routine"}:
    issues.append(ValidationIssue("error", config.sheet_name, None, "KAI", title, f"Invalid KAI Nature enum: {nature}"))
```

- [ ] **Step 2: Add PMID/PNID invariant checks**

Inside `validate_output_rows`:

```python
pmid = norm_text(row_map.get("Position Master ID (Required)"))
pnid = norm_text(row_map.get("Position Nomenklatur ID"))
if pmid and pnid:
    issues.append(ValidationIssue("error", config.sheet_name, None, record_type or "row", title, "Invalid upload scope: row has both PMID and PNID."))
```

- [ ] **Step 3: Run full tests**

```bash
python3 -m unittest tests/test_kpi_bulk_transform.py
```

Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add scripts/kpi_bulk_transform.py tests/test_kpi_bulk_transform.py
git commit -m "Add final KPI upload validation gate"
```

## Task 6: Recap and README

**Files:**
- Modify: `scripts/build_conversion_recap.py`
- Modify: `README.md`

- [ ] **Step 1: Recap enum issues**

In `load_report_counts`, count messages containing `enum_issue category=`, `mapping_corrected`, and `mapping_conflict`.

```python
if "enum_issue category=" in message:
    counts["enum_issue"] += 1
if "category=cross_column" in message:
    counts["cross_column_enum"] += 1
if "mapping_corrected" in message:
    counts["mapping_corrected"] += 1
if "mapping_conflict" in message:
    counts["mapping_conflict"] += 1
```

- [ ] **Step 2: Add workbook recap columns**

Add these keys to `workbook_rows`:

```python
"Enum Issues": report_counts["enum_issue"],
"Cross Column Enum": report_counts["cross_column_enum"],
"Mapping Corrected": report_counts["mapping_corrected"],
"Mapping Conflict": report_counts["mapping_conflict"],
```

- [ ] **Step 3: Add README runbook**

Append section:

````markdown
## Run Converter Without Codex

From this repo:

```bash
cd /Users/alfredoteja/Documents/portaverse-kpi-bulk-upload-transformer
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
```

Run a ZIP batch:

```bash
python3 scripts/kpi_bulk_transform.py \
  --source "/absolute/path/to/KAMUS KPI GROUP 2.zip" \
  --template "input/KPI Upload Template.xlsx" \
  --mapping "configs/production_position_reference.json" \
  --output-dir "output/group2_conversion_$(date +%Y%m%d_%H%M)"
```

Run one workbook:

```bash
python3 scripts/kpi_bulk_transform.py \
  --source "/absolute/path/to/source.xlsx" \
  --template "input/KPI Upload Template.xlsx" \
  --config "configs/pre_restructure_positions_rw_reviewed_20260609.json" \
  --mapping "configs/production_position_reference.json" \
  --output "output/single_conversion.xlsx" \
  --report "output/single_conversion.report.csv"
```

Build recap:

```bash
python3 scripts/build_conversion_recap.py \
  --output-dir "output/group2_conversion_YYYYMMDD_HHMM" \
  --config "configs/pre_restructure_positions_rw_reviewed_20260609.json" \
  --reference "configs/production_position_reference.json" \
  --output "output/group2_conversion_recap.xlsx" \
  --report-scope "Group 2"
```
````

- [ ] **Step 4: Run tests**

```bash
python3 -m unittest tests/test_kpi_bulk_transform.py
python3 -m py_compile scripts/build_conversion_recap.py scripts/kpi_bulk_transform.py
```

Expected: PASS and no syntax errors.

- [ ] **Step 5: Commit**

```bash
git add README.md scripts/build_conversion_recap.py
git commit -m "Document standalone KPI conversion workflow"
```

## Task 7: Batch Verification

**Files:**
- No code changes unless a test reveals a real defect.

- [ ] **Step 1: Run Group 1 HO verification batch**

```bash
python3 scripts/kpi_bulk_transform.py \
  --source "/Users/alfredoteja/Downloads/KAMUS KPI HO PRE-RESTRUCTURE-20260602T070011Z-3-001.zip" \
  --template "input/KPI Upload Template.xlsx" \
  --mapping "configs/production_position_reference.json" \
  --output-dir "output/group1_ho_hardened_verify"
```

Expected: Generated workbook/report files. Exit code is `0` only if no upload blockers remain.

- [ ] **Step 2: Run Group 3 verification batch**

```bash
python3 scripts/kpi_bulk_transform.py \
  --source "/Users/alfredoteja/Downloads/KAMUS KPI PELINDO GROUP 3 (AFILIASI, NON CLUSTER, DANA PENSIUN)-20260607T162631Z-3-001.zip" \
  --template "input/KPI Upload Template.xlsx" \
  --mapping "configs/production_position_reference.json" \
  --output-dir "output/group3_hardened_verify"
```

Expected: Generated workbook/report files. Remaining errors are mapping/source blockers, not invalid enum pass-through.

- [ ] **Step 3: Build recaps**

```bash
python3 scripts/build_conversion_recap.py \
  --output-dir "output/group1_ho_hardened_verify" \
  --config "configs/pre_restructure_positions_rw_reviewed_20260609.json" \
  --reference "configs/production_position_reference.json" \
  --output "output/group1_ho_hardened_verify_recap.xlsx" \
  --report-scope "Group 1 HO Hardened Verify"
```

- [ ] **Step 4: Commit final passing state**

```bash
git add scripts/kpi_bulk_transform.py scripts/build_conversion_recap.py tests/test_kpi_bulk_transform.py README.md
git commit -m "Harden KPI converter for enum and position scope issues"
```

## Self-Review

- Spec coverage: enum normalization, cross-column spotlight, PMID/PNID scope, final gate, recap, tests, and no-Codex usage are covered.
- Placeholder scan: no unfinished markers or open implementation blanks.
- Type consistency: plan uses `NormalizedEnum`, `NormalizationStatus`, and existing `ValidationIssue`, `PositionConfig`, `build_upload_rows`, `validate_output_rows`, `refresh_configs_from_mapping`.
