# Diagram Logic Ideal KPI Converter

Dokumen ini memvisualisasikan logic ideal converter KPI bulk upload, terutama pemetaan posisi dan penulisan kamus KPI ke template upload.

## 1. Alur Besar Converter

```mermaid
flowchart TD
    A["Input Kamus KPI<br/>.xlsx atau .zip"] --> B["Baca template upload resmi<br/>24 kolom"]
    B --> C["Baca production position reference<br/>configs/production_position_reference.json"]
    C --> D["Load atau discover config posisi<br/>per workbook + per sheet"]
    D --> E["Resolve identity posisi<br/>sheet name, Nama Posisi, group, alias"]
    E --> F["Validasi mapping posisi<br/>PMID vs PNID"]
    F --> G{Mapping clear?}
    G -- "Tidak" --> H["Set mapping_conflict<br/>skip sheet + tulis error report"]
    G -- "Ya" --> I["Parse blok KPI<br/>Impact, Output, KAI"]
    I --> J["Normalisasi enum dan field wajib<br/>period, polarity, cascading, ownership, nature"]
    J --> K["Tulis rows ke KPI Template"]
    K --> L["Final validation gate<br/>schema, enum, PMID/PNID, weight, report"]
    L --> M{Ada error?}
    M -- "Ya" --> N["Workbook inspeksi boleh ada<br/>status belum siap upload"]
    M -- "Tidak" --> O["Upload-ready workbook<br/>manifest + report"]
```

## 2. Logic Pemetaan Posisi

```mermaid
flowchart TD
    A["Mulai dari satu worksheet"] --> B["Ambil kandidat identity<br/>position_name, sheet_name, group_name, lookup_names"]
    B --> C["Normalisasi teks<br/>contoh DH -> Department Head, SPV -> Supervisor"]
    C --> D["Cari di production reference<br/>position_master_rows + rows"]
    D --> E{Ada satu kandidat jelas?}

    E -- "Tidak ada / banyak kandidat" --> X["mapping_conflict<br/>PMID kosong, PNID kosong, sheet diskip"]

    E -- "Ya" --> F{Production scope?}

    F -- "position_master_type_id == 5" --> G["Structural"]
    G --> H["Gunakan position_master_id sebagai PMID"]
    H --> I["Output:<br/>Position Master ID terisi<br/>Position Nomenklatur ID kosong"]

    F -- "type selain 5 dan punya cluster_id" --> J["Non-structural"]
    J --> K["Gunakan rows.cluster_id sebagai PNID"]
    K --> L["Output:<br/>Position Nomenklatur ID terisi<br/>Position Master ID kosong"]

    F -- "Tidak cukup data scope" --> X

    subgraph Rules["Aturan anti-salah mapping"]
        R1["PNID adalah rows.cluster_id<br/>bukan mapping row id dan bukan internal PMID"]
        R2["Angka yang sama bisa muncul sebagai PMID dan PNID"]
        R3["Jangan pilih ID dari numeric collision"]
        R4["Pilih berdasarkan resolved identity + production scope"]
        R5["Jika ragu, fail closed sebagai mapping_conflict"]
    end

    I --> V["Validasi output row"]
    L --> V
    V --> W{Row punya tepat satu identity?}
    W -- "PMID dan PNID dua-duanya terisi" --> Y["Error: double identity"]
    W -- "Dua-duanya kosong" --> Z["Error: blank identity"]
    W -- "Tepat satu terisi" --> AA["Mapping valid untuk ditulis"]
```

## 3. Logic Penulisan Kamus ke Template Upload

```mermaid
flowchart TD
    A["Mapping posisi valid"] --> B["Baca sheet Kamus KPI"]
    B --> C["Temukan header block<br/>KPI Impact, KPI Output, KAI"]
    C --> D["Parse KPI Impact"]
    D --> E["Parse child KPI Output"]
    E --> F["Parse child KAI"]

    D --> G["Backfill field yang kosong<br/>dari shared KPI title jika aman"]
    E --> G
    F --> G

    G --> H["Normalisasi field upload"]
    H --> H1["Period<br/>Triwulan -> TRIWULANAN"]
    H --> H2["Polarity<br/>Positif -> POSITIVE"]
    H --> H3["Cascading<br/>Direct/Indirect/Duplicate"]
    H --> H4["Ownership Type<br/>Specific/Shared/Common"]
    H --> H5["Nature of Work KAI<br/>Routine / Non Routine"]

    H1 --> I["Build row IMPACT"]
    H2 --> I
    H3 --> I
    H4 --> I
    H5 --> I
    I --> J["Build row OUTPUT dengan parent IMPACT"]
    J --> K["Build row KAI dengan parent OUTPUT"]

    K --> L["Isi identity posisi ke setiap row"]
    L --> M{Scope posisi}
    M -- "Structural" --> N["Set PMID<br/>PNID blank"]
    M -- "Non-structural" --> O["Set PNID<br/>PMID blank"]

    N --> P["Validasi row"]
    O --> P
    P --> Q{Valid?}
    Q -- "Tidak" --> R["Tulis report error/warning<br/>jangan tandai upload-ready"]
    Q -- "Ya" --> S["Tulis ke sheet KPI Template"]
```

## 4. Mental Model Singkat

```mermaid
flowchart LR
    A["Structural position"] --> B["PMID"]
    B --> C["Satu jabatan struktural spesifik"]

    D["Non-structural position"] --> E["PNID = cluster_id"]
    E --> F["Importer backend expand ke PMID terkait"]

    G["Unclear mapping"] --> H["mapping_conflict"]
    H --> I["Skip sheet, review manual"]
```

## Prinsip Desain

1. `configs/production_position_reference.json` adalah source of truth offline.
2. `position_master_type_id == 5` berarti structural.
3. PNID selalu `rows[].cluster_id`.
4. Structural menulis PMID saja.
5. Non-structural menulis PNID saja.
6. Converter tidak boleh menulis PMID dan PNID sekaligus.
7. Converter tidak boleh menulis row tanpa PMID/PNID.
8. Mapping yang ambigu harus menjadi `mapping_conflict`, bukan ditebak.
