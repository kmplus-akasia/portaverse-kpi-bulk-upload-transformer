# Domain docs

How engineering skills should consume this repo's domain documentation when exploring the codebase.

## Before exploring, read these

- `CONTEXT.md` at the repo root, when present.
- `CONTEXT-MAP.md` at the repo root, when present; it points at one `CONTEXT.md` per context.
- Relevant ADRs under `docs/adr/`.
- Relevant converter specs and diagrams under `docs/converter/` and `docs/superpowers/specs/`.

If any of these files do not exist, proceed silently. Do not create a context document or ADR until a domain term or architectural decision needs to be recorded.

## Layout

This is a single-context repo:

```text
/
├── CONTEXT.md
├── docs/adr/
├── docs/converter/
├── docs/superpowers/specs/
└── scripts/
```

## Vocabulary and conflicts

Use established domain terms such as PMID, PNID, PMVID, structural, non-structural, active reference, review artifact, and upload-ready. If an output conflicts with an ADR or an approved converter spec, surface that conflict explicitly rather than silently overriding it.
