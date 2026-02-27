# RKS v2.0 — Python AI Agent + Web UI (MVP)

This workspace contains a minimal, working implementation of the **Reflexive Knowledge System (RKS v2.0)** described in your spec.

## What this MVP implements

- **Node schema**: definition, causal chain, boundary condition, failure mode, links, assumptions, status
- **Contradiction registry**: stores conditional contradictions (Node A vs Node B)
- **Hybrid contract gate**:
  - Level 1 (AI): create node skeleton
  - Level 2 (Human): fill causal/boundary/failure + real examples
  - Level 3 (AI validation): recompute **NMS** + enforce gates
- **NMS scoring (0–5)** per the spec
- **Explore/Build/Active** status:
  - If Level 2 is incomplete, selecting **Active** is automatically downgraded to **Build**
- **Reflection logging**: prediction/outcome/delta run logs
- **File-based persistence**: JSONL in `data/`

## Run the Web UI

From the workspace root:

```powershell
F:/00AiAgent/.venv/Scripts/python.exe -m pip install -r requirements.txt
F:/00AiAgent/.venv/Scripts/python.exe -m uvicorn webapp.main:app --reload --port 8000
```

Open:

- http://127.0.0.1:8000/

## Data files

- `data/nodes.jsonl` (append-only, latest record wins)
- `data/contradictions.jsonl`
- `data/runs.jsonl`

## Dev note

This is a **single-user MVP** meant to prove the RKS loop end-to-end.
Next upgrades (optional): add auth/multi-user, add real LLM integration for ontology tagging/mechanism suggestion, and add governance policies (entropy budget, ontology freeze threshold) as automated checks.
