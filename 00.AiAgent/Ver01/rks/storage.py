from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from pydantic import TypeAdapter

from .models import Contradiction, Node, RunLog


@dataclass(frozen=True)
class StoragePaths:
    nodes_path: Path
    contradictions_path: Path
    runs_path: Path


class FileStorage:
    """Simple JSONL storage (append-only + in-memory index on load).

    This is intentionally minimal for a single-user MVP.
    """

    def __init__(self, data_dir: Path):
        data_dir.mkdir(parents=True, exist_ok=True)
        self.paths = StoragePaths(
            nodes_path=data_dir / "nodes.jsonl",
            contradictions_path=data_dir / "contradictions.jsonl",
            runs_path=data_dir / "runs.jsonl",
        )

        self._node_adapter = TypeAdapter(Node)
        self._contr_adapter = TypeAdapter(Contradiction)
        self._run_adapter = TypeAdapter(RunLog)

    def _read_jsonl(self, path: Path) -> list[dict]:
        if not path.exists():
            return []
        rows: list[dict] = []
        with path.open("r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                rows.append(json.loads(line))
        return rows

    def _append_jsonl(self, path: Path, obj: dict) -> None:
        with path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(obj, ensure_ascii=False) + "\n")

    # Nodes
    def list_nodes(self) -> list[Node]:
        return [self._node_adapter.validate_python(x) for x in self._read_jsonl(self.paths.nodes_path)]

    def upsert_node(self, node: Node) -> None:
        # Append-only: latest record wins (simple and robust)
        self._append_jsonl(self.paths.nodes_path, node.model_dump(mode="json"))

    def get_node(self, node_id: str) -> Node | None:
        nodes = self.list_nodes()
        for node in reversed(nodes):
            if node.node_id == node_id:
                return node
        return None

    # Contradictions
    def list_contradictions(self) -> list[Contradiction]:
        return [
            self._contr_adapter.validate_python(x)
            for x in self._read_jsonl(self.paths.contradictions_path)
        ]

    def add_contradiction(self, contradiction: Contradiction) -> None:
        self._append_jsonl(self.paths.contradictions_path, contradiction.model_dump(mode="json"))

    # Runs
    def list_runs(self) -> list[RunLog]:
        return [self._run_adapter.validate_python(x) for x in self._read_jsonl(self.paths.runs_path)]

    def add_run(self, run: RunLog) -> None:
        self._append_jsonl(self.paths.runs_path, run.model_dump(mode="json"))


def latest_by_id(items: Iterable, attr: str) -> dict[str, object]:
    out: dict[str, object] = {}
    for item in items:
        out[getattr(item, attr)] = item
    return out
