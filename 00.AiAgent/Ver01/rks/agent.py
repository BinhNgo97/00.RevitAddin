from __future__ import annotations

from dataclasses import dataclass

from .models import Contradiction, ContradictionCreate, Layer, Node, NodeCreate, NodePatch, NodeStatus, RunLog
from .scoring import compute_nms
from .storage import FileStorage, latest_by_id
from .validation import validate_node_for_graph


@dataclass(frozen=True)
class AgentConfig:
    data_dir: str = "data"


class RKSAgent:
    """RKS v2.0 agent orchestrator (Explore/Build, gates, and logging)."""

    def __init__(self, storage: FileStorage):
        self.storage = storage

    # ---- Node lifecycle ----
    def create_node_skeleton(self, req: NodeCreate) -> Node:
        node = Node(
            layer=req.layer,
            title=req.title.strip(),
            definition=req.definition.strip(),
            mechanism_hint=req.mechanism_hint.strip(),
            domain_context=req.domain_context.strip(),
            status=req.status,
        )
        node.node_maturity_score = compute_nms(node)
        self.storage.upsert_node(node)
        return node

    def patch_node(self, node_id: str, patch: NodePatch) -> Node:
        existing = self.storage.get_node(node_id)
        if existing is None:
            raise KeyError(f"Node not found: {node_id}")

        updated = existing.model_copy(deep=True)
        for k, v in patch.model_dump(exclude_unset=True).items():
            setattr(updated, k, v)

        updated.touch()
        updated.node_maturity_score = compute_nms(updated)

        # Gate: Level 2 required for Active
        result = validate_node_for_graph(updated)
        if updated.status == NodeStatus.ACTIVE and not result.is_complete:
            updated.status = NodeStatus.BUILD

        self.storage.upsert_node(updated)
        return updated

    def list_latest_nodes(self) -> list[Node]:
        nodes = self.storage.list_nodes()
        latest = latest_by_id(nodes, "node_id")
        return sorted(latest.values(), key=lambda n: n.updated_at, reverse=True)  # type: ignore[arg-type]

    # ---- Contradictions ----
    def create_contradiction(self, req: ContradictionCreate) -> Contradiction:
        # Ensure both nodes exist
        if self.storage.get_node(req.node_a) is None:
            raise KeyError(f"Node A not found: {req.node_a}")
        if self.storage.get_node(req.node_b) is None:
            raise KeyError(f"Node B not found: {req.node_b}")

        c = Contradiction(
            node_a=req.node_a,
            node_b=req.node_b,
            condition_a=req.condition_a.strip(),
            condition_b=req.condition_b.strip(),
            resolution_trigger=req.resolution_trigger.strip(),
        )
        self.storage.add_contradiction(c)
        return c

    def list_contradictions(self) -> list[Contradiction]:
        return self.storage.list_contradictions()

    # ---- Reflection / runs ----
    def log_run(self, run: RunLog) -> RunLog:
        self.storage.add_run(run)
        return run

    def list_runs(self) -> list[RunLog]:
        return self.storage.list_runs()

    # ---- Expansion / compression (MVP stubs) ----
    def can_expand(self, node: Node) -> bool:
        return compute_nms(node) >= 3

    def explain_gates(self, node: Node) -> dict:
        res = validate_node_for_graph(node)
        return {
            "nms": res.nms,
            "is_complete": res.is_complete,
            "reasons": res.reasons,
            "can_expand": res.nms >= 3,
        }
