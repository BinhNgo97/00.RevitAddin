from __future__ import annotations

from dataclasses import dataclass

from .models import Node, NodeStatus
from .scoring import compute_nms


@dataclass(frozen=True)
class NodeValidationResult:
    is_complete: bool
    reasons: list[str]
    nms: int


def validate_node_for_graph(node: Node) -> NodeValidationResult:
    """Validate a node against the RKS v2.0 gate rules.

    - If any of causal_chain/boundary_condition/failure_mode is missing -> INCOMPLETE
      and should not be linked to the main graph.
    - If Level 2 not complete -> node cannot be Active.
    """

    reasons: list[str] = []

    if not node.causal_chain.strip():
        reasons.append("Missing causal_chain (Level 2)")
    if not node.boundary_condition.strip():
        reasons.append("Missing boundary_condition (Level 2)")
    if not node.failure_mode.strip():
        reasons.append("Missing failure_mode (Level 2)")

    nms = compute_nms(node)
    is_complete = len(reasons) == 0

    if node.status == NodeStatus.ACTIVE and not is_complete:
        reasons.append("Active nodes require Level-2 completion")

    return NodeValidationResult(is_complete=is_complete, reasons=reasons, nms=nms)
