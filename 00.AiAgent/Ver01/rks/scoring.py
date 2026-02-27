from __future__ import annotations

from .models import Node


def compute_nms(node: Node) -> int:
    """Compute Node Maturity Score (0-5) following RKS v2.0.

    Interprets the textual spec as:
    0 = idea
    1 = has definition
    2 = has a mechanism (use mechanism_hint OR causal chain presence as proxy)
    3 = has >= 3 linked nodes
    4 = has real-world evidence/examples
    5 = cross-domain validated

    Returns an integer 0..5.
    """

    score = 0

    if node.title.strip():
        score = 0

    if node.definition.strip():
        score = max(score, 1)

    if node.mechanism_hint.strip() or node.causal_chain.strip():
        score = max(score, 2)

    if len([x for x in node.linked_nodes if x.strip()]) >= 3:
        score = max(score, 3)

    if any(e.strip() for e in node.evidence_examples):
        score = max(score, 4)

    if node.cross_domain_validated:
        score = max(score, 5)

    return score
