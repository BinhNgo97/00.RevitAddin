from __future__ import annotations

from datetime import datetime
from enum import Enum
from typing import Annotated, Any
from uuid import uuid4

from pydantic import BaseModel, Field


class Layer(str, Enum):
    ONTOLOGY = "Ontology"
    MECHANISM = "Mechanism"
    DOMAIN = "Domain"
    ACTION = "Action"
    REFLECTION = "Reflection"
    GOVERNANCE = "Governance"


class NodeStatus(str, Enum):
    EXPLORE = "Explore"
    BUILD = "Build"
    ACTIVE = "Active"


def _new_id(prefix: str) -> str:
    return f"{prefix}_{uuid4().hex[:12]}"


class Node(BaseModel):
    node_id: str = Field(default_factory=lambda: _new_id("N"))
    layer: Layer

    title: str = Field(min_length=1, max_length=200)

    definition: str = ""
    causal_chain: str = ""
    boundary_condition: str = ""
    failure_mode: str = ""

    linked_nodes: list[str] = Field(default_factory=list)
    assumption_ledger: list[str] = Field(default_factory=list)

    node_maturity_score: int = 0
    status: NodeStatus = NodeStatus.EXPLORE

    mechanism_hint: str = ""  # AI-suggested placeholder (Level 1)
    domain_context: str = ""  # optional

    evidence_examples: list[str] = Field(default_factory=list)  # Level 2: real examples
    cross_domain_validated: bool = False

    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime = Field(default_factory=datetime.utcnow)

    def touch(self) -> None:
        self.updated_at = datetime.utcnow()


class NodeCreate(BaseModel):
    layer: Layer
    title: str
    definition: str = ""
    mechanism_hint: str = ""
    domain_context: str = ""
    status: NodeStatus = NodeStatus.EXPLORE


class NodePatch(BaseModel):
    title: str | None = None
    definition: str | None = None
    causal_chain: str | None = None
    boundary_condition: str | None = None
    failure_mode: str | None = None

    linked_nodes: list[str] | None = None
    assumption_ledger: list[str] | None = None

    mechanism_hint: str | None = None
    domain_context: str | None = None

    evidence_examples: list[str] | None = None
    cross_domain_validated: bool | None = None

    status: NodeStatus | None = None


class Contradiction(BaseModel):
    contradiction_id: str = Field(default_factory=lambda: _new_id("C"))
    node_a: str
    node_b: str

    condition_a: str = ""
    condition_b: str = ""
    resolution_trigger: str = ""

    created_at: datetime = Field(default_factory=datetime.utcnow)


class ContradictionCreate(BaseModel):
    node_a: str
    node_b: str
    condition_a: str = ""
    condition_b: str = ""
    resolution_trigger: str = ""


class RunLog(BaseModel):
    run_id: str = Field(default_factory=lambda: _new_id("R"))
    problem: str

    prediction: str = ""
    outcome: str = ""
    delta: str = ""

    suspected_layer: Layer | None = None
    notes: str = ""

    related_nodes: list[str] = Field(default_factory=list)
    created_at: datetime = Field(default_factory=datetime.utcnow)


JsonDict = Annotated[dict[str, Any], Field(default_factory=dict)]
