from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from rks.agent import RKSAgent
from rks.models import ContradictionCreate, Layer, NodeCreate, NodePatch, NodeStatus, RunLog
from rks.storage import FileStorage


def create_app() -> FastAPI:
    app = FastAPI(title="RKS v2.0 (MVP)")

    base_dir = Path(__file__).resolve().parent
    templates = Jinja2Templates(directory=str(base_dir / "templates"))
    app.mount("/static", StaticFiles(directory=str(base_dir / "static")), name="static")

    storage = FileStorage(data_dir=(base_dir.parent / "data"))
    agent = RKSAgent(storage=storage)

    @app.get("/", response_class=HTMLResponse)
    def index(request: Request):
        nodes = agent.list_latest_nodes()
        contradictions = agent.list_contradictions()
        runs = list(reversed(agent.list_runs()))[:20]
        return templates.TemplateResponse(
            "index.html",
            {
                "request": request,
                "nodes": nodes,
                "contradictions": contradictions,
                "runs": runs,
                "layers": [x.value for x in Layer],
                "statuses": [x.value for x in NodeStatus],
            },
        )

    @app.post("/nodes/create")
    def create_node(
        layer: str = Form(...),
        title: str = Form(...),
        definition: str = Form(""),
        mechanism_hint: str = Form(""),
        domain_context: str = Form(""),
        status: str = Form(NodeStatus.EXPLORE.value),
    ):
        node = agent.create_node_skeleton(
            NodeCreate(
                layer=Layer(layer),
                title=title,
                definition=definition,
                mechanism_hint=mechanism_hint,
                domain_context=domain_context,
                status=NodeStatus(status),
            )
        )
        return RedirectResponse(url=f"/?focus={node.node_id}", status_code=303)

    @app.post("/nodes/patch")
    def patch_node(
        node_id: str = Form(...),
        title: str = Form(""),
        definition: str = Form(""),
        causal_chain: str = Form(""),
        boundary_condition: str = Form(""),
        failure_mode: str = Form(""),
        linked_nodes: str = Form(""),
        assumption_ledger: str = Form(""),
        evidence_examples: str = Form(""),
        cross_domain_validated: str = Form("off"),
        status: str = Form(NodeStatus.BUILD.value),
    ):
        patch = NodePatch(
            title=title,
            definition=definition,
            causal_chain=causal_chain,
            boundary_condition=boundary_condition,
            failure_mode=failure_mode,
            linked_nodes=[x.strip() for x in linked_nodes.split("\n") if x.strip()],
            assumption_ledger=[x.strip() for x in assumption_ledger.split("\n") if x.strip()],
            evidence_examples=[x.strip() for x in evidence_examples.split("\n") if x.strip()],
            cross_domain_validated=(cross_domain_validated.lower() in {"on", "true", "1", "yes"}),
            status=NodeStatus(status),
        )

        agent.patch_node(node_id=node_id, patch=patch)
        return RedirectResponse(url=f"/?focus={node_id}", status_code=303)

    @app.post("/contradictions/create")
    def create_contradiction(
        node_a: str = Form(...),
        node_b: str = Form(...),
        condition_a: str = Form(""),
        condition_b: str = Form(""),
        resolution_trigger: str = Form(""),
    ):
        agent.create_contradiction(
            ContradictionCreate(
                node_a=node_a,
                node_b=node_b,
                condition_a=condition_a,
                condition_b=condition_b,
                resolution_trigger=resolution_trigger,
            )
        )
        return RedirectResponse(url="/", status_code=303)

    @app.post("/runs/log")
    def log_run(
        problem: str = Form(...),
        prediction: str = Form(""),
        outcome: str = Form(""),
        delta: str = Form(""),
        suspected_layer: str = Form(""),
        related_nodes: str = Form(""),
        notes: str = Form(""),
    ):
        agent.log_run(
            RunLog(
                problem=problem,
                prediction=prediction,
                outcome=outcome,
                delta=delta,
                suspected_layer=Layer(suspected_layer) if suspected_layer else None,
                related_nodes=[x.strip() for x in related_nodes.split("\n") if x.strip()],
                notes=notes,
            )
        )
        return RedirectResponse(url="/", status_code=303)

    return app


app = create_app()
