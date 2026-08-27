"""İnci Akü PPWR Yazılımı — FastAPI backend.

Reads controlled delivery sets from Inci_Aku_PPWR_PIMS/output (READ-ONLY).
Never writes into delivery folders.
"""

from __future__ import annotations

import json
import os
import subprocess
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, Response, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel

from auth import (  # noqa: E402
    ChangePasswordRequest,
    CreateUserRequest,
    LoginRequest,
    ResetPasswordRequest,
    auth_status,
    change_own_password,
    clear_session_cookie,
    create_user,
    is_public_path,
    list_users_public,
    login as auth_login,
    public_user_from_payload,
    require_admin,
    require_user,
    reset_password,
    set_session_cookie,
    set_user_active,
)
from gaps import (  # noqa: E402
    delete_assignment,
    list_assignments,
    open_candidates_folder,
    save_assignment,
    scan_gaps,
    suggest,
    workspace_status,
)
from masters import (  # noqa: E402
    get_bom,
    get_product,
    industrial_summary,
    search_components,
    search_products,
    starter_summary,
)
from packs import (  # noqa: E402
    build_candidate_pack,
    delete_pack,
    get_pack,
    list_packs,
    open_pack_file,
    open_packs_folder,
    resolve_pack_file,
)
from engine import (  # noqa: E402
    engine_status,
    open_candidate_delivery_folder,
    open_candidate_engine,
    rebuild_candidate_engine,
)
from workspace_engine import (  # noqa: E402
    open_workspace_engine,
    open_workspace_engine_folder,
    rebuild_workspace_engine,
    workspace_engine_status,
)
from workspace_store import (  # noqa: E402
    bulk_create_from_codes,
    complete_all_incomplete,
    complete_pdfs,
    create_variant,
    desktop_zip_drop,
    get_workspace_product,
    list_products as ws_list_products,
    list_revisions,
    open_workspace_file,
    open_workspace_folder,
    recent_activity,
    resolve_workspace_file,
    revise_product,
    search_workspace,
    workspace_status as ws_status,
    EXPORTS as WS_EXPORTS,
    WEB_MODE as WS_WEB_MODE,
)
from customers import (  # noqa: E402
    customer_coverage,
    delete_customer,
    ensure_customer_packs,
    get_customer,
    list_customers,
    save_customer,
    zip_from_customer,
)
from suppliers import (  # noqa: E402
    analyze_document as supplier_analyze_document,
    component_matrix,
    delete_document as supplier_delete_document,
    delete_supplier,
    get_analysis as supplier_get_analysis,
    get_supplier,
    link_component as supplier_link_component,
    list_documents as supplier_list_documents,
    list_links as supplier_list_links,
    list_suppliers,
    open_document as supplier_open_document,
    open_supplier_folder,
    resolve_document_path as supplier_resolve_document,
    save_supplier,
    suppliers_for_component,
    unlink_component as supplier_unlink_component,
    update_link as supplier_update_link,
    upload_document as supplier_upload_document,
)
from pipeline import (  # noqa: E402
    bilingual_preview_texts,
    bilingual_set_preview,
    catalog as photo_catalog,
    export_preview_xlsx,
    resolve_set_photos,
    serve_photo,
    status as pipeline_status,
)

ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "app" / "config" / "app.config.json"
APP_VERSION = os.environ.get("INCI_PPWR_VERSION", "1.0.0")
WEB_MODE = (
    os.environ.get("INCI_PPWR_WEB", "").strip().lower() in {"1", "true", "yes"}
    or bool(os.environ.get("RENDER"))
    or WS_WEB_MODE
)


def _resolve_data_path(raw: str, *, fallbacks: list[Path] | None = None) -> Path:
    p = Path(raw)
    if not p.is_absolute():
        p = (ROOT / p).resolve()
    if p.exists():
        return p
    for alt in fallbacks or []:
        if alt.exists():
            return alt.resolve()
    return p


def load_config() -> dict:
    data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    env_root = os.environ.get("INCI_PPWR_DELIVERY_ROOT")
    if env_root:
        data["deliveryRoot"] = env_root
    env_cand = os.environ.get("INCI_PPWR_CANDIDATES_ROOT")
    if env_cand:
        data["candidatesRoot"] = env_cand
    data["deliveryRoot"] = str(
        _resolve_data_path(
            str(data.get("deliveryRoot") or "delivery"),
            fallbacks=[
                ROOT / "delivery",
                Path(r"C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output"),
            ],
        )
    )
    data["candidatesRoot"] = str(
        _resolve_data_path(str(data.get("candidatesRoot") or "candidates"), fallbacks=[ROOT / "candidates"])
    )
    return data


CFG = load_config()
DELIVERY_ROOT = Path(CFG["deliveryRoot"])
SCOPES = CFG["scopes"]
READ_ONLY = bool(CFG.get("readOnlyDeliveries", True))

STEMS = [
    ("01_Technical_File", "Technical File"),
    ("02_EU_DoC", "EU DoC"),
    ("03_Label", "Label"),
    ("04_Shipment_Statement", "Shipment Statement"),
]

app = FastAPI(
    title="İnci Akü PPWR Compliance Suite",
    version=APP_VERSION,
    docs_url=None if WEB_MODE else "/docs",
    redoc_url=None if WEB_MODE else "/redoc",
    openapi_url=None if WEB_MODE else "/openapi.json",
)

_cors = os.environ.get("INCI_PPWR_CORS_ORIGINS", "*")
_cors_list = [o.strip() for o in _cors.split(",") if o.strip()] or ["*"]
# Cookie auth cannot use wildcard origin with credentials
_cred = True
if _cors_list == ["*"]:
    _cred = False
app.add_middleware(
    CORSMiddleware,
    allow_origins=_cors_list,
    allow_credentials=_cred,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def auth_middleware(request: Request, call_next):
    path = request.url.path
    if is_public_path(path):
        return await call_next(request)
    if path.startswith("/api"):
        try:
            request.state.user = require_user(request)
        except HTTPException as exc:
            return JSONResponse({"detail": exc.detail}, status_code=exc.status_code)
    return await call_next(request)


@app.post("/api/auth/login")
def api_auth_login(req: LoginRequest, response: Response):
    result = auth_login(req.username, req.password)
    set_session_cookie(response, result["token"])
    return result


@app.post("/api/auth/logout")
def api_auth_logout(response: Response):
    clear_session_cookie(response)
    return {"ok": True}


@app.get("/api/auth/me")
def api_auth_me(request: Request):
    payload = require_user(request)
    return {"user": public_user_from_payload(payload)}


@app.get("/api/auth/status")
def api_auth_status():
    return auth_status()


@app.get("/api/auth/users")
def api_auth_users(request: Request):
    require_admin(request)
    return {"users": list_users_public()}


@app.post("/api/auth/users")
def api_auth_create_user(req: CreateUserRequest, request: Request):
    require_admin(request)
    return create_user(req)


@app.post("/api/auth/users/{user_id}/password")
def api_auth_reset_password(user_id: str, req: ResetPasswordRequest, request: Request):
    require_admin(request)
    return reset_password(user_id, req.password)


@app.post("/api/auth/password")
def api_auth_change_password(req: ChangePasswordRequest, request: Request):
    payload = require_user(request)
    uid = str(payload.get("sub") or payload.get("id") or "")
    return change_own_password(uid, req.current_password, req.new_password)


@app.post("/api/auth/users/{user_id}/active")
def api_auth_set_active(user_id: str, request: Request, active: bool = True):
    require_admin(request)
    return set_user_active(user_id, active)


class GapSuggestRequest(BaseModel):
    description: str
    product_code: str | None = None
    top_n: int = 5


class GapAssignRequest(BaseModel):
    product_code: str
    set_code: str
    description: str = ""
    note: str = ""
    form: str | None = None


class BilingualTextRequest(BaseModel):
    texts: list[str]
    kind: str = "component"  # component | product


class PipelineExportRequest(BaseModel):
    set_code: str
    scope: str = "starter"


class PackBuildRequest(BaseModel):
    product_code: str
    description: str = ""
    set_code: str = ""
    scope: str = "starter"
    skip_pdf: bool = False


class PackOpenRequest(BaseModel):
    product_code: str
    file: str


class VariantCreateRequest(BaseModel):
    product_code: str
    description: str = ""
    set_code: str = ""
    scope: str = "starter"
    reason: str = "İlk yayın"
    skip_pdf: bool = False
    issue: bool = True


class VariantReviseRequest(BaseModel):
    reason: str
    description: str | None = None
    set_code: str | None = None
    scope: str | None = None
    skip_pdf: bool = False


class WorkspaceOpenRequest(BaseModel):
    product_code: str
    file: str
    revision: str | None = None


class DesktopDropRequest(BaseModel):
    codes_text: str
    label: str = "MULTI"


class CustomerSaveRequest(BaseModel):
    name: str
    codes_text: str = ""
    product_codes: list[str] | None = None
    note: str = ""
    id: str | None = None


class CustomerZipRequest(BaseModel):
    label: str | None = None


class BulkCreateRequest(BaseModel):
    codes_text: str
    scope: str = "starter"
    reason: str = "Master’dan toplu oluşturma"
    skip_pdf: bool = False


class CustomerEnsureRequest(BaseModel):
    scope: str = "starter"
    skip_pdf: bool = False
    then_zip: bool = True
    label: str | None = None


class SupplierSaveRequest(BaseModel):
    name: str
    code: str = ""
    country: str = ""
    status: str = "ACTIVE"
    external_ref: str = ""
    note: str = ""
    contact: str = ""
    materials: str = ""
    id: str | None = None


class SupplierLinkRequest(BaseModel):
    component_code: str
    description: str = ""
    set_code: str = ""
    note: str = ""
    preferred: bool = False
    scope: str = "starter"


class SupplierDeclPatch(BaseModel):
    status: str | None = None
    evidence_date: str | None = None
    evidence_doc_id: str | None = None
    note: str | None = None
    substance_name: str | None = None
    candidate_list_date: str | None = None


class SupplierLinkUpdateRequest(BaseModel):
    preferred: bool | None = None
    note: str | None = None
    material_family: str | None = None
    recycled_content_pct: float | None = None
    recyclability_note: str | None = None
    heavy_metals: SupplierDeclPatch | None = None
    svhc: SupplierDeclPatch | None = None
    pfas: SupplierDeclPatch | None = None


@app.get("/api/suppliers")
def api_suppliers(q: str = ""):
    return list_suppliers(q)


@app.post("/api/suppliers/open-folder")
def api_suppliers_open_folder(supplier_id: str | None = None):
    if WEB_MODE:
        raise HTTPException(400, "Klasör açma yalnızca masaüstü kurulumunda")
    return open_supplier_folder(supplier_id)


@app.get("/api/suppliers/{supplier_id}")
def api_supplier(supplier_id: str):
    return get_supplier(supplier_id)


@app.post("/api/suppliers")
def api_supplier_save(req: SupplierSaveRequest):
    return save_supplier(
        name=req.name,
        code=req.code,
        country=req.country,
        status=req.status,
        external_ref=req.external_ref,
        note=req.note,
        contact=req.contact,
        materials=req.materials,
        supplier_id=req.id,
    )


@app.delete("/api/suppliers/{supplier_id}")
def api_supplier_delete(supplier_id: str):
    return delete_supplier(supplier_id)


@app.get("/api/suppliers/{supplier_id}/documents")
def api_supplier_docs(supplier_id: str):
    return supplier_list_documents(supplier_id)


@app.post("/api/suppliers/{supplier_id}/documents")
async def api_supplier_upload(
    supplier_id: str,
    file: UploadFile = File(...),
    doc_type: str = Form("TDS"),
    title: str = Form(""),
    note: str = Form(""),
    analyze: bool = Form(True),
):
    return supplier_upload_document(
        supplier_id,
        file,
        doc_type=doc_type,
        title=title,
        note=note,
        analyze=analyze,
    )


@app.post("/api/suppliers/{supplier_id}/documents/{doc_id}/analyze")
def api_supplier_analyze(supplier_id: str, doc_id: str):
    return supplier_analyze_document(supplier_id, doc_id)


@app.get("/api/suppliers/{supplier_id}/analyses/{analysis_id}")
def api_supplier_analysis(supplier_id: str, analysis_id: str):
    return supplier_get_analysis(supplier_id, analysis_id)


@app.post("/api/suppliers/{supplier_id}/documents/{doc_id}/open")
def api_supplier_open_doc(supplier_id: str, doc_id: str):
    return supplier_open_document(supplier_id, doc_id)


@app.get("/api/suppliers/{supplier_id}/documents/{doc_id}/file")
def api_supplier_doc_file(supplier_id: str, doc_id: str):
    path = supplier_resolve_document(supplier_id, doc_id)
    return FileResponse(path, filename=path.name)


@app.delete("/api/suppliers/{supplier_id}/documents/{doc_id}")
def api_supplier_delete_doc(supplier_id: str, doc_id: str):
    return supplier_delete_document(supplier_id, doc_id)


@app.get("/api/suppliers/{supplier_id}/links")
def api_supplier_links(supplier_id: str):
    return supplier_list_links(supplier_id)


@app.post("/api/suppliers/{supplier_id}/links")
def api_supplier_link(supplier_id: str, req: SupplierLinkRequest):
    return supplier_link_component(
        supplier_id,
        component_code=req.component_code,
        description=req.description,
        set_code=req.set_code,
        note=req.note,
        preferred=req.preferred,
        scope=req.scope,
    )


@app.put("/api/suppliers/{supplier_id}/links/{link_id}")
def api_supplier_update_link(supplier_id: str, link_id: str, req: SupplierLinkUpdateRequest):
    patch = req.model_dump(exclude_unset=True)
    for key in ("heavy_metals", "svhc", "pfas"):
        block = patch.get(key)
        if isinstance(block, dict):
            patch[key] = {k: v for k, v in block.items() if v is not None}
    return supplier_update_link(supplier_id, link_id, patch)


@app.delete("/api/suppliers/{supplier_id}/links/{link_id}")
def api_supplier_unlink(supplier_id: str, link_id: str):
    return supplier_unlink_component(supplier_id, link_id)


@app.get("/api/components/search")
def api_components_search(q: str = "", kind: str = "starter", limit: int = 80):
    return search_components(kind, q, limit=max(1, min(limit, 300)))


@app.get("/api/components/matrix")
def api_components_matrix(
    q: str = "",
    kind: str = "starter",
    limit: int = 80,
    linked_only: bool = False,
):
    return component_matrix(
        q=q,
        kind=kind,
        limit=max(1, min(limit, 300)),
        linked_only=linked_only,
    )


@app.get("/api/components/{component_code}/suppliers")
def api_component_suppliers(component_code: str):
    return suppliers_for_component(component_code)


@app.get("/api/customers")
def api_customers():
    return list_customers()


@app.get("/api/customers/{customer_id}")
def api_customer(customer_id: str):
    return get_customer(customer_id)


@app.get("/api/customers/{customer_id}/coverage")
def api_customer_coverage(customer_id: str):
    return customer_coverage(customer_id)


@app.post("/api/customers")
def api_customer_save(req: CustomerSaveRequest):
    return save_customer(
        name=req.name,
        codes_text=req.codes_text,
        product_codes=req.product_codes,
        note=req.note,
        customer_id=req.id,
    )


@app.delete("/api/customers/{customer_id}")
def api_customer_delete(customer_id: str):
    return delete_customer(customer_id)


@app.post("/api/customers/{customer_id}/desktop-zip")
def api_customer_zip(customer_id: str, req: CustomerZipRequest = CustomerZipRequest()):
    return zip_from_customer(customer_id, label=req.label)


@app.post("/api/customers/{customer_id}/ensure-packs")
def api_customer_ensure(customer_id: str, req: CustomerEnsureRequest = CustomerEnsureRequest()):
    return ensure_customer_packs(
        customer_id,
        scope=req.scope,
        skip_pdf=req.skip_pdf,
        then_zip=req.then_zip,
        label=req.label,
    )


@app.post("/api/workspace/bulk-create")
def api_ws_bulk(req: BulkCreateRequest):
    flat: list[str] = []
    for ln in (req.codes_text or "").replace(",", "\n").replace(";", "\n").splitlines():
        part = ln.strip()
        if part:
            flat.append(part)
    return bulk_create_from_codes(
        codes=flat,
        scope=req.scope,
        reason=req.reason,
        skip_pdf=req.skip_pdf,
    )


@app.post("/api/workspace/products/{product_code}/complete-pdfs")
def api_ws_complete_pdfs(product_code: str, revision: str | None = None):
    return complete_pdfs(product_code, revision)


@app.post("/api/workspace/complete-incomplete")
def api_ws_complete_all():
    return complete_all_incomplete()


@app.get("/api/workspace/activity")
def api_ws_activity(limit: int = 30):
    return recent_activity(limit=max(1, min(limit, 100)))


@app.get("/api/workspace/status")
def api_ws_status():
    return ws_status()


@app.get("/api/workspace/search")
def api_ws_search(q: str = "", limit: int = 100):
    return search_workspace(q, limit=max(1, min(limit, 500)))


@app.get("/api/workspace/products")
def api_ws_products():
    return {"products": ws_list_products()}


@app.get("/api/workspace/products/{product_code}")
def api_ws_product(product_code: str):
    return get_workspace_product(product_code)


@app.get("/api/workspace/products/{product_code}/revisions")
def api_ws_revisions(product_code: str):
    return {"product_code": product_code, "revisions": list_revisions(product_code)}


@app.post("/api/workspace/products")
def api_ws_create(req: VariantCreateRequest):
    return create_variant(
        product_code=req.product_code,
        description=req.description,
        set_code=req.set_code,
        scope=req.scope,
        reason=req.reason,
        skip_pdf=req.skip_pdf,
        issue=req.issue,
    )


@app.post("/api/workspace/products/{product_code}/revise")
def api_ws_revise(product_code: str, req: VariantReviseRequest):
    return revise_product(
        product_code=product_code,
        reason=req.reason,
        description=req.description,
        set_code=req.set_code,
        scope=req.scope,
        skip_pdf=req.skip_pdf,
    )


@app.post("/api/workspace/open")
def api_ws_open(req: WorkspaceOpenRequest):
    return open_workspace_file(req.product_code, req.file, req.revision)


@app.get("/api/workspace/file")
def api_ws_file(product_code: str, file: str, revision: str | None = None):
    path = resolve_workspace_file(product_code, file, revision)
    return FileResponse(path, filename=path.name)


@app.get("/api/workspace/zip-download")
def api_ws_zip_download(name: str):
    safe = Path(name).name
    if not safe.lower().endswith(".zip") or ".." in safe:
        raise HTTPException(400, "Invalid zip name")
    for base in (WS_EXPORTS, Path.home() / "Desktop"):
        path = (base / safe).resolve()
        try:
            path.relative_to(base.resolve())
        except ValueError:
            continue
        if path.exists():
            return FileResponse(path, filename=safe, media_type="application/zip")
    raise HTTPException(404, "ZIP not found")


@app.post("/api/workspace/open-folder")
def api_ws_open_folder(product_code: str | None = None):
    if WEB_MODE:
        raise HTTPException(400, "Klasör açma yalnızca masaüstü kurulumunda")
    return open_workspace_folder(product_code)


@app.post("/api/workspace/desktop-zip")
def api_ws_desktop_zip(req: DesktopDropRequest):
    flat: list[str] = []
    for ln in (req.codes_text or "").replace(",", "\n").replace(";", "\n").splitlines():
        part = ln.strip()
        if part:
            flat.append(part)
    return desktop_zip_drop(codes=flat, label=req.label)


@app.get("/api/workspace/engine/status")
def api_ws_engine_status():
    return workspace_engine_status()


@app.post("/api/workspace/engine/rebuild")
def api_ws_engine_rebuild(issued_only: bool = False):
    return rebuild_workspace_engine(issued_only=issued_only)


@app.post("/api/workspace/engine/open")
def api_ws_engine_open():
    if WEB_MODE:
        raise HTTPException(400, "Document Engine yalnızca masaüstü kurulumunda")
    return open_workspace_engine()


@app.post("/api/workspace/engine/open-folder")
def api_ws_engine_open_folder():
    if WEB_MODE:
        raise HTTPException(400, "Klasör açma yalnızca masaüstü kurulumunda")
    return open_workspace_engine_folder()


@app.get("/api/packs")
def api_list_packs():
    return list_packs()


@app.get("/api/packs/file")
def api_pack_file(product_code: str, file: str):
    path = resolve_pack_file(product_code, file)
    return FileResponse(path, filename=path.name)


@app.get("/api/packs/{product_code}")
def api_get_pack(product_code: str):
    return get_pack(product_code)


@app.post("/api/packs/build")
def api_build_pack(req: PackBuildRequest):
    result = build_candidate_pack(
        product_code=req.product_code,
        description=req.description or None,
        set_code=req.set_code or None,
        scope=req.scope,
        skip_pdf=req.skip_pdf,
    )
    # auto-refresh candidate engine when pack is complete
    if result.get("complete"):
        try:
            result["engine"] = rebuild_candidate_engine(require_complete=True)
        except Exception as e:
            result["engine_error"] = str(e)
    return result


@app.post("/api/packs/open")
def api_open_pack_file(req: PackOpenRequest):
    return open_pack_file(req.product_code, req.file)


@app.post("/api/packs/open-folder")
def api_open_packs_folder():
    if WEB_MODE:
        raise HTTPException(400, "Klasör açma yalnızca masaüstü kurulumunda")
    return open_packs_folder()


@app.delete("/api/packs/{product_code}")
def api_delete_pack(product_code: str):
    return delete_pack(product_code)


@app.get("/api/engine/status")
def api_engine_status():
    return engine_status()


@app.post("/api/engine/rebuild")
def api_engine_rebuild():
    return rebuild_candidate_engine(require_complete=True)


@app.post("/api/engine/open")
def api_engine_open():
    if WEB_MODE:
        raise HTTPException(400, "Document Engine yalnızca masaüstü kurulumunda")
    return open_candidate_engine()


@app.post("/api/engine/open-folder")
def api_engine_open_folder():
    if WEB_MODE:
        raise HTTPException(400, "Klasör açma yalnızca masaüstü kurulumunda")
    return open_candidate_delivery_folder()


@app.get("/api/pipeline/status")
def api_pipeline_status():
    return pipeline_status()


@app.get("/api/pipeline/photos/catalog")
def api_photo_catalog(scope: str | None = None):
    return photo_catalog(scope)


@app.get("/api/pipeline/photos/resolve")
def api_photo_resolve(set_code: str, scope: str = "starter"):
    return resolve_set_photos(set_code, scope)


@app.get("/api/pipeline/photos/file")
def api_photo_file(rel: str):
    return serve_photo(rel)


@app.post("/api/pipeline/bilingual/texts")
def api_bilingual_texts(req: BilingualTextRequest):
    return bilingual_preview_texts(req.texts, kind=req.kind)


@app.get("/api/pipeline/bilingual/set")
def api_bilingual_set(set_code: str, scope: str = "starter"):
    return bilingual_set_preview(set_code, scope)


@app.post("/api/pipeline/export-preview")
def api_pipeline_export(req: PipelineExportRequest):
    return export_preview_xlsx(req.set_code, req.scope)


@app.get("/api/gaps/workspace")
def gaps_workspace():
    return workspace_status()


@app.get("/api/gaps/scan")
def gaps_scan(kind: str = "starter", limit: int = 200):
    return scan_gaps(kind, limit=max(1, min(limit, 500)))


@app.post("/api/gaps/suggest")
def gaps_suggest(req: GapSuggestRequest):
    return suggest(req.description, req.product_code, top_n=max(1, min(req.top_n, 12)))


@app.get("/api/gaps/assignments")
def gaps_assignments():
    return list_assignments()


@app.post("/api/gaps/assignments")
def gaps_save(req: GapAssignRequest):
    return save_assignment(req.model_dump())


@app.delete("/api/gaps/assignments/{assignment_id}")
def gaps_delete(assignment_id: str):
    return delete_assignment(assignment_id)


@app.post("/api/gaps/open-folder")
def gaps_open_folder():
    if WEB_MODE:
        raise HTTPException(400, "Klasör açma yalnızca masaüstü kurulumunda")
    return open_candidates_folder()


@app.get("/api/masters/summary")
def masters_summary():
    out: dict = {}
    try:
        out["starter"] = starter_summary()
    except Exception as e:
        out["starter"] = {"error": str(e)}
    try:
        out["industrial"] = industrial_summary()
    except Exception as e:
        out["industrial"] = {"error": str(e)}
    return out


@app.get("/api/masters/{kind}/products")
def master_products(kind: str, q: str | None = None, limit: int = 50):
    return search_products(kind, q, limit=max(1, min(limit, 200)))


@app.get("/api/masters/{kind}/products/{code}")
def master_product(kind: str, code: str):
    return get_product(kind, code)


@app.get("/api/masters/{kind}/sets/{set_code}/bom")
def master_bom(kind: str, set_code: str):
    return get_bom(kind, set_code)


def scope_path(scope_key: str) -> Path:
    if scope_key not in SCOPES:
        raise HTTPException(404, f"Unknown scope: {scope_key}")
    p = DELIVERY_ROOT / SCOPES[scope_key]
    if not p.exists():
        raise HTTPException(404, f"Delivery folder missing: {p}")
    return p


def docs_subdir(scope_key: str) -> str:
    return {
        "starter": "01_PRODUCTS",
        "industrial": "01_PRODUCTS",
        "container": "01_CONFIGS",
        "component": "01_VARIANTS",
    }[scope_key]


@app.get("/api/health")
def health():
    payload = {
        "ok": True,
        "version": APP_VERSION,
        "product": "İnci Akü PPWR Compliance Suite",
        "mode": "web" if WEB_MODE else "desktop",
        "auth_required": auth_status()["auth_required"],
        "readOnly": READ_ONLY,
        "exists": DELIVERY_ROOT.exists(),
    }
    if not WEB_MODE:
        payload.update(
            {
                "deliveryRoot": str(DELIVERY_ROOT),
                "candidatesRoot": str(ROOT / "candidates"),
                "packsRoot": str(ROOT / "candidates" / "packs"),
                "workspaceRoot": str(ROOT / "workspace"),
                "workspaceEngineRoot": str(ROOT / "workspace" / "PPWR_WORKSPACE_ENGINE"),
                "suppliersRoot": str(ROOT / "workspace" / "suppliers"),
            }
        )
    else:
        payload["deliveryRoot"] = "/data/delivery" if os.environ.get("RENDER") else "configured"
    return payload


@app.get("/api/scopes")
def list_scopes():
    out = []
    for key, folder in SCOPES.items():
        root = DELIVERY_ROOT / folder
        docs = root / docs_subdir(key)
        n = 0
        if docs.exists():
            n = sum(1 for p in docs.iterdir() if p.is_dir())
        engine = next((root / "00_CONTROL").glob("*ENGINE*.xlsx"), None) if (root / "00_CONTROL").exists() else None
        launcher = root / "00_AC_DOCUMENT_ENGINE.cmd"
        out.append(
            {
                "key": key,
                "folder": folder,
                "path": str(root),
                "records": n,
                "engine": str(engine) if engine else None,
                "launcher": str(launcher) if launcher.exists() else None,
                "exists": root.exists(),
            }
        )
    return {"scopes": out}


@app.get("/api/scopes/{scope_key}/keys")
def list_keys(scope_key: str, q: str | None = None, limit: int = 100):
    root = scope_path(scope_key)
    docs = root / docs_subdir(scope_key)
    keys = sorted(p.name for p in docs.iterdir() if p.is_dir()) if docs.exists() else []
    if q:
        qq = q.strip().upper()
        keys = [k for k in keys if qq in k.upper()]
    total = len(keys)
    keys = keys[: max(1, min(limit, 500))]
    return {"scope": scope_key, "total": total, "keys": keys}


@app.get("/api/scopes/{scope_key}/records/{key}")
def get_record(scope_key: str, key: str):
    root = scope_path(scope_key)
    folder = root / docs_subdir(scope_key) / key
    if not folder.is_dir():
        raise HTTPException(404, f"Record not found: {key}")
    files = []
    for stem, label in STEMS:
        for ext, kind in (("docx", "WORD"), ("pdf", "PDF")):
            path = folder / f"{stem}.{ext}"
            files.append(
                {
                    "stem": stem,
                    "label": label,
                    "kind": kind,
                    "name": path.name,
                    "exists": path.exists(),
                    "path": str(path) if path.exists() else None,
                    "size": path.stat().st_size if path.exists() else 0,
                }
            )
    return {"scope": scope_key, "key": key, "folder": str(folder), "files": files}


class OpenRequest(BaseModel):
    scope: str
    key: str
    file: str  # e.g. 01_Technical_File.docx


@app.post("/api/open")
def open_file(req: OpenRequest):
    """Open (desktop) or prepare download (web) for a delivery file."""
    root = scope_path(req.scope)
    path = (root / docs_subdir(req.scope) / req.key / req.file).resolve()
    # path traversal guard
    try:
        path.relative_to((root / docs_subdir(req.scope)).resolve())
    except ValueError as e:
        raise HTTPException(400, "Invalid path") from e
    if not path.exists():
        raise HTTPException(404, f"File missing: {path}")
    if WEB_MODE or not hasattr(os, "startfile"):
        return {
            "opened": str(path),
            "download": True,
            "download_url": (
                f"/api/file?scope={req.scope}&key={req.key}&file={req.file}"
            ),
        }
    os.startfile(str(path))  # Windows
    return {"opened": str(path), "download": False}


@app.post("/api/scopes/{scope_key}/open-engine")
def open_engine(scope_key: str):
    if WEB_MODE:
        raise HTTPException(400, "Document Engine launcher yalnızca masaüstü kurulumunda")
    root = scope_path(scope_key)
    cmd = root / "00_AC_DOCUMENT_ENGINE.cmd"
    if not cmd.exists():
        raise HTTPException(404, "Engine launcher missing")
    subprocess.Popen(["cmd", "/c", str(cmd)], cwd=str(root))
    return {"launched": str(cmd)}


@app.get("/api/file")
def download_file(scope: str, key: str, file: str):
    root = scope_path(scope)
    path = (root / docs_subdir(scope) / key / file).resolve()
    try:
        path.relative_to((root / docs_subdir(scope)).resolve())
    except ValueError as e:
        raise HTTPException(400, "Invalid path") from e
    if not path.exists():
        raise HTTPException(404, "Missing")
    return FileResponse(path, filename=path.name)


# ── Production: serve built React UI from app/dist ──────────────────────────
_UI_DIST = ROOT / "app" / "dist"
if _UI_DIST.is_dir() and (WEB_MODE or os.environ.get("INCI_PPWR_SERVE_UI", "").strip() in {"1", "true", "yes"}):
    from fastapi.responses import FileResponse as _FR
    from fastapi.staticfiles import StaticFiles

    _assets = _UI_DIST / "assets"
    if _assets.is_dir():
        app.mount("/assets", StaticFiles(directory=str(_assets)), name="ui-assets")

    @app.get("/")
    def ui_index():
        return _FR(
            _UI_DIST / "index.html",
            headers={"Cache-Control": "no-store, no-cache, must-revalidate", "Pragma": "no-cache"},
        )

    @app.get("/{full_path:path}")
    def ui_spa(full_path: str):
        blocked = ("api/", "docs", "redoc", "openapi.json")
        if any(full_path == b or full_path.startswith(b) for b in blocked):
            raise HTTPException(404, "Not found")
        candidate = (_UI_DIST / full_path).resolve()
        try:
            candidate.relative_to(_UI_DIST.resolve())
        except ValueError as e:
            raise HTTPException(404, "Not found") from e
        if candidate.is_file():
            return _FR(candidate)
        return _FR(
            _UI_DIST / "index.html",
            headers={"Cache-Control": "no-store, no-cache, must-revalidate", "Pragma": "no-cache"},
        )
