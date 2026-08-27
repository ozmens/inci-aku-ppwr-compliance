import { setWebMode } from "./runtime";
import { triggerDownload } from "./download";

const BASE = "";

export type ScopeInfo = {
  key: string;
  folder: string;
  path: string;
  records: number;
  engine: string | null;
  launcher: string | null;
  exists: boolean;
};

export type RecordDetail = {
  scope: string;
  key: string;
  folder: string;
  files: {
    stem: string;
    label: string;
    kind: string;
    name: string;
    exists: boolean;
    path: string | null;
    size: number;
  }[];
};

export type BomLine = {
  component_code: string;
  description: string;
  qty: number | string | null;
  uom: string;
  unit_weight?: number | string | null;
  line_weight?: number | string | null;
};

export type MasterProduct = {
  product_code: string;
  description: string;
  set_code: string;
  status: string;
  tare_kg: number | string | null;
  bom?: BomLine[];
};

export type MasterSummary = {
  starter?: Record<string, unknown> & { products?: number; unique_sets?: number; error?: string };
  industrial?: Record<string, unknown> & { products?: number; error?: string };
};

export type BomResponse = {
  set_code: string;
  lines: BomLine[];
  meta: {
    final_id?: string;
    tare_kg?: number | string | null;
    description?: string;
    product_count?: number | string | null;
  };
};

export type GapWorkspace = {
  candidates_root: string;
  assignments_file: string;
  xlsx: string;
  count: number;
  frozen_deliveries_writable: boolean;
  write_policy: string;
};

export type GapSuggestion = {
  set_code: string;
  peer_products: number;
  tare_kg?: number | string | null;
  bom_lines: number;
  final_id?: string;
  description?: string;
};

export type GapSuggestResult = {
  product_code: string;
  description: string;
  form: string;
  suggestions: GapSuggestion[];
};

export type GapAssignment = {
  id: string;
  product_code: string;
  description: string;
  form: string;
  set_code: string;
  final_id?: string;
  tare_kg?: number | string | null;
  bom_lines?: number;
  note?: string;
  status: string;
  created_at?: string;
  updated_at?: string;
};

export type GapScanResult = {
  kind: string;
  count: number;
  gaps: {
    product_code: string;
    description: string;
    status: string;
    set_code: string;
    form: string;
  }[];
  note?: string;
};

export type PipelineStatus = {
  photo_root: string;
  photo_root_exists: boolean;
  image_count: number;
  mapping_count: number;
  bilingual_rule: string;
  write_policy: string;
  candidates_root: string;
};

export type BilingualLine = {
  component_code: string;
  qty: number | string | null;
  uom: string;
  tr: string;
  en: string;
  unit_weight?: number | string | null;
  line_weight?: number | string | null;
};

export type BilingualSetPreview = {
  set_code: string;
  scope: string;
  meta: Record<string, unknown>;
  lines: BilingualLine[];
  rule: string;
};

export type PhotoHit = {
  component_code: string;
  name_tr: string;
  name_en: string;
  note: string;
  rel: string;
  exists: boolean;
};

export type PhotoResolveResult = {
  set_code: string;
  scope: string;
  photo_root: string;
  matched: number;
  missing: { component_code: string; description: string }[];
  photos: PhotoHit[];
  bom_line_count: number;
};

export type PackFile = {
  stem: string;
  kind: string;
  name: string;
  exists: boolean;
  size: number;
};

export type CandidatePack = {
  product_code: string;
  path: string;
  meta?: Record<string, unknown>;
  files: PackFile[];
  complete: boolean;
  docx: number;
  pdf: number;
};

export type PackBuildResult = {
  pack: Record<string, unknown> & {
    product_code: string;
    photos?: number;
    pdf_ok?: number;
    pdf_fail?: number;
    set_code?: string;
    tare_kg?: number | string | null;
  };
  folder: string;
  files: PackFile[];
  complete: boolean;
  write_policy: string;
};

export type WsProduct = {
  product_code: string;
  description: string;
  set_code: string;
  current_revision: string | null;
  status: string;
  revision_count: number;
  complete: boolean;
};

export type WsRevision = {
  revision: string;
  status: string;
  reason: string;
  built_at?: string;
  set_code?: string;
  complete: boolean;
  files: PackFile[];
};

export type WsProductDetail = {
  product: {
    product_code: string;
    description: string;
    set_code: string;
    current_revision: string | null;
    status: string;
  };
  revisions: WsRevision[];
  current_files: PackFile[];
  folder: string;
};

export type DesktopDropResult = {
  zip: string;
  desktop: string;
  included: { product_code: string; revision: string; files: number }[];
  missing: { product_code: string; error: string; revision?: string }[];
  count_ok: number;
  count_missing: number;
  customer_id?: string;
  customer_name?: string;
  download_url?: string;
  zip_name?: string;
  downloadHref?: string;
  pack?: string;
  note?: string;
};

export type CustomerCard = {
  id: string;
  name: string;
  code_count: number;
  product_codes: string[];
  note?: string;
  updated_at?: string;
};

export type CustomerCoverage = {
  customer_id: string;
  customer_name: string;
  total: number;
  ready: number;
  incomplete: number;
  missing: number;
  zip_ready: boolean;
  rows: {
    product_code: string;
    state: "ready" | "incomplete" | "missing";
    in_workspace: boolean;
    complete: boolean;
    status: string | null;
    revision: string | null;
    description: string;
    set_code: string;
  }[];
};

export type SupplierCard = {
  id: string;
  code: string;
  name: string;
  country?: string;
  status?: string;
  external_ref?: string;
  note?: string;
  contact?: string;
  materials?: string;
  doc_count: number;
  link_count?: number;
  has_tds: boolean;
  has_analysis: boolean;
  has_certificate?: boolean;
  readiness: "ready" | "partial" | "empty";
  updated_at?: string;
};

export type SubstanceDecl = {
  status?: string;
  evidence_date?: string;
  evidence_doc_id?: string;
  note?: string;
  substance_name?: string;
  candidate_list_date?: string;
};

export type SupplierLink = {
  id: string;
  component_code: string;
  description?: string;
  set_code?: string;
  note?: string;
  preferred?: boolean;
  scope?: string;
  linked_at?: string;
  updated_at?: string;
  material_family?: string;
  recycled_content_pct?: number | null;
  recyclability_note?: string;
  heavy_metals?: SubstanceDecl;
  svhc?: SubstanceDecl;
  pfas?: SubstanceDecl;
};

export type SupplierDocument = {
  id: string;
  doc_type: string;
  title: string;
  original_name: string;
  stored_name: string;
  note?: string;
  uploaded_at?: string;
  size?: number;
  exists?: boolean;
  analysis_id?: string | null;
};

export type SupplierAnalysis = {
  id: string;
  supplier_id: string;
  document_id: string;
  doc_type?: string;
  title?: string;
  analyzed_at?: string;
  char_count?: number;
  language_guess?: string;
  topics?: Record<string, number>;
  matched_terms?: string[];
  ppwr_signals?: string[];
  text_preview?: string;
  extract_ok?: boolean;
};

export type BulkCreateResult = {
  requested: number;
  created: {
    product_code: string;
    revision: string;
    status: string;
    complete: boolean;
    set_code?: string;
  }[];
  skipped: { product_code: string; reason: string }[];
  failed: { product_code: string; error: string }[];
  count_created: number;
  count_skipped: number;
  count_failed: number;
};

// Document generation (DOCX build + PDF conversion) can take minutes,
// especially via MS Word COM automation — give those calls a long budget.
const DOC_GEN_TIMEOUT_MS = 300000;

async function j<T>(path: string, init?: RequestInit): Promise<T> {
  const headers = new Headers(init?.headers || {});
  const token = localStorage.getItem("inci_ppwr_token");
  if (token && !headers.has("Authorization")) {
    headers.set("Authorization", `Bearer ${token}`);
  }
  const r = await fetch(`${BASE}${path}`, {
    ...init,
    headers,
    credentials: "include",
    signal: init?.signal ?? AbortSignal.timeout(15000),
  });
  if (r.status === 401 && !path.startsWith("/api/auth/")) {
    localStorage.removeItem("inci_ppwr_token");
    localStorage.removeItem("inci_ppwr_user");
    if (!window.location.pathname.startsWith("/login")) {
      window.location.assign("/login");
    }
  }
  if (!r.ok) throw new Error(`${r.status} ${await r.text()}`);
  if (r.status === 204) return undefined as T;
  return r.json() as Promise<T>;
}

export type AuthUser = {
  id: string;
  username: string;
  display_name?: string;
  role: string;
};

export const api = {
  authStatus: () =>
    j<{ auth_required: boolean; auth_disabled: boolean; users: number }>("/api/auth/status"),
  authLogin: async (username: string, password: string) => {
    const r = await j<{ token: string; user: AuthUser; expires_hours: number }>("/api/auth/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username, password }),
    });
    localStorage.setItem("inci_ppwr_token", r.token);
    localStorage.setItem("inci_ppwr_user", JSON.stringify(r.user));
    return r;
  },
  authLogout: async () => {
    try {
      await j<{ ok: boolean }>("/api/auth/logout", { method: "POST" });
    } finally {
      localStorage.removeItem("inci_ppwr_token");
      localStorage.removeItem("inci_ppwr_user");
    }
  },
  authMe: () => j<{ user: AuthUser }>("/api/auth/me"),
  authUsers: () => j<{ users: AuthUser[] }>("/api/auth/users"),
  authCreateUser: (body: {
    username: string;
    password: string;
    display_name?: string;
    role?: string;
  }) =>
    j<AuthUser>("/api/auth/users", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    }),
  authResetPassword: (id: string, password: string) =>
    j<{ ok: boolean }>(`/api/auth/users/${encodeURIComponent(id)}/password`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ password }),
    }),
  authChangePassword: (current_password: string, new_password: string) =>
    j<{ ok: boolean }>("/api/auth/password", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ current_password, new_password }),
    }),
  authSetActive: (id: string, active: boolean) =>
    j<AuthUser>(
      `/api/auth/users/${encodeURIComponent(id)}/active?active=${active ? "true" : "false"}`,
      { method: "POST" },
    ),
  health: async () => {
    const h = await j<{
      ok: boolean;
      deliveryRoot?: string;
      version?: string;
      candidatesRoot?: string;
      auth_required?: boolean;
      mode?: string;
      product?: string;
    }>("/api/health");
    if (h.mode) setWebMode(h.mode === "web");
    return h;
  },
  scopes: () => j<{ scopes: ScopeInfo[] }>("/api/scopes"),
  keys: (scope: string, q: string, limit = 80) =>
    j<{ total: number; keys: string[] }>(
      `/api/scopes/${scope}/keys?q=${encodeURIComponent(q)}&limit=${limit}`,
    ),
  record: (scope: string, key: string) =>
    j<RecordDetail>(`/api/scopes/${scope}/records/${encodeURIComponent(key)}`),
  openFile: async (scope: string, key: string, file: string) => {
    const r = await j<{ opened: string; download?: boolean; download_url?: string }>("/api/open", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ scope, key, file }),
    });
    if (r.download && r.download_url) {
      r.download_url = triggerDownload(r.download_url, file);
    }
    return r;
  },
  openEngine: (scope: string) =>
    j<{ launched: string }>(`/api/scopes/${scope}/open-engine`, { method: "POST" }),
  mastersSummary: () => j<MasterSummary>("/api/masters/summary"),
  masterProducts: (kind: string, q: string, limit = 80) =>
    j<{ count: number; products: MasterProduct[] }>(
      `/api/masters/${kind}/products?q=${encodeURIComponent(q)}&limit=${limit}`,
    ),
  masterProduct: (kind: string, code: string) =>
    j<MasterProduct>(`/api/masters/${kind}/products/${encodeURIComponent(code)}`),
  masterBom: (kind: string, setCode: string) =>
    j<BomResponse>(`/api/masters/${kind}/sets/${encodeURIComponent(setCode)}/bom`),
  gapsWorkspace: () => j<GapWorkspace>("/api/gaps/workspace"),
  gapsScan: (kind = "starter", limit = 100) =>
    j<GapScanResult>(`/api/gaps/scan?kind=${kind}&limit=${limit}`),
  gapsSuggest: (description: string, product_code?: string) =>
    j<GapSuggestResult>("/api/gaps/suggest", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ description, product_code, top_n: 5 }),
    }),
  gapsAssignments: () =>
    j<{ count: number; assignments: GapAssignment[]; candidates_root: string }>(
      "/api/gaps/assignments",
    ),
  gapsSave: (body: {
    product_code: string;
    set_code: string;
    description?: string;
    note?: string;
    form?: string;
  }) =>
    j<{ saved: GapAssignment; count: number }>("/api/gaps/assignments", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    }),
  gapsDelete: (id: string) =>
    j<{ deleted: string; count: number }>(`/api/gaps/assignments/${encodeURIComponent(id)}`, {
      method: "DELETE",
    }),
  gapsOpenFolder: () => j<{ opened: string }>("/api/gaps/open-folder", { method: "POST" }),
  pipelineStatus: () => j<PipelineStatus>("/api/pipeline/status"),
  pipelineBilingualTexts: (texts: string[], kind: "component" | "product" = "component") =>
    j<{ kind: string; items: { tr: string; en: string; display: string }[] }>(
      "/api/pipeline/bilingual/texts",
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ texts, kind }),
      },
    ),
  pipelineBilingualSet: (setCode: string, scope = "starter") =>
    j<BilingualSetPreview>(
      `/api/pipeline/bilingual/set?set_code=${encodeURIComponent(setCode)}&scope=${encodeURIComponent(scope)}`,
    ),
  pipelineResolvePhotos: (setCode: string, scope = "starter") =>
    j<PhotoResolveResult>(
      `/api/pipeline/photos/resolve?set_code=${encodeURIComponent(setCode)}&scope=${encodeURIComponent(scope)}`,
    ),
  pipelineExport: (setCode: string, scope = "starter") =>
    j<{
      path: string;
      set_code: string;
      bilingual_lines: number;
      photos_matched: number;
      photos_missing: number;
    }>("/api/pipeline/export-preview", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ set_code: setCode, scope }),
    }),
  packsList: () =>
    j<{ packs_root: string; count: number; packs: CandidatePack[] }>("/api/packs"),
  packsGet: (code: string) =>
    j<{
      product_code: string;
      folder: string;
      meta: Record<string, unknown>;
      files: PackFile[];
      complete: boolean;
    }>(`/api/packs/${encodeURIComponent(code)}`),
  packsBuild: (body: {
    product_code: string;
    description?: string;
    set_code?: string;
    scope?: string;
    skip_pdf?: boolean;
  }) =>
    j<PackBuildResult>("/api/packs/build", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
      signal: AbortSignal.timeout(DOC_GEN_TIMEOUT_MS),
    }),
  packsOpen: async (product_code: string, file: string) => {
    const r = await j<{ opened: string; download?: boolean; download_url?: string }>(
      "/api/packs/open",
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ product_code, file }),
      },
    );
    if (r.download && r.download_url) {
      r.download_url = triggerDownload(r.download_url, file);
    }
    return r;
  },
  packsOpenFolder: () => j<{ opened: string }>("/api/packs/open-folder", { method: "POST" }),
  packsDelete: (code: string) =>
    j<{ deleted: string; count: number }>(`/api/packs/${encodeURIComponent(code)}`, {
      method: "DELETE",
    }),
  engineStatus: () =>
    j<{
      delivery_root: string;
      engine: string | null;
      engine_exists: boolean;
      launcher: string | null;
      launcher_exists: boolean;
      packs: number;
      complete_packs: number;
      link_rule: string;
    }>("/api/engine/status"),
  engineRebuild: () =>
    j<{
      engine: string;
      launcher: string;
      records: number;
      qa: string;
      verify: { checked: number; missing: string[]; pass: boolean };
    }>("/api/engine/rebuild", { method: "POST" }),
  engineOpen: () => j<{ launched: string }>("/api/engine/open", { method: "POST" }),
  engineOpenFolder: () =>
    j<{ opened: string }>("/api/engine/open-folder", { method: "POST" }),
  wsStatus: () =>
    j<{
      workspace: string;
      products: number;
      issued: number;
      revision_scheme: string;
    }>("/api/workspace/status"),
  wsSearch: (q = "", limit = 100) =>
    j<{ source: string; q: string; total: number; products: WsProduct[] }>(
      `/api/workspace/search?q=${encodeURIComponent(q)}&limit=${limit}`,
    ),
  wsProducts: () => j<{ products: WsProduct[] }>("/api/workspace/products"),
  wsProduct: (code: string) =>
    j<WsProductDetail>(`/api/workspace/products/${encodeURIComponent(code)}`),
  wsCreate: (body: {
    product_code: string;
    description?: string;
    set_code?: string;
    scope?: string;
    reason?: string;
    skip_pdf?: boolean;
  }) =>
    j<{
      product: Record<string, string>;
      revision: { revision: string; status: string; reason?: string };
      files: PackFile[];
      complete: boolean;
    }>("/api/workspace/products", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    }),
  wsRevise: (
    code: string,
    body: {
      reason: string;
      description?: string;
      set_code?: string;
      scope?: string;
      skip_pdf?: boolean;
    },
  ) =>
    j<{
      product: Record<string, string>;
      revision: { revision: string; status: string };
      files: PackFile[];
      complete: boolean;
      superseded?: string;
    }>(`/api/workspace/products/${encodeURIComponent(code)}/revise`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
      signal: AbortSignal.timeout(DOC_GEN_TIMEOUT_MS),
    }),
  wsOpen: async (product_code: string, file: string, revision?: string) => {
    const r = await j<{
      opened: string;
      download?: boolean;
      download_url?: string;
    }>("/api/workspace/open", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ product_code, file, revision }),
    });
    if (r.download && r.download_url) {
      r.download_url = triggerDownload(r.download_url, file);
    } else if (r.download) {
      const q = new URLSearchParams({ product_code, file });
      if (revision) q.set("revision", revision);
      r.download_url = triggerDownload(`/api/workspace/file?${q}`, file);
    }
    return r;
  },
  wsDesktopZip: async (codes_text: string, label = "MULTI") => {
    const r = await j<DesktopDropResult>("/api/workspace/desktop-zip", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ codes_text, label }),
      signal: AbortSignal.timeout(DOC_GEN_TIMEOUT_MS),
    });
    if (r.download_url) {
      r.downloadHref = triggerDownload(r.download_url, r.zip_name);
    }
    return r;
  },
  wsBulkCreate: (codes_text: string, scope = "starter", skip_pdf = false) =>
    j<BulkCreateResult>("/api/workspace/bulk-create", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        codes_text,
        scope,
        reason: "Bulk import from master",
        skip_pdf,
      }),
      signal: AbortSignal.timeout(DOC_GEN_TIMEOUT_MS),
    }),
  wsCompletePdfs: (code: string, revision?: string) =>
    j<{
      product_code: string;
      revision: string;
      complete: boolean;
      status: string;
      pdf_ok: number;
      pdf_fail: number;
    }>(
      `/api/workspace/products/${encodeURIComponent(code)}/complete-pdfs${
        revision ? `?revision=${encodeURIComponent(revision)}` : ""
      }`,
      { method: "POST", signal: AbortSignal.timeout(DOC_GEN_TIMEOUT_MS) },
    ),
  wsCompleteAll: () =>
    j<{ count: number; completed: number; results: unknown[] }>(
      "/api/workspace/complete-incomplete",
      { method: "POST", signal: AbortSignal.timeout(DOC_GEN_TIMEOUT_MS) },
    ),
  wsActivity: (limit = 20) =>
    j<{ count: number; events: { at: string; action: string; [k: string]: unknown }[] }>(
      `/api/workspace/activity?limit=${limit}`,
    ),
  wsEngineStatus: () =>
    j<{
      delivery_root: string;
      engine: string | null;
      engine_exists: boolean;
      launcher: string | null;
      launcher_exists: boolean;
      complete_products: number;
      link_rule: string;
      role: string;
    }>("/api/workspace/engine/status"),
  wsEngineRebuild: (issued_only = false) =>
    j<{
      engine: string;
      launcher: string;
      records: number;
      qa: string;
      verify: { checked: number; missing: string[]; pass: boolean };
      link_rule: string;
    }>(
      `/api/workspace/engine/rebuild${issued_only ? "?issued_only=true" : ""}`,
      { method: "POST" },
    ),
  wsEngineOpen: () =>
    j<{ launched: string }>("/api/workspace/engine/open", { method: "POST" }),
  wsEngineOpenFolder: () =>
    j<{ opened: string }>("/api/workspace/engine/open-folder", { method: "POST" }),
  customersList: () =>
    j<{ count: number; customers: CustomerCard[] }>("/api/customers"),
  customersSave: (body: {
    name: string;
    codes_text?: string;
    product_codes?: string[];
    note?: string;
    id?: string;
  }) =>
    j<{ id: string; name: string; product_codes: string[] }>("/api/customers", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    }),
  customersDelete: (id: string) =>
    j<{ deleted: string }>(`/api/customers/${encodeURIComponent(id)}`, { method: "DELETE" }),
  customersCoverage: (id: string) =>
    j<CustomerCoverage>(`/api/customers/${encodeURIComponent(id)}/coverage`),
  customersZip: async (id: string, label?: string) => {
    const r = await j<DesktopDropResult>(`/api/customers/${encodeURIComponent(id)}/desktop-zip`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ label: label || null }),
      signal: AbortSignal.timeout(DOC_GEN_TIMEOUT_MS),
    });
    if (r.download_url) {
      r.downloadHref = triggerDownload(r.download_url, r.zip_name);
    }
    return r;
  },
  customersEnsure: async (id: string, then_zip = true) => {
    const r = await j<{
      customer_name?: string;
      bulk: BulkCreateResult;
      zip?: DesktopDropResult;
    }>(`/api/customers/${encodeURIComponent(id)}/ensure-packs`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ then_zip, scope: "starter" }),
      signal: AbortSignal.timeout(DOC_GEN_TIMEOUT_MS),
    });
    if (r.zip?.download_url) {
      r.zip.downloadHref = triggerDownload(r.zip.download_url, r.zip.zip_name);
    }
    return r;
  },  suppliersList: (q = "") =>
    j<{ count: number; suppliers: SupplierCard[] }>(
      `/api/suppliers?q=${encodeURIComponent(q)}`,
    ),
  suppliersGet: (id: string) =>
    j<
      SupplierCard & {
        documents: SupplierDocument[];
        links: SupplierLink[];
        folder: string;
      }
    >(`/api/suppliers/${encodeURIComponent(id)}`),
  suppliersSave: (body: {
    name: string;
    code?: string;
    country?: string;
    status?: string;
    external_ref?: string;
    note?: string;
    contact?: string;
    materials?: string;
    id?: string;
  }) =>
    j<SupplierCard & { documents: SupplierDocument[] }>("/api/suppliers", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    }),
  suppliersDelete: (id: string) =>
    j<{ deleted: string }>(`/api/suppliers/${encodeURIComponent(id)}`, { method: "DELETE" }),
  suppliersUpload: async (
    id: string,
    file: File,
    opts: { doc_type?: string; title?: string; note?: string; analyze?: boolean } = {},
  ) => {
    const fd = new FormData();
    fd.append("file", file);
    fd.append("doc_type", opts.doc_type || "TDS");
    fd.append("title", opts.title || "");
    fd.append("note", opts.note || "");
    fd.append("analyze", String(opts.analyze !== false));
    const r = await fetch(`${BASE}/api/suppliers/${encodeURIComponent(id)}/documents`, {
      method: "POST",
      body: fd,
    });
    if (!r.ok) throw new Error(`${r.status} ${await r.text()}`);
    return r.json() as Promise<{
      document: SupplierDocument;
      analysis: SupplierAnalysis | null;
      supplier: SupplierCard & { documents: SupplierDocument[] };
    }>;
  },
  suppliersAnalyze: (id: string, docId: string) =>
    j<SupplierAnalysis>(
      `/api/suppliers/${encodeURIComponent(id)}/documents/${encodeURIComponent(docId)}/analyze`,
      { method: "POST" },
    ),
  suppliersAnalysis: (id: string, analysisId: string) =>
    j<SupplierAnalysis>(
      `/api/suppliers/${encodeURIComponent(id)}/analyses/${encodeURIComponent(analysisId)}`,
    ),
  suppliersOpenDoc: async (id: string, docId: string) => {
    const r = await j<{
      opened: string;
      download?: boolean;
      download_url?: string;
    }>(
      `/api/suppliers/${encodeURIComponent(id)}/documents/${encodeURIComponent(docId)}/open`,
      { method: "POST" },
    );
    if (r.download && r.download_url) {
      r.download_url = triggerDownload(r.download_url);
    }
    return r;
  },
  suppliersDeleteDoc: (id: string, docId: string) =>
    j<{ deleted: string }>(
      `/api/suppliers/${encodeURIComponent(id)}/documents/${encodeURIComponent(docId)}`,
      { method: "DELETE" },
    ),
  suppliersOpenFolder: (id?: string) =>
    j<{ opened: string }>(
      `/api/suppliers/open-folder${id ? `?supplier_id=${encodeURIComponent(id)}` : ""}`,
      { method: "POST" },
    ),
  componentsSearch: (q = "", kind = "starter", limit = 40) =>
    j<{
      total: number;
      components: {
        component_code: string;
        description: string;
        set_codes: string[];
        set_count: number;
      }[];
    }>(
      `/api/components/search?q=${encodeURIComponent(q)}&kind=${encodeURIComponent(kind)}&limit=${limit}`,
    ),
  componentsMatrix: (q = "", kind = "starter", limit = 80, linked_only = false) =>
    j<{
      total: number;
      linked_components: number;
      components: {
        component_code: string;
        description?: string;
        set_codes?: string[];
        set_count?: number;
        supplier_count: number;
        has_tds_supplier: boolean;
        preferred_supplier?: string | null;
        coverage: string;
        suppliers: {
          supplier_id: string;
          supplier_name: string;
          preferred: boolean;
          has_tds: boolean;
          readiness: string;
          material_family?: string;
          heavy_metals_status?: string;
          svhc_status?: string;
          pfas_status?: string;
        }[];
      }[];
    }>(
      `/api/components/matrix?q=${encodeURIComponent(q)}&kind=${encodeURIComponent(kind)}&limit=${limit}&linked_only=${linked_only}`,
    ),
  suppliersLink: (
    id: string,
    body: {
      component_code: string;
      description?: string;
      set_code?: string;
      note?: string;
      preferred?: boolean;
      scope?: string;
    },
  ) =>
    j<{ link: SupplierLink; supplier: SupplierCard & { links: SupplierLink[] } }>(
      `/api/suppliers/${encodeURIComponent(id)}/links`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      },
    ),
  suppliersUnlink: (id: string, linkId: string) =>
    j<{ deleted: string }>(
      `/api/suppliers/${encodeURIComponent(id)}/links/${encodeURIComponent(linkId)}`,
      { method: "DELETE" },
    ),
  suppliersUpdateLink: (
    id: string,
    linkId: string,
    body: {
      preferred?: boolean;
      note?: string;
      material_family?: string;
      recycled_content_pct?: number | null;
      recyclability_note?: string;
      heavy_metals?: SubstanceDecl;
      svhc?: SubstanceDecl;
      pfas?: SubstanceDecl;
    },
  ) =>
    j<{ link: SupplierLink; supplier: SupplierCard & { links: SupplierLink[] } }>(
      `/api/suppliers/${encodeURIComponent(id)}/links/${encodeURIComponent(linkId)}`,
      {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      },
    ),
  componentSuppliers: (componentCode: string) =>
    j<{
      component_code: string;
      count: number;
      suppliers: {
        supplier_id: string;
        supplier_name: string;
        preferred: boolean;
        has_tds: boolean;
        readiness: string;
      }[];
    }>(`/api/components/${encodeURIComponent(componentCode)}/suppliers`),
};
