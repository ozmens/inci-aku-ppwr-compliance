/** Customer-facing Turkish labels. Never render raw backend keys. */

export const SCOPE_LABEL: Record<string, string> = {
  starter: "STARTER",
  industrial: "INDUSTRIAL",
  container: "CONTAINER",
  component: "COMPONENT",
  workspace: "Revizyon",
};

export const SCOPE_OPTIONS = [
  { value: "starter", label: "STARTER — bireysel" },
  { value: "industrial", label: "INDUSTRIAL — endüstriyel" },
  { value: "container", label: "CONTAINER — konteyner" },
  { value: "component", label: "COMPONENT — bileşen" },
] as const;

export function scopeLabel(key: string | null | undefined): string {
  const k = String(key || "").trim().toLowerCase();
  if (!k) return "—";
  return SCOPE_LABEL[k] || String(key);
}

export const STATUS_LABEL: Record<string, string> = {
  DRAFT: "Taslak",
  ISSUED: "Yayınlandı",
  SUPERSEDED: "Eski revizyon",
  ACTIVE: "Aktif",
  PENDING: "Beklemede",
  INACTIVE: "Pasif",
  CONTROLLED: "Kontrollü",
  "CONTROLLED PACKAGING SET": "Kontrollü ambalaj seti",
  "BOM DATA REQUIRED": "Ambalaj verisi eksik",
};

export function statusLabel(value: string | null | undefined): string {
  const raw = String(value || "").trim();
  if (!raw) return "—";
  return STATUS_LABEL[raw.toUpperCase()] || STATUS_LABEL[raw] || raw;
}

export const COVERAGE_LABEL: Record<string, string> = {
  covered: "Tam",
  linked: "Bağlı",
  gap: "Eksik",
  partial: "Kısmi",
  ready: "Hazır",
  empty: "Boş",
  incomplete: "PDF eksik",
  missing: "Eksik",
};

export function coverageLabel(value: string | null | undefined): string {
  const k = String(value || "").trim().toLowerCase();
  if (!k) return "—";
  return COVERAGE_LABEL[k] || String(value);
}

export const ACTION_LABEL: Record<string, string> = {
  desktop_zip: "ZIP paketi",
  customer_zip: "Müşteri ZIP",
  create_variant: "Yeni paket",
  revise: "Revizyon",
  issue: "Yayınlandı",
  complete_pdfs: "PDF tamamlandı",
  complete_all_incomplete: "Eksik PDF tamamlandı",
  bulk_create: "Toplu oluşturma",
  engine_rebuild: "İndeks yenilendi",
  workspace_engine_export: "Doküman indeksi güncellendi",
  supplier_save: "Tedarikçi kaydı",
  supplier_delete: "Tedarikçi silindi",
  supplier_upload: "Belge yüklendi",
  supplier_analyze: "Belge analizi",
  supplier_doc_delete: "Belge silindi",
  supplier_link: "Tedarikçi bağlandı",
  supplier_link_update: "Bileşen beyanı",
  supplier_unlink: "Bağlantı kaldırıldı",
};

export function actionLabel(action: string | null | undefined): string {
  const raw = String(action || "").trim();
  if (!raw) return "İşlem";
  if (ACTION_LABEL[raw]) return ACTION_LABEL[raw];
  const snake = raw.toLowerCase().replace(/\s+/g, "_");
  return ACTION_LABEL[snake] || "İşlem";
}

export function actionBadge(action: string | null | undefined): string {
  const a = String(action || "").toLowerCase();
  if (a.includes("zip") || a.includes("customer") || a.includes("desktop")) return "purple";
  if (a.includes("complete") || a.includes("issue") || a.includes("engine")) return "green";
  return "";
}

const REASON_MAP: Record<string, string> = {
  "bulk smoke": "Toplu oluşturma",
  "initial issue": "İlk yayın",
  "label correction smoke test": "Etiket düzeltmesi",
  "bulk import from master": "Master’dan toplu oluşturma",
};

export function reasonLabel(reason: string | null | undefined): string {
  const raw = String(reason || "").trim();
  if (!raw) return "";
  const mapped = REASON_MAP[raw.toLowerCase()];
  if (mapped) return mapped;
  if (/\bsmoke\b/i.test(raw)) {
    const cleaned = raw
      .replace(/\bsmoke\s*(test)?\b/gi, "")
      .replace(/\s+/g, " ")
      .replace(/^[-–—·\s]+|[-–—·\s]+$/g, "")
      .trim();
    return cleaned || "Toplu oluşturma";
  }
  const ensure = raw.match(/^ensure packs for customer\s*(.*)$/i);
  if (ensure) {
    const name = (ensure[1] || "").trim();
    return name ? `Müşteri Paketi: ${name}` : "Müşteri Paketi";
  }
  return raw;
}

export function qaLabel(qa: unknown): string {
  const s = String(qa || "").trim().toUpperCase();
  if (s === "PASS") return "Uygun";
  if (s === "FAIL") return "Uygun değil";
  return "";
}

export const DOC_TYPE_LABEL: Record<string, string> = {
  TDS: "TDS",
  ANALYSIS: "Analiz",
  CERTIFICATE: "Sertifika",
  HEAVY_METALS: "Ağır Metal Beyanı",
  SVHC: "REACH / SVHC Beyanı",
  PFAS: "PFAS Beyanı",
  OTHER: "Diğer",
};

export function docTypeLabel(value: string | null | undefined): string {
  const raw = String(value || "").trim();
  if (!raw) return "—";
  return DOC_TYPE_LABEL[raw.toUpperCase()] || raw;
}

export const MATERIAL_FAMILY_LABEL: Record<string, string> = {
  PE: "PE",
  PP: "PP",
  PET: "PET",
  PAPER: "Kâğıt",
  CARDBOARD: "Karton",
  WOOD: "Ahşap",
  STEEL: "Çelik",
  MIXED: "Karışık",
  OTHER: "Diğer",
};

export const HM_STATUS_LABEL: Record<string, string> = {
  unknown: "Seçilmedi",
  compliant: "Uygun (≤100 mg/kg)",
  non_compliant: "Uygun değil",
  no_evidence: "Kanıt yok",
};

export const SVHC_STATUS_LABEL: Record<string, string> = {
  unknown: "Seçilmedi",
  none: "SVHC yok",
  present: "SVHC var",
  no_declaration: "Beyan yok",
};

export const PFAS_STATUS_LABEL: Record<string, string> = {
  unknown: "Seçilmedi",
  not_added: "Kasıtlı yok",
  present: "Var",
  not_applicable: "Gıda teması yok (N/A)",
};

export function declStatusLabel(
  kind: "hm" | "svhc" | "pfas",
  value: string | null | undefined,
): string {
  const k = String(value || "unknown").trim().toLowerCase();
  if (kind === "hm") return HM_STATUS_LABEL[k] || k;
  if (kind === "svhc") return SVHC_STATUS_LABEL[k] || k;
  return PFAS_STATUS_LABEL[k] || k;
}

export function fileKindLabel(kind: string | null | undefined): string {
  const k = String(kind || "").toUpperCase();
  if (k === "WORD" || k === "DOCX") return "Word";
  if (k === "PDF") return "PDF";
  return String(kind || "—");
}

export type ActivityEvent = {
  at?: string;
  action?: string;
  product_code?: unknown;
  supplier_id?: unknown;
  customer_id?: unknown;
  doc_id?: unknown;
  revision?: unknown;
  records?: unknown;
  count?: unknown;
  count_ok?: unknown;
  name?: unknown;
  qa?: unknown;
  scope?: unknown;
  doc_type?: unknown;
  component_code?: unknown;
  [k: string]: unknown;
};

export function eventDetail(ev: ActivityEvent): string {
  const parts: string[] = [];
  if (ev.revision) parts.push(String(ev.revision));
  if (ev.records != null && ev.records !== "") parts.push(`${ev.records} kayıt`);
  if (ev.count_ok != null && ev.count_ok !== "") parts.push(`${ev.count_ok} hazır`);
  if (ev.count != null && ev.count !== "" && ev.records == null) parts.push(`${ev.count} paket`);
  if (ev.name) parts.push(String(ev.name));
  if (ev.component_code) parts.push(String(ev.component_code));
  if (ev.doc_type) parts.push(docTypeLabel(String(ev.doc_type)));
  const qa = qaLabel(ev.qa);
  if (qa) parts.push(qa);
  return parts.filter(Boolean).join(" · ") || "—";
}

export function eventCode(ev: ActivityEvent): string {
  return (
    String(ev.product_code || "") ||
    String(ev.supplier_id || "") ||
    String(ev.customer_id || "") ||
    String(ev.doc_id || "") ||
    "—"
  );
}

export const BILINGUAL_RULE_TR = "Türkçe düz · İngilizce italik";

const MEASURE_LABEL: Record<string, string> = {
  "MASS-BASED / N/A": "Ağırlık bazlı / Yok",
  "N/A / MASS-BASED": "Yok / Ağırlık bazlı",
  "N/A / MASS BASED": "Yok / Ağırlık bazlı",
  "MASS BASED / N/A": "Ağırlık bazlı / Yok",
  "MASS-BASED": "Ağırlık bazlı",
  "MASS BASED": "Ağırlık bazlı",
  "N/A": "Yok",
};

export function measureLabel(value: unknown, empty = "—"): string {
  if (value == null || value === "") return empty;
  if (typeof value === "number" && Number.isFinite(value)) return String(value);
  const raw = String(value).trim();
  if (!raw) return empty;
  const key = raw.replace(/\s+/g, " ").toUpperCase();
  if (MEASURE_LABEL[key]) return MEASURE_LABEL[key];
  return raw
    .replace(/MASS[-\s]?BASED/gi, "Ağırlık bazlı")
    .replace(/\bN\/A\b/gi, "Yok");
}
