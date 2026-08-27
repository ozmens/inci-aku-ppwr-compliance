import { FormEvent, useEffect, useState } from "react";
import { Link, useSearchParams } from "react-router-dom";
import { api } from "../api";
import { coverageLabel, declStatusLabel, MATERIAL_FAMILY_LABEL } from "../labels";

type MatrixRow = {
  component_code: string;
  description?: string;
  set_codes?: string[];
  set_count?: number;
  supplier_count: number;
  has_tds_supplier: boolean;
  preferred_supplier?: string | null;
  coverage: "covered" | "linked" | "gap";
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
};

export default function Components() {
  const [searchParams] = useSearchParams();
  const [q, setQ] = useState(() => searchParams.get("q") || "");
  const [linkedOnly, setLinkedOnly] = useState(true);
  const [rows, setRows] = useState<MatrixRow[]>([]);
  const [linkedTotal, setLinkedTotal] = useState(0);
  const [selected, setSelected] = useState<MatrixRow | null>(null);
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);

  async function load(query = q, onlyLinked = linkedOnly) {
    setErr("");
    setBusy(true);
    try {
      const r = await api.componentsMatrix(query, "starter", 80, onlyLinked);
      const list = r.components as MatrixRow[];
      setRows(list);
      setLinkedTotal(r.linked_components || 0);
      if (list[0]) {
        const keep = selected
          ? list.find((c) => c.component_code === selected.component_code)
          : null;
        setSelected(keep || list[0]);
      } else {
        setSelected(null);
      }
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  useEffect(() => {
    const fromUrl = searchParams.get("q") || "";
    if (fromUrl !== q) setQ(fromUrl);
    void load(fromUrl || q, fromUrl ? false : linkedOnly);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [searchParams]);

  function onSearch(e: FormEvent) {
    e.preventDefault();
    void load();
  }

  return (
    <section>
      <p className="eyebrow">Bileşen · tedarikçi kapsamı</p>
      <h1>Bileşen Matrisi</h1>
      <p className="lead">
        Ambalaj bileşeninin tedarikçisi var mı, TDS yüklü mü? Detay:{" "}
        <Link to="/suppliers">Tedarikçi</Link> · set: <Link to="/bom">Ambalaj BOM</Link>.
      </p>

      <form className="search-bar" onSubmit={onSearch}>
        <input
          value={q}
          onChange={(e) => setQ(e.target.value)}
          placeholder="Bileşen kodu / açıklama…"
        />
        <label className="muted" style={{ display: "flex", alignItems: "center", gap: "0.35rem" }}>
          <input
            type="checkbox"
            checked={linkedOnly}
            onChange={(e) => {
              setLinkedOnly(e.target.checked);
              void load(q, e.target.checked);
            }}
          />
          yalnız bağlılar
        </label>
        <button type="submit" disabled={busy}>
          {busy ? "…" : "Ara"}
        </button>
      </form>

      {err && <p className="error">{err}</p>}
      <p className="meta">
        {rows.length} satır · sistemde bağlı bileşen {linkedTotal}
      </p>

      <div className="split">
        <ul className="key-list">
          {rows.map((r) => (
            <li key={r.component_code}>
              <button
                type="button"
                className={selected?.component_code === r.component_code ? "picked-rev" : ""}
                onClick={() => setSelected(r)}
              >
                {r.component_code}
                <span
                  className={
                    r.coverage === "covered"
                      ? "act-badge green"
                      : r.coverage === "gap"
                        ? "act-badge purple"
                        : "act-badge"
                  }
                  style={{ marginLeft: "0.4rem" }}
                >
                  {coverageLabel(r.coverage)}
                </span>
                <span className="muted">
                  {r.preferred_supplier ? ` · ${r.preferred_supplier}` : ""}
                  {r.has_tds_supplier ? " · TDS" : ""}
                </span>
              </button>
            </li>
          ))}
          {rows.length === 0 && (
            <li>
              <p className="muted" style={{ padding: "0.75rem" }}>
                {linkedOnly
                  ? "Henüz bağlı bileşen yok — Tedarikçi sayfasından bağlayın."
                  : "Sonuç yok."}
              </p>
            </li>
          )}
        </ul>

        <div className="detail">
          {!selected && <p className="muted">Bileşen seçin</p>}
          {selected && (
            <>
              <h2>{selected.component_code}</h2>
              <p className="meta">{selected.description || "—"}</p>
              <p className="meta">
                Kapsam: <strong>{coverageLabel(selected.coverage)}</strong>
                {selected.set_codes?.[0] ? ` · set ${selected.set_codes[0]}` : ""}
                {typeof selected.set_count === "number" && selected.set_count > 0
                  ? ` · ${selected.set_count} set`
                  : ""}
              </p>
              <h3 className="section-title">
                Tedarikçiler ({selected.supplier_count})
              </h3>
              {selected.suppliers.length === 0 ? (
                <p className="muted">
                  Bağlı tedarikçi yok.{" "}
                  <Link to="/suppliers">Tedarikçi</Link> sayfasından bağlayın.
                </p>
              ) : (
                <ul className="activity-list">
                  {selected.suppliers.map((s) => (
                    <li key={s.supplier_id}>
                      <strong>{s.supplier_name}</strong>
                      <span className="muted">
                        {" "}
                        · {coverageLabel(s.readiness)}
                        {s.preferred ? " · tercih" : ""}
                        {s.has_tds ? " · TDS var" : " · TDS yok"}
                        {s.material_family
                          ? ` · ${MATERIAL_FAMILY_LABEL[s.material_family] || s.material_family}`
                          : ""}
                        {` · ağır metal ${declStatusLabel("hm", s.heavy_metals_status)}`}
                        {` · SVHC ${declStatusLabel("svhc", s.svhc_status)}`}
                        {` · PFAS ${declStatusLabel("pfas", s.pfas_status)}`}
                      </span>
                    </li>
                  ))}
                </ul>
              )}
              {selected.set_codes?.[0] && (
                <p className="meta" style={{ marginTop: "0.75rem" }}>
                  <Link to={`/bom/${encodeURIComponent(selected.set_codes[0])}`}>
                    Ambalaj BOM: {selected.set_codes[0]}
                  </Link>
                </p>
              )}
            </>
          )}
        </div>
      </div>
    </section>
  );
}
