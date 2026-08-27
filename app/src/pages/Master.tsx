import { FormEvent, useEffect, useState } from "react";
import { Link } from "react-router-dom";
import { api, MasterProduct, MasterSummary } from "../api";
import { measureLabel, SCOPE_OPTIONS, statusLabel } from "../labels";

type Kind = "starter" | "industrial";

export default function Master() {
  const [kind, setKind] = useState<Kind>("starter");
  const [summary, setSummary] = useState<MasterSummary | null>(null);
  const [q, setQ] = useState("");
  const [products, setProducts] = useState<MasterProduct[]>([]);
  const [selected, setSelected] = useState<MasterProduct | null>(null);
  const [err, setErr] = useState("");
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    api.mastersSummary().then(setSummary).catch((e) => setErr(String(e)));
  }, []);

  useEffect(() => {
    setSelected(null);
    setProducts([]);
    void onSearch();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [kind]);

  async function onSearch(e?: FormEvent) {
    e?.preventDefault();
    setErr("");
    setSelected(null);
    setLoading(true);
    try {
      const r = await api.masterProducts(kind, q, 80);
      setProducts(r.products);
      if (r.products[0]) {
        await selectProduct(r.products[0].product_code);
      }
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setLoading(false);
    }
  }

  async function selectProduct(code: string) {
    setErr("");
    try {
      setSelected(await api.masterProduct(kind, code));
    } catch (ex) {
      setErr(String(ex));
    }
  }

  const live = summary?.[kind];
  const productCount =
    live && !("error" in live) && typeof live.products === "number" ? live.products : null;

  return (
    <section>
      <p className="eyebrow">Ürün ve set referansı</p>
      <h1>Master Veri</h1>
      <p className="lead">
        Ürün kartı ve ambalaj BOM satırlarını inceleyin. Bu ekran yalnızca okur; kayıtlı
        dosyaları değiştirmez.
      </p>

      <div className="scope-switcher">
        <button
          type="button"
          className={`scope-switch ${kind === "starter" ? "active" : ""}`}
          onClick={() => setKind("starter")}
        >
          STARTER
        </button>
        <button
          type="button"
          className={`scope-switch purple ${kind === "industrial" ? "active" : ""}`}
          onClick={() => setKind("industrial")}
        >
          INDUSTRIAL
        </button>
      </div>

      {summary && (
        <div className="kpi-row kpi-row-3">
          <article className="kpi">
            <strong>
              {"error" in (summary.starter || {})
                ? "—"
                : (summary.starter as { products?: number })?.products?.toLocaleString() ?? "—"}
            </strong>
            <span>STARTER ürün</span>
          </article>
          <article className="kpi">
            <strong>
              {"error" in (summary.industrial || {})
                ? "—"
                : (summary.industrial as { products?: number })?.products?.toLocaleString() ?? "—"}
            </strong>
            <span>INDUSTRIAL ürün</span>
          </article>
          <article className="kpi">
            <strong>
              {kind === "starter" && live && !("error" in live) && "unique_sets" in live
                ? String((live as { unique_sets: number }).unique_sets)
                : "—"}
            </strong>
            <span>Benzersiz ambalaj seti</span>
          </article>
        </div>
      )}

      <form className="search-bar" onSubmit={onSearch}>
        <select value={kind} onChange={(e) => setKind(e.target.value as Kind)}>
          {SCOPE_OPTIONS.filter((o) => o.value === "starter" || o.value === "industrial").map(
            (o) => (
              <option key={o.value} value={o.value}>
                {o.label}
              </option>
            ),
          )}
        </select>
        <input
          value={q}
          onChange={(e) => setQ(e.target.value)}
          placeholder="Ürün kodu veya açıklama…"
        />
        <button type="submit" disabled={loading}>
          {loading ? "…" : "Ara"}
        </button>
      </form>

      {productCount != null && (
        <p className="meta">
          Master: {productCount.toLocaleString()} ürün · gösterilen {products.length}
        </p>
      )}
      {err && <p className="error">{err}</p>}

      <div className="split">
        <ul className="key-list">
          {products.map((p) => (
            <li key={p.product_code}>
              <button
                type="button"
                className={selected?.product_code === p.product_code ? "picked-rev" : ""}
                onClick={() => void selectProduct(p.product_code)}
              >
                {p.product_code}
                {p.set_code ? <span className="muted"> · {p.set_code}</span> : null}
              </button>
            </li>
          ))}
          {products.length === 0 && !loading && (
            <li>
              <p className="muted" style={{ padding: "0.75rem" }}>
                Ara ile master ürünleri getirin.
              </p>
            </li>
          )}
        </ul>
        <div className="detail">
          {!selected && <p className="muted">Ürün seçin</p>}
          {selected && (
            <>
              <h2>{selected.product_code}</h2>
              <p className="meta">{selected.description || "—"}</p>
              <dl className="facts">
                <div>
                  <dt>Durum</dt>
                  <dd>{statusLabel(selected.status)}</dd>
                </div>
                <div>
                  <dt>Set</dt>
                  <dd>
                    {selected.set_code ? (
                      <Link to={`/bom/${encodeURIComponent(selected.set_code)}`}>
                        {selected.set_code}
                      </Link>
                    ) : (
                      "—"
                    )}
                  </dd>
                </div>
                <div>
                  <dt>Dara (kg)</dt>
                  <dd>{selected.tare_kg ?? "—"}</dd>
                </div>
              </dl>
              {selected.bom && selected.bom.length > 0 && (
                <>
                  <h3>Ambalaj listesi ({selected.bom.length})</h3>
                  <table className="data-table">
                    <thead>
                      <tr>
                        <th>Bileşen</th>
                        <th>Adet</th>
                        <th>Birim</th>
                        <th>Satır kg</th>
                      </tr>
                    </thead>
                    <tbody>
                      {selected.bom.map((line, i) => (
                        <tr key={`${line.component_code}-${i}`}>
                          <td>
                            <strong>{line.component_code}</strong>
                            <div className="muted">{line.description}</div>
                          </td>
                          <td>{measureLabel(line.qty)}</td>
                          <td>{measureLabel(line.uom)}</td>
                          <td>{measureLabel(line.line_weight)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </>
              )}
              {kind === "starter" && selected.set_code && (!selected.bom || selected.bom.length === 0) && (
                <p className="muted">
                  Ambalaj listesi yok veya set atanmamış.{" "}
                  <Link to={`/bom/${encodeURIComponent(selected.set_code)}`}>Set detayı</Link>
                </p>
              )}
            </>
          )}
        </div>
      </div>
    </section>
  );
}
