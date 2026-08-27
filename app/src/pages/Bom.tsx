import { FormEvent, useEffect, useState } from "react";
import { Link, useParams } from "react-router-dom";
import { api, BomResponse } from "../api";
import { measureLabel } from "../labels";

const RECENT_KEY = "inci_ppwr_bom_sets";

export default function Bom() {
  const { setCode: routeSet } = useParams();
  const [setCode, setSetCode] = useState(routeSet ? decodeURIComponent(routeSet) : "");
  const [data, setData] = useState<BomResponse | null>(null);
  const [recents, setRecents] = useState<string[]>([]);
  const [err, setErr] = useState("");
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    try {
      const raw = sessionStorage.getItem(RECENT_KEY);
      setRecents(raw ? (JSON.parse(raw) as string[]).slice(0, 8) : []);
    } catch {
      setRecents([]);
    }
  }, []);

  function pushRecent(code: string) {
    try {
      const next = [code, ...recents.filter((x) => x !== code)].slice(0, 8);
      sessionStorage.setItem(RECENT_KEY, JSON.stringify(next));
      setRecents(next);
    } catch {
      /* ignore */
    }
  }

  async function load(code: string) {
    const sc = code.trim();
    if (!sc) return;
    setErr("");
    setLoading(true);
    try {
      setData(await api.masterBom("starter", sc));
      pushRecent(sc);
    } catch (ex) {
      setData(null);
      setErr(String(ex));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    if (routeSet) {
      const decoded = decodeURIComponent(routeSet);
      setSetCode(decoded);
      void load(decoded);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [routeSet]);

  function onSearch(e: FormEvent) {
    e.preventDefault();
    void load(setCode);
  }

  return (
    <section>
      <p className="eyebrow">Ambalaj seti</p>
      <h1>Ambalaj BOM</h1>
      <p className="lead">
        Set kodunu yazın; bileşen listesini görün. Tedarikçi kapsamı için{" "}
        <Link to="/components">bileşen matrisine</Link> gidin.
      </p>
      {recents.length > 0 && (
        <div className="customer-chips" style={{ marginBottom: "0.75rem" }}>
          {recents.map((c) => (
            <button
              key={c}
              type="button"
              className={data?.set_code === c ? "picked-rev" : ""}
              onClick={() => {
                setSetCode(c);
                void load(c);
              }}
            >
              {c}
            </button>
          ))}
        </div>
      )}
      <form className="search-bar" onSubmit={onSearch}>
        <input
          value={setCode}
          onChange={(e) => setSetCode(e.target.value)}
          placeholder="Ambalaj set kodu…"
        />
        <button type="submit" disabled={loading}>
          {loading ? "…" : "Getir"}
        </button>
        {data?.set_code && (
          <Link
            className="btn-ghost"
            to={`/dil-foto`}
            style={{ display: "inline-flex", alignItems: "center", padding: "0.55rem 0.9rem" }}
            onClick={() => {
              try {
                sessionStorage.setItem(
                  "inci_ppwr_pipeline_sets",
                  JSON.stringify(
                    [data.set_code, ...JSON.parse(sessionStorage.getItem("inci_ppwr_pipeline_sets") || "[]")].slice(
                      0,
                      6,
                    ),
                  ),
                );
              } catch {
                /* ignore */
              }
            }}
          >
            Dil & Foto →
          </Link>
        )}
      </form>
      {err && <p className="error">{err}</p>}
      {data && (
        <div className="detail">
          <h2>{data.set_code}</h2>
          {data.meta && Object.keys(data.meta).length > 0 && (
            <dl className="facts">
              <div>
                <dt>Nihai kod</dt>
                <dd>{data.meta.final_id || "—"}</dd>
              </div>
              <div>
                <dt>Dara (kg)</dt>
                <dd>{data.meta.tare_kg ?? "—"}</dd>
              </div>
              <div>
                <dt>Bu setteki ürün</dt>
                <dd>{data.meta.product_count ?? "—"}</dd>
              </div>
            </dl>
          )}
          {data.meta?.description && <p className="meta">{data.meta.description}</p>}
          <p className="meta">{data.lines.length} ambalaj satırı</p>
          <table className="data-table">
            <thead>
              <tr>
                <th>Bileşen</th>
                <th>Adet</th>
                <th>Birim</th>
                <th>Birim kg</th>
                <th>Satır kg</th>
              </tr>
            </thead>
            <tbody>
              {data.lines.map((line, i) => (
                <tr key={`${line.component_code}-${i}`}>
                  <td>
                    <strong>
                      <Link to={`/components?q=${encodeURIComponent(line.component_code)}`}>
                        {line.component_code}
                      </Link>
                    </strong>
                    <div className="muted">{line.description}</div>
                  </td>
                  <td>{measureLabel(line.qty)}</td>
                  <td>{measureLabel(line.uom)}</td>
                  <td>{measureLabel(line.unit_weight)}</td>
                  <td>{measureLabel(line.line_weight)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {data.lines.length === 0 && <p className="muted">Bu set için ambalaj satırı yok.</p>}
        </div>
      )}
      {!data && !err && !loading && (
        <p className="muted" style={{ marginTop: "1rem" }}>
          Set kodu girin veya son bakılanlardan seçin.
        </p>
      )}
    </section>
  );
}
