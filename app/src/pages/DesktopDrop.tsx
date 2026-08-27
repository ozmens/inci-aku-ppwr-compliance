import { FormEvent, useEffect, useMemo, useState } from "react";
import { Link } from "react-router-dom";
import { api, CustomerCard, DesktopDropResult } from "../api";
import DownloadLink from "../components/DownloadLink";
import { isWebMode } from "../runtime";

export default function DesktopDrop() {
  const [codesText, setCodesText] = useState("");
  const [label, setLabel] = useState("MUSTERI");
  const [customers, setCustomers] = useState<CustomerCard[]>([]);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [result, setResult] = useState<DesktopDropResult | null>(null);
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const web = isWebMode();

  useEffect(() => {
    api
      .customersList()
      .then((r) => setCustomers(r.customers))
      .catch(() => setCustomers([]));
  }, []);

  const codeCount = useMemo(() => {
    return codesText
      .split(/[\s,;]+/)
      .map((x) => x.trim())
      .filter(Boolean).length;
  }, [codesText]);

  function loadCustomer(c: CustomerCard) {
    setSelectedId(c.id);
    setLabel(c.name);
    setCodesText((c.product_codes || []).join("\n"));
    setResult(null);
    setErr("");
  }

  function clearForm() {
    setSelectedId(null);
    setLabel("MUSTERI");
    setCodesText("");
    setResult(null);
    setErr("");
  }

  async function onDrop(e: FormEvent) {
    e.preventDefault();
    setErr("");
    setResult(null);
    setBusy(true);
    try {
      setResult(await api.wsDesktopZip(codesText, label.trim() || "MULTI"));
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  return (
    <section>
      <p className="eyebrow">Çoklu ürün · ZIP</p>
      <h1>Paket ZIP</h1>
      <p className="lead">
        Ürün kodlarını yapıştırın veya kayıtlı <Link to="/customers">müşteri kartı</Link> seçin.
        Revizyonda hazır olan Word + PDF paketleri tek ZIP olarak iner.
      </p>

      {customers.length > 0 && (
        <div className="customer-chips" style={{ marginBottom: "1rem" }}>
          {customers.map((c) => (
            <button
              key={c.id}
              type="button"
              className={selectedId === c.id ? "picked-rev" : ""}
              onClick={() => loadCustomer(c)}
            >
              {c.name} ({c.code_count})
            </button>
          ))}
          <button type="button" className="btn-ghost" onClick={clearForm}>
            Temizle
          </button>
        </div>
      )}

      <form onSubmit={onDrop}>
        <label className="block-label">
          Etiket (ZIP adı)
          <input value={label} onChange={(e) => setLabel(e.target.value)} placeholder="MUSTERI_ADI" />
        </label>
        <label className="block-label">
          Ürün kodları
          <textarea
            value={codesText}
            onChange={(e) => {
              setCodesText(e.target.value);
              setSelectedId(null);
            }}
            rows={8}
            placeholder={"1000069\n1000070\n1000071"}
            required
          />
        </label>
        <p className="meta">{codeCount} kod · ZIP yalnızca hazır paketleri alır</p>
        <div className="engine-actions">
          <button type="submit" disabled={busy || codeCount === 0}>
            {busy ? "Hazırlanıyor…" : web ? "ZIP indir" : "Masaüstüne ZIP"}
          </button>
          <Link to="/customers" className="btn-ghost" style={{ display: "inline-flex", alignItems: "center", padding: "0.55rem 0.9rem" }}>
            Müşteri kartları
          </Link>
          <Link to="/workspace" className="btn-ghost" style={{ display: "inline-flex", alignItems: "center", padding: "0.55rem 0.9rem" }}>
            Revizyon
          </Link>
        </div>
      </form>

      {err && <p className="error">{err}</p>}
      {result && (
        <div className="detail" style={{ marginTop: "1rem" }}>
          <p className={result.count_missing ? "error" : "ok"}>
            ZIP hazır · OK {result.count_ok} · eksik {result.count_missing}
            {result.zip_name ? ` · ${result.zip_name}` : ""}
          </p>
          {(result.downloadHref || result.download_url) && (
            <p style={{ marginTop: "0.5rem" }}>
              <DownloadLink href={result.downloadHref || result.download_url || ""} label="ZIP’i indir" />
            </p>
          )}
          {result.included.length > 0 && (
            <>
              <h3 className="section-title">Dahil</h3>
              <ul className="meta">
                {result.included.map((x) => (
                  <li key={x.product_code}>
                    <Link to={`/search?q=${encodeURIComponent(x.product_code)}&source=workspace`}>
                      {x.product_code}
                    </Link>{" "}
                    · {x.revision} · {x.files} dosya
                  </li>
                ))}
              </ul>
            </>
          )}
          {result.missing.length > 0 && (
            <>
              <h3 className="section-title">Eksikler</h3>
              <ul className="meta">
                {result.missing.map((x) => (
                  <li key={x.product_code + (x.error || "")}>
                    <Link to={`/workspace`}>{x.product_code}</Link>: {x.error}
                    {" · "}
                    <Link to={`/packs`}>Aday Paket</Link>
                  </li>
                ))}
              </ul>
            </>
          )}
        </div>
      )}
    </section>
  );
}
