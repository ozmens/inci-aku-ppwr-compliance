import { FormEvent, useEffect, useMemo, useState } from "react";
import { Link } from "react-router-dom";
import { api, CustomerCard, CustomerCoverage, DesktopDropResult } from "../api";
import DownloadLink from "../components/DownloadLink";
import { isWebMode } from "../runtime";
import { statusLabel } from "../labels";

const STATE_TR: Record<string, string> = {
  ready: "Hazır",
  incomplete: "PDF eksik",
  missing: "Eksik",
};

export default function Customers() {
  const [customers, setCustomers] = useState<CustomerCard[]>([]);
  const [filter, setFilter] = useState("");
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [name, setName] = useState("");
  const [codesText, setCodesText] = useState("");
  const [note, setNote] = useState("");
  const [err, setErr] = useState("");
  const [okMsg, setOkMsg] = useState("");
  const [busy, setBusy] = useState(false);
  const [zipResult, setZipResult] = useState<DesktopDropResult | null>(null);
  const [coverage, setCoverage] = useState<CustomerCoverage | null>(null);
  const web = isWebMode();

  async function refresh() {
    const r = await api.customersList();
    setCustomers(r.customers);
  }

  async function loadCoverage(id: string) {
    try {
      setCoverage(await api.customersCoverage(id));
    } catch {
      setCoverage(null);
    }
  }

  useEffect(() => {
    refresh().catch((e) => setErr(String(e)));
  }, []);

  const filtered = useMemo(() => {
    const q = filter.trim().toLowerCase();
    if (!q) return customers;
    return customers.filter(
      (c) =>
        c.name.toLowerCase().includes(q) ||
        (c.product_codes || []).some((code) => code.toLowerCase().includes(q)) ||
        (c.note || "").toLowerCase().includes(q),
    );
  }, [customers, filter]);

  function loadCard(c: CustomerCard) {
    setSelectedId(c.id);
    setName(c.name);
    setCodesText((c.product_codes || []).join("\n"));
    setNote(c.note || "");
    setZipResult(null);
    setOkMsg("");
    void loadCoverage(c.id);
  }

  function clearForm() {
    setSelectedId(null);
    setName("");
    setCodesText("");
    setNote("");
    setZipResult(null);
    setCoverage(null);
  }

  async function onSave(e: FormEvent) {
    e.preventDefault();
    setErr("");
    setOkMsg("");
    setBusy(true);
    try {
      const saved = await api.customersSave({
        name: name.trim(),
        codes_text: codesText,
        note: note.trim(),
        id: selectedId || undefined,
      });
      setSelectedId(saved.id);
      setOkMsg(`Kaydedildi: ${saved.name} · ${saved.product_codes.length} kod`);
      await refresh();
      await loadCoverage(saved.id);
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  async function onDelete() {
    if (!selectedId) return;
    setErr("");
    try {
      await api.customersDelete(selectedId);
      clearForm();
      await refresh();
    } catch (ex) {
      setErr(String(ex));
    }
  }

  async function onZip() {
    if (!selectedId) return;
    setErr("");
    setZipResult(null);
    setBusy(true);
    try {
      setZipResult(await api.customersZip(selectedId));
      await loadCoverage(selectedId);
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  async function onEnsure() {
    if (!selectedId) return;
    setErr("");
    setOkMsg("");
    setZipResult(null);
    setBusy(true);
    try {
      const r = await api.customersEnsure(selectedId, true);
      setOkMsg(
        `Eksikler: +${r.bulk.count_created} · atlandı ${r.bulk.count_skipped} · hata ${r.bulk.count_failed}`,
      );
      if (r.zip) setZipResult(r.zip);
      if (r.bulk.failed[0]) {
        setErr(r.bulk.failed.map((f) => `${f.product_code}: ${f.error}`).join(" · "));
      }
      await loadCoverage(selectedId);
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  return (
    <section>
      <p className="eyebrow">Müşteri · ZIP</p>
      <h1>Müşteri Paketi</h1>
      <p className="lead">
        Müşteri kartına ürün kodlarını yazın. ZIP yalnızca Technical File ve EU DoC PDF
        dosyalarını içerir. Kod listesini elle yapıştırmak için{" "}
        <Link to="/drop">Paket ZIP</Link>.
      </p>

      <div className="split">
        <div>
          <h2 className="section-title">Kartlar ({filtered.length}/{customers.length})</h2>
          <label className="block-label">
            Ara
            <input
              value={filter}
              onChange={(e) => setFilter(e.target.value)}
              placeholder="isim veya ürün kodu"
            />
          </label>
          <ul className="key-list">
            {filtered.map((c) => (
              <li key={c.id}>
                <button
                  type="button"
                  className={selectedId === c.id ? "picked-rev" : ""}
                  onClick={() => loadCard(c)}
                >
                  {c.name}
                  <span className="muted"> · {c.code_count} kod</span>
                </button>
              </li>
            ))}
            {filtered.length === 0 && (
              <li>
                <p className="muted" style={{ padding: "0.75rem" }}>
                  {customers.length === 0 ? "Henüz müşteri yok." : "Filtreye uyan kart yok."}
                </p>
              </li>
            )}
          </ul>
          <button type="button" style={{ marginTop: "0.75rem" }} onClick={clearForm}>
            Yeni kart
          </button>
        </div>

        <div className="detail">
          <form onSubmit={onSave}>
            <label className="block-label">
              Müşteri adı
              <input value={name} onChange={(e) => setName(e.target.value)} required />
            </label>
            <label className="block-label">
              Ürün kodları
              <textarea
                value={codesText}
                onChange={(e) => setCodesText(e.target.value)}
                rows={6}
                placeholder={"1000069\n1000070\n1000071"}
                required
              />
            </label>
            <label className="block-label">
              Not
              <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="isteğe bağlı" />
            </label>
            <div className="engine-actions">
              <button type="submit" disabled={busy}>
                {busy ? "…" : "Kaydet"}
              </button>
              <button type="button" disabled={!selectedId || busy} onClick={() => void onZip()}>
                {web ? "ZIP indir" : "Masaüstüne ZIP"}
              </button>
              <button type="button" disabled={!selectedId || busy} onClick={() => void onEnsure()}>
                Eksikleri üret + ZIP
              </button>
              <button
                type="button"
                className="btn-ghost"
                disabled={!selectedId}
                onClick={() => void onDelete()}
              >
                Sil
              </button>
            </div>
          </form>

          {coverage && (
            <div style={{ marginTop: "1rem" }}>
              <h3 className="section-title">Hazırlık Özeti</h3>
              <p className="meta">
                {coverage.ready}/{coverage.total} hazır
                {coverage.missing ? ` · eksik ${coverage.missing}` : ""}
                {coverage.incomplete ? ` · PDF eksik ${coverage.incomplete}` : ""}
                {coverage.zip_ready ? " · ZIP için uygun" : " · önce eksikleri tamamlayın"}
              </p>
              <ul className="activity-list">
                {coverage.rows.map((row) => (
                  <li key={row.product_code}>
                    <Link to={`/search?q=${encodeURIComponent(row.product_code)}&source=workspace`}>
                      <strong>{row.product_code}</strong>
                    </Link>
                    <span
                      className={
                        row.state === "ready"
                          ? "act-badge green"
                          : row.state === "incomplete"
                            ? "act-badge"
                            : "act-badge purple"
                      }
                      style={{ marginLeft: "0.4rem" }}
                    >
                      {STATE_TR[row.state] || row.state}
                    </span>
                    <span className="muted">
                      {" "}
                      ·{" "}
                      {row.state === "ready"
                        ? `${row.revision} · ${statusLabel(row.status)}`
                        : row.state === "incomplete"
                          ? `${row.revision || "—"} · PDF/DOCX eksik`
                          : "kayıt yok"}
                    </span>
                  </li>
                ))}
              </ul>
            </div>
          )}

          {err && <p className="error">{err}</p>}
          {okMsg && <p className="ok">{okMsg}</p>}
          {zipResult && (
            <div className="ok" style={{ marginTop: "0.75rem" }}>
              <p>
                ZIP hazır · yalnızca Technical File + EU DoC PDF · OK {zipResult.count_ok} · eksik{" "}
                {zipResult.count_missing}
                {zipResult.zip_name ? ` · ${zipResult.zip_name}` : ""}
              </p>
              {(zipResult.downloadHref || zipResult.download_url) && (
                <p style={{ marginTop: "0.4rem" }}>
                  <DownloadLink
                    href={zipResult.downloadHref || zipResult.download_url || ""}
                    label="ZIP’i indir"
                  />
                </p>
              )}
            </div>
          )}
        </div>
      </div>
    </section>
  );
}
