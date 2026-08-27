import { FormEvent, useEffect, useState } from "react";
import { Link } from "react-router-dom";
import { api, WsProduct, WsProductDetail } from "../api";
import FilePairList, { downloadKindLabel } from "../components/FilePairList";
import { useLastDownload } from "../components/useLastDownload";
import { isWebMode } from "../runtime";
import { reasonLabel, SCOPE_OPTIONS, statusLabel } from "../labels";

type WsEngineInfo = {
  delivery_root: string;
  engine_exists: boolean;
  launcher_exists: boolean;
  complete_products: number;
  link_rule: string;
};

export default function Workspace() {
  const [products, setProducts] = useState<WsProduct[]>([]);
  const [selected, setSelected] = useState<WsProductDetail | null>(null);
  const [productCode, setProductCode] = useState("");
  const [description, setDescription] = useState("");
  const [setCode, setSetCode] = useState("");
  const [scope, setScope] = useState("starter");
  const [reason, setReason] = useState("İlk yayın");
  const [reviseReason, setReviseReason] = useState("");
  const [err, setErr] = useState("");
  const [okMsg, setOkMsg] = useState("");
  const [busy, setBusy] = useState(false);
  const [viewRev, setViewRev] = useState<string | null>(null);
  const [bulkText, setBulkText] = useState("");
  const [engine, setEngine] = useState<WsEngineInfo | null>(null);
  const [listFilter, setListFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState<string | null>(null);
  const { capture, LastDownloadBar } = useLastDownload();

  function matchesListFilter(p: WsProduct) {
    if (statusFilter && (p.status || "").toUpperCase() !== statusFilter) return false;
    const q = listFilter.trim().toLowerCase();
    if (!q) return true;
    return (
      p.product_code.toLowerCase().includes(q) ||
      (p.current_revision || "").toLowerCase().includes(q) ||
      (p.status || "").toLowerCase().includes(q) ||
      (p.description || "").toLowerCase().includes(q) ||
      (p.set_code || "").toLowerCase().includes(q)
    );
  }

  const filteredProducts = products.filter(matchesListFilter);

  async function refresh() {
    const [r, e] = await Promise.all([api.wsProducts(), api.wsEngineStatus()]);
    setProducts(r.products);
    setEngine(e);
  }

  useEffect(() => {
    refresh().catch((e) => setErr(String(e)));
  }, []);

  async function onRebuildEngine() {
    setErr("");
    setOkMsg("");
    setBusy(true);
    try {
      const r = await api.wsEngineRebuild();
      setOkMsg(`Doküman indeksi güncellendi · ${r.records} kayıt`);
      await refresh();
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  async function load(code: string) {
    setErr("");
    const d = await api.wsProduct(code);
    setSelected(d);
    setProductCode(d.product.product_code);
    setDescription(String(d.product.description || ""));
    setSetCode(String(d.product.set_code || ""));
    setViewRev(d.product.current_revision || null);
  }

  async function onBulk(e: FormEvent) {
    e.preventDefault();
    setErr("");
    setOkMsg("");
    setBusy(true);
    try {
      const r = await api.wsBulkCreate(bulkText, scope);
      setOkMsg(
        `Toplu: +${r.count_created} · atlandı ${r.count_skipped} · hata ${r.count_failed}`,
      );
      if (r.failed[0]) setErr(r.failed.map((f) => `${f.product_code}: ${f.error}`).join(" · "));
      await refresh();
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  async function onCreate(e: FormEvent) {
    e.preventDefault();
    setErr("");
    setOkMsg("");
    setBusy(true);
    try {
      const r = await api.wsCreate({
        product_code: productCode.trim(),
        description: description.trim(),
        set_code: setCode.trim(),
        scope,
        reason: reason.trim() || "İlk yayın",
      });
      setOkMsg(
        `${r.product.product_code} · ${r.revision.revision} · ${statusLabel(r.revision.status)}`,
      );
      await refresh();
      await load(r.product.product_code);
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  async function onRevise(e: FormEvent) {
    e.preventDefault();
    if (!selected) return;
    setErr("");
    setOkMsg("");
    setBusy(true);
    try {
      const r = await api.wsRevise(selected.product.product_code, {
        reason: reviseReason.trim(),
        description: description.trim(),
        set_code: setCode.trim(),
        scope,
      });
      setOkMsg(
        `Revize: ${r.superseded || "—"} → ${r.revision.revision} · ${statusLabel(r.revision.status)}`,
      );
      setReviseReason("");
      await refresh();
      await load(selected.product.product_code);
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  const files =
    selected &&
    (viewRev && viewRev !== selected.product.current_revision
      ? selected.revisions.find((r) => r.revision === viewRev)?.files
      : selected.current_files);

  return (
    <section>
      <p className="eyebrow">Resmi kaynak · revizyon</p>
      <h1>Revizyon Yönetimi</h1>
      <p className="lead">
        Ürün paketlerinin resmi kaynağı. Taslak, yayın ve revizyon adımlarını yönetin; Word ve PDF
        müşteri paketine hazır hale getirin. Toplu indirme: <Link to="/drop">Paket ZIP</Link>.
      </p>

      {engine && (
        <div className="detail" style={{ marginBottom: "1rem" }}>
          <h2 className="section-title">Doküman indeksi</h2>
          <p className="meta">
            {engine.complete_products} tam paket
            {engine.engine_exists ? " · indeks hazır" : ""}
          </p>
          <div className="engine-actions">
            <button type="button" disabled={busy} onClick={() => void onRebuildEngine()}>
              İndeksi yenile
            </button>
            {!isWebMode() && (
              <>
                <button
                  type="button"
                  disabled={!engine.launcher_exists}
                  onClick={() => api.wsEngineOpen().catch((e) => setErr(String(e)))}
                >
                  İndeksi aç
                </button>
                <button
                  type="button"
                  onClick={() => api.wsEngineOpenFolder().catch((e) => setErr(String(e)))}
                >
                  Klasörü aç
                </button>
              </>
            )}
          </div>
        </div>
      )}

      <form className="pack-form" onSubmit={onCreate}>
        <label>
          Ürün kodu
          <input value={productCode} onChange={(e) => setProductCode(e.target.value)} required />
        </label>
        <label className="grow">
          Açıklama
          <input value={description} onChange={(e) => setDescription(e.target.value)} />
        </label>
        <label>
          Set
          <input value={setCode} onChange={(e) => setSetCode(e.target.value)} placeholder="ST-…" />
        </label>
        <label>
          Kapsam
          <select value={scope} onChange={(e) => setScope(e.target.value)}>
            {SCOPE_OPTIONS.filter((o) => o.value === "starter" || o.value === "industrial").map(
              (o) => (
                <option key={o.value} value={o.value}>
                  {o.label}
                </option>
              ),
            )}
          </select>
        </label>
        <label className="grow">
          İlk gerekçe
          <input value={reason} onChange={(e) => setReason(e.target.value)} />
        </label>
        <button type="submit" disabled={busy}>
          {busy ? "…" : "Yeni paket (Rev.00)"}
        </button>
      </form>

      <form className="bulk-form" onSubmit={onBulk}>
        <label className="block-label">
          Toplu master import (kod listesi → Rev.00)
          <textarea
            value={bulkText}
            onChange={(e) => setBulkText(e.target.value)}
            rows={4}
            placeholder={"1000069\n1000070\n1000071"}
          />
        </label>
        <button type="submit" disabled={busy || !bulkText.trim()}>
          {busy ? "…" : "Eksik olanları üret"}
        </button>
        <button
          type="button"
          disabled={busy}
          style={{ marginLeft: "0.5rem" }}
          onClick={() => {
            setBusy(true);
            setErr("");
            api
              .wsCompleteAll()
              .then((r) => {
                setOkMsg(`PDF tamamla: ${r.completed}/${r.count} paket`);
                return refresh();
              })
              .catch((e) => setErr(String(e)))
              .finally(() => setBusy(false));
          }}
        >
          Eksik PDF’leri tamamla
        </button>
      </form>

      {err && <p className="error">{err}</p>}
      {okMsg && <p className="ok">{okMsg}</p>}

      <div className="split" style={{ marginTop: "1rem" }}>
        <div>
          <h2 className="section-title">
            Ürünler ({filteredProducts.length}
            {listFilter.trim() ? `/${products.length}` : ""})
          </h2>
          <form
            className="inline"
            style={{ marginBottom: "0.75rem" }}
            onSubmit={(e) => e.preventDefault()}
          >
            <input
              value={listFilter}
              onChange={(e) => setListFilter(e.target.value)}
              placeholder="Kod / rev / durum / set filtre…"
              aria-label="Ürün listesi filtresi"
            />
            {listFilter && (
              <button type="button" className="btn-ghost" onClick={() => setListFilter("")}>
                Temizle
              </button>
            )}
          </form>
          <div className="customer-chips" style={{ marginBottom: "0.75rem" }}>
            {([null, "DRAFT", "ISSUED", "SUPERSEDED"] as const).map((s) => (
              <button
                key={s || "ALL"}
                type="button"
                className={statusFilter === s ? "picked-rev" : ""}
                onClick={() => setStatusFilter(s)}
              >
                {s ? statusLabel(s) : "Tümü"}
              </button>
            ))}
          </div>
          <ul className="key-list">
            {filteredProducts.map((p) => (
              <li key={p.product_code}>
                <button
                  type="button"
                  className={selected?.product.product_code === p.product_code ? "picked-rev" : ""}
                  onClick={() => void load(p.product_code).catch((e) => setErr(String(e)))}
                >
                  {p.product_code}
                  <span className="muted"> · {p.current_revision}</span>
                  <span
                    className={
                      p.status === "ISSUED"
                        ? "act-badge green"
                        : p.status === "SUPERSEDED"
                          ? "act-badge purple"
                          : "act-badge"
                    }
                    style={{ marginLeft: "0.35rem" }}
                  >
                    {statusLabel(p.status)}
                  </span>
                  {!p.complete && (
                    <span className="act-badge purple" style={{ marginLeft: "0.25rem" }}>
                      eksik
                    </span>
                  )}
                </button>
              </li>
            ))}
            {products.length === 0 && (
              <li>
                <p className="muted" style={{ padding: "0.75rem" }}>
                  Henüz ürün yok.
                </p>
              </li>
            )}
            {products.length > 0 && filteredProducts.length === 0 && (
              <li>
                <p className="muted" style={{ padding: "0.75rem" }}>
                  Filtreyle eşleşen ürün yok.
                </p>
              </li>
            )}
          </ul>
        </div>
        <div className="detail">
          {!selected && <p className="muted">Ürün seçin veya yeni paket oluşturun</p>}
          {selected && (
            <>
              <h2>
                {selected.product.product_code}{" "}
                <span
                  className={
                    selected.product.status === "ISSUED"
                      ? "act-badge green"
                      : selected.product.status === "SUPERSEDED"
                        ? "act-badge purple"
                        : "act-badge"
                  }
                >
                  {statusLabel(selected.product.status)}
                </span>
              </h2>
              <p className="meta">{selected.product.description}</p>
              <p className="meta">
                Güncel: {selected.product.current_revision} · Set:{" "}
                {selected.product.set_code ? (
                  <Link to={`/bom/${encodeURIComponent(selected.product.set_code)}`}>
                    {selected.product.set_code}
                  </Link>
                ) : (
                  "—"
                )}
                {" · "}
                <Link to={`/customers`}>Müşteri</Link>
                {" · "}
                <Link to={`/drop`}>Paket ZIP</Link>
              </p>

              <h3 className="section-title">Revizyonlar</h3>
              <ul className="rev-list">
                {selected.revisions.map((r) => (
                  <li key={r.revision}>
                    <button
                      type="button"
                      className={viewRev === r.revision ? "picked-rev" : ""}
                      onClick={() => setViewRev(r.revision)}
                    >
                      {r.revision} · {statusLabel(r.status)}
                      {r.reason ? ` — ${reasonLabel(r.reason)}` : ""}
                    </button>
                  </li>
                ))}
              </ul>

              <form className="revise-form" onSubmit={onRevise}>
                <label className="grow">
                  Revizyon gerekçesi (zorunlu)
                  <input
                    value={reviseReason}
                    onChange={(e) => setReviseReason(e.target.value)}
                    placeholder="örn. BOM / etiket düzeltmesi"
                    required
                  />
                </label>
                <button type="submit" disabled={busy || reviseReason.trim().length < 3}>
                  Yeni revizyon üret
                </button>
              </form>

              <h3 className="section-title">Dosyalar ({viewRev || "güncel"})</h3>
              <LastDownloadBar />
              <FilePairList
                files={files || []}
                onOpen={(f) =>
                  api
                    .wsOpen(selected.product.product_code, f.name, viewRev || undefined)
                    .then((r) => capture(r.download_url, downloadKindLabel(f.kind, f.name)))
                    .catch((e) => setErr(String(e)))
                }
              />
              {selected.product.status === "DRAFT" ||
              (files || []).some((f) => !f.exists) ? (
                <button
                  type="button"
                  style={{ marginTop: "0.75rem" }}
                  disabled={busy}
                  onClick={() => {
                    setBusy(true);
                    api
                      .wsCompletePdfs(selected.product.product_code, viewRev || undefined)
                      .then(async (r) => {
                        setOkMsg(
                          `PDF: ${r.product_code} ${r.revision} → ${statusLabel(r.status)}`,
                        );
                        await load(selected.product.product_code);
                        await refresh();
                      })
                      .catch((e) => setErr(String(e)))
                      .finally(() => setBusy(false));
                  }}
                >
                  Bu paketin PDF’lerini tamamla
                </button>
              ) : null}
            </>
          )}
        </div>
      </div>
    </section>
  );
}
