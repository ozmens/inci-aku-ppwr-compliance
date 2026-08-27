import { FormEvent, useEffect, useMemo, useState } from "react";
import { Link, useSearchParams } from "react-router-dom";
import { api, CandidatePack, PackBuildResult } from "../api";
import FilePairList, { downloadKindLabel } from "../components/FilePairList";
import { useLastDownload } from "../components/useLastDownload";
import { isWebMode } from "../runtime";
import { SCOPE_OPTIONS } from "../labels";

type EngineInfo = {
  delivery_root: string;
  engine: string | null;
  engine_exists: boolean;
  launcher: string | null;
  launcher_exists: boolean;
  packs: number;
  complete_packs: number;
  link_rule: string;
};

export default function PackBuilder() {
  const [searchParams] = useSearchParams();
  const [packs, setPacks] = useState<CandidatePack[]>([]);
  const [productCode, setProductCode] = useState("");
  const [description, setDescription] = useState("");
  const [setCode, setSetCode] = useState("");
  const [scope, setScope] = useState("starter");
  const [skipPdf, setSkipPdf] = useState(false);
  const [result, setResult] = useState<PackBuildResult | null>(null);
  const [selected, setSelected] = useState<CandidatePack | null>(null);
  const [engine, setEngine] = useState<EngineInfo | null>(null);
  const [err, setErr] = useState("");
  const [okMsg, setOkMsg] = useState("");
  const [busy, setBusy] = useState(false);
  const [listFilter, setListFilter] = useState("");
  const [completeOnly, setCompleteOnly] = useState<"all" | "ok" | "gap">("all");
  const { capture, LastDownloadBar } = useLastDownload();

  async function refresh() {
    const [r, e] = await Promise.all([api.packsList(), api.engineStatus()]);
    setPacks(r.packs);
    setEngine(e);
    return r.packs;
  }

  useEffect(() => {
    const code = searchParams.get("code") || "";
    const set = searchParams.get("set") || "";
    const desc = searchParams.get("desc") || "";
    if (code) setProductCode(code);
    if (set) setSetCode(set);
    if (desc) setDescription(desc);
  }, [searchParams]);

  useEffect(() => {
    refresh()
      .then((list) => {
        const pref = searchParams.get("code");
        if (pref && list.some((p) => p.product_code === pref)) {
          void loadPack(pref);
        } else if (list[0]) {
          void loadPack(list[0].product_code);
        }
      })
      .catch((e) => setErr(String(e)));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const filteredPacks = useMemo(() => {
    const q = listFilter.trim().toLowerCase();
    return packs.filter((p) => {
      if (completeOnly === "ok" && !p.complete) return false;
      if (completeOnly === "gap" && p.complete) return false;
      if (!q) return true;
      const set = String(p.meta?.set_code || "").toLowerCase();
      const desc = String(p.meta?.description || "").toLowerCase();
      return (
        p.product_code.toLowerCase().includes(q) ||
        set.includes(q) ||
        desc.includes(q)
      );
    });
  }, [packs, listFilter, completeOnly]);

  async function onBuild(e: FormEvent) {
    e.preventDefault();
    setErr("");
    setOkMsg("");
    setResult(null);
    setBusy(true);
    try {
      const r = await api.packsBuild({
        product_code: productCode.trim(),
        description: description.trim(),
        set_code: setCode.trim(),
        scope,
        skip_pdf: skipPdf,
      });
      setResult(r);
      setSelected({
        product_code: r.pack.product_code,
        path: r.folder,
        meta: r.pack,
        files: r.files,
        complete: r.complete,
        docx: r.files.filter((f) => f.kind === "WORD" && f.exists).length,
        pdf: r.files.filter((f) => f.kind === "PDF" && f.exists).length,
      });
      if (r.complete) setOkMsg("Paket hazır · doküman indeksi güncellendi");
      await refresh();
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  async function loadPack(code: string) {
    setErr("");
    try {
      const p = await api.packsGet(code);
      setSelected({
        product_code: p.product_code,
        path: p.folder,
        meta: p.meta,
        files: p.files,
        complete: p.complete,
        docx: p.files.filter((f) => f.kind === "WORD" && f.exists).length,
        pdf: p.files.filter((f) => f.kind === "PDF" && f.exists).length,
      });
      setProductCode(p.product_code);
      setDescription(String(p.meta?.description || ""));
      setSetCode(String(p.meta?.set_code || ""));
    } catch (ex) {
      setErr(String(ex));
    }
  }

  async function onDelete(code: string) {
    setErr("");
    try {
      await api.packsDelete(code);
      if (selected?.product_code === code) setSelected(null);
      await refresh();
    } catch (ex) {
      setErr(String(ex));
    }
  }

  async function onRebuildEngine() {
    setErr("");
    setOkMsg("");
    setBusy(true);
    try {
      const r = await api.engineRebuild();
      setOkMsg(`Doküman indeksi güncellendi · ${r.records} kayıt`);
      await refresh();
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  return (
    <section>
      <p className="eyebrow">Aday doküman paketi</p>
      <h1>Aday Paket</h1>
      <p className="lead">
        Eksik ürün için 4 Word + 4 PDF paketini üretin. PDF otomatik üretilir. Hazır paketi
        revizyon yönetimine alıp müşteri ZIP’ine ekleyebilirsiniz.
      </p>

      <div className="banner-safe">
        <strong>Onaylı arşiv korunur</strong>
        <span>Çıktılar aday çalışma alanına yazılır</span>
        {!isWebMode() && (
          <button type="button" onClick={() => api.packsOpenFolder().catch((e) => setErr(String(e)))}>
            Klasörü aç
          </button>
        )}
      </div>

      {engine && (
        <div className="detail" style={{ marginBottom: "1rem" }}>
          <h2 className="section-title">Doküman indeksi</h2>
          <p className="meta">
            {engine.complete_packs}/{engine.packs} tam paket
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
                  onClick={() => api.engineOpen().catch((e) => setErr(String(e)))}
                >
                  İndeksi aç
                </button>
                <button type="button" onClick={() => api.engineOpenFolder().catch((e) => setErr(String(e)))}>
                  Klasörü aç
                </button>
              </>
            )}
          </div>
        </div>
      )}

      <form className="pack-form" onSubmit={onBuild}>
        <label>
          Ürün kodu
          <input
            value={productCode}
            onChange={(e) => setProductCode(e.target.value)}
            placeholder="örn. 1009999"
            required
          />
        </label>
        <label className="grow">
          Açıklama
          <input
            value={description}
            onChange={(e) => setDescription(e.target.value)}
            placeholder="Master’dan otomatik veya elle…"
          />
        </label>
        <label>
          Set kodu
          <input
            value={setCode}
            onChange={(e) => setSetCode(e.target.value)}
            placeholder="ST-021-STD-03"
          />
        </label>
        <label>
          Kapsam
          <select value={scope} onChange={(e) => setScope(e.target.value)}>
            {SCOPE_OPTIONS.map((o) => (
              <option key={o.value} value={o.value}>
                {o.label}
              </option>
            ))}
          </select>
        </label>
        <label className="check">
          <input type="checkbox" checked={skipPdf} onChange={(e) => setSkipPdf(e.target.checked)} />
          Önce yalnızca Word (PDF sonra)
        </label>
        <button type="submit" disabled={busy}>
          {busy ? "Üretiliyor…" : "Paket üret"}
        </button>
      </form>
      <p className="meta">Master ürün kodu girerseniz açıklama ve set boş kalabilir.</p>

      {err && <p className="error">{err}</p>}
      {okMsg && <p className="ok">{okMsg}</p>}
      {result && (
        <p className={result.complete ? "ok" : "error"}>
          {result.complete ? "Tam paket (4 Word + 4 PDF)" : "Eksik dosya var — Word veya PDF tamamlanmadı"}
        </p>
      )}

      <div className="split" style={{ marginTop: "1.25rem" }}>
        <div>
          <h2 className="section-title">
            Aday paketler ({filteredPacks.length}
            {listFilter || completeOnly !== "all" ? `/${packs.length}` : ""})
          </h2>
          <form className="inline" style={{ marginBottom: "0.75rem" }} onSubmit={(e) => e.preventDefault()}>
            <input
              value={listFilter}
              onChange={(e) => setListFilter(e.target.value)}
              placeholder="Kod / set / açıklama…"
              aria-label="Aday paket filtresi"
            />
          </form>
          <div className="customer-chips" style={{ marginBottom: "0.75rem" }}>
            {(
              [
                ["all", "Tümü"],
                ["ok", "Tam"],
                ["gap", "Eksik"],
              ] as const
            ).map(([k, label]) => (
              <button
                key={k}
                type="button"
                className={completeOnly === k ? "picked-rev" : ""}
                onClick={() => setCompleteOnly(k)}
              >
                {label}
              </button>
            ))}
          </div>
          <ul className="key-list">
            {filteredPacks.map((p) => (
              <li key={p.product_code}>
                <button
                  type="button"
                  className={selected?.product_code === p.product_code ? "picked-rev" : ""}
                  onClick={() => void loadPack(p.product_code)}
                >
                  {p.product_code}
                  <span className="muted">
                    {" "}
                    · {p.docx}W/{p.pdf}P
                  </span>
                  <span
                    className={p.complete ? "act-badge green" : "act-badge purple"}
                    style={{ marginLeft: "0.35rem" }}
                  >
                    {p.complete ? "tam" : "eksik"}
                  </span>
                </button>
              </li>
            ))}
            {packs.length === 0 && (
              <li>
                <p className="muted" style={{ padding: "0.75rem" }}>
                  Henüz aday paket yok.
                </p>
              </li>
            )}
            {packs.length > 0 && filteredPacks.length === 0 && (
              <li>
                <p className="muted" style={{ padding: "0.75rem" }}>
                  Filtreyle eşleşen paket yok.
                </p>
              </li>
            )}
          </ul>
        </div>
        <div className="detail">
          {!selected && <p className="muted">Paket seçin veya üretin — Word ve PDF burada açılır.</p>}
          {selected && (
            <>
              <h2>
                {selected.product_code}{" "}
                <span className={selected.complete ? "act-badge green" : "act-badge purple"}>
                  {selected.complete ? "tam" : "eksik"}
                </span>
              </h2>
              <p className="meta">
                <Link to={`/workspace`}>Revizyon</Link>
                {" · "}
                <Link to={`/search?q=${encodeURIComponent(selected.product_code)}&source=workspace`}>
                  Ara
                </Link>
              </p>
              <dl className="facts">
                <div>
                  <dt>Set</dt>
                  <dd>
                    {selected.meta?.set_code ? (
                      <Link to={`/bom/${encodeURIComponent(String(selected.meta.set_code))}`}>
                        {String(selected.meta.set_code)}
                      </Link>
                    ) : (
                      "—"
                    )}
                  </dd>
                </div>
                <div>
                  <dt>Dara (kg)</dt>
                  <dd>{String(selected.meta?.tare_kg ?? "—")}</dd>
                </div>
                <div>
                  <dt>Fotoğraf</dt>
                  <dd>{String(selected.meta?.photos ?? "—")}</dd>
                </div>
              </dl>
              <FilePairList
                files={selected.files}
                onOpen={(f) =>
                  api
                    .packsOpen(selected.product_code, f.name)
                    .then((r) => capture(r.download_url, downloadKindLabel(f.kind, f.name)))
                    .catch((e) => setErr(String(e)))
                }
              />
              <LastDownloadBar />
              <button
                type="button"
                className="btn-ghost"
                style={{ marginTop: "0.75rem" }}
                onClick={() => void onDelete(selected.product_code)}
              >
                Aday paketi sil
              </button>
            </>
          )}
        </div>
      </div>
    </section>
  );
}
