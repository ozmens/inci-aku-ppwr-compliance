import { FormEvent, useEffect, useState } from "react";
import { Link } from "react-router-dom";
import {
  api,
  BilingualSetPreview,
  PhotoResolveResult,
  PipelineStatus,
} from "../api";

import { withAccessToken } from "../download";
import { BILINGUAL_RULE_TR, measureLabel, SCOPE_OPTIONS } from "../labels";

const PHOTO_URL = (rel: string) =>
  withAccessToken(`/api/pipeline/photos/file?rel=${encodeURIComponent(rel)}`);

const RECENT_KEY = "inci_ppwr_pipeline_sets";

export default function Pipeline() {
  const [status, setStatus] = useState<PipelineStatus | null>(null);
  const [setCode, setSetCode] = useState("ST-021-STD-03");
  const [scope, setScope] = useState("starter");
  const [bilingual, setBilingual] = useState<BilingualSetPreview | null>(null);
  const [photos, setPhotos] = useState<PhotoResolveResult | null>(null);
  const [sampleTr, setSampleTr] = useState("KARTON SEPERATÖR");
  const [sampleEn, setSampleEn] = useState("");
  const [recents, setRecents] = useState<string[]>([]);
  const [err, setErr] = useState("");
  const [okMsg, setOkMsg] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    api.pipelineStatus().then(setStatus).catch((e) => setErr(String(e)));
    try {
      const raw = sessionStorage.getItem(RECENT_KEY);
      setRecents(raw ? (JSON.parse(raw) as string[]).slice(0, 6) : []);
    } catch {
      setRecents([]);
    }
  }, []);

  function pushRecent(code: string) {
    try {
      const next = [code, ...recents.filter((x) => x !== code)].slice(0, 6);
      sessionStorage.setItem(RECENT_KEY, JSON.stringify(next));
      setRecents(next);
    } catch {
      /* ignore */
    }
  }

  async function onPreview(e?: FormEvent) {
    e?.preventDefault();
    setErr("");
    setOkMsg("");
    setBusy(true);
    try {
      const code = setCode.trim();
      const [bi, ph] = await Promise.all([
        api.pipelineBilingualSet(code, scope),
        api.pipelineResolvePhotos(code, scope),
      ]);
      setBilingual(bi);
      setPhotos(ph);
      pushRecent(code);
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  async function onTranslateSample(e: FormEvent) {
    e.preventDefault();
    setErr("");
    try {
      const r = await api.pipelineBilingualTexts([sampleTr], "component");
      setSampleEn(r.items[0]?.en || "");
    } catch (ex) {
      setErr(String(ex));
    }
  }

  async function onExport() {
    setErr("");
    setOkMsg("");
    setBusy(true);
    try {
      const r = await api.pipelineExport(setCode.trim(), scope);
      setOkMsg(
        `Önizleme hazır · BOM ${r.bilingual_lines} satır · foto ${r.photos_matched}/${r.photos_matched + r.photos_missing}`,
      );
      pushRecent(setCode.trim());
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  return (
    <section>
      <p className="eyebrow">Çeviri · fotoğraf eşleme</p>
      <h1>Dil & Foto</h1>
      <p className="lead">
        Ambalaj BOM satırlarını Türkçe / İngilizce görün ve bileşen fotoğraflarını eşleyin.
        Dışa aktarım onaylı arşivi değiştirmez.
      </p>

      {status && (
        <div className="kpi-row kpi-row-3">
          <article className="kpi">
            <strong>{status.image_count}</strong>
            <span>Fotoğraf</span>
          </article>
          <article className="kpi">
            <strong>{status.mapping_count}</strong>
            <span>Eşleme satırı</span>
          </article>
          <article className="kpi pass">
            <strong>TR/EN</strong>
            <span>{BILINGUAL_RULE_TR}</span>
          </article>
        </div>
      )}
      {status && (
        <p className="meta">
          {status.image_count} fotoğraf · {status.mapping_count} eşleme
        </p>
      )}

      <form className="search-bar" onSubmit={onTranslateSample}>
        <input
          value={sampleTr}
          onChange={(e) => setSampleTr(e.target.value)}
          placeholder="Bileşen TR metni…"
        />
        <button type="submit">Çevir</button>
      </form>
      {sampleEn && (
        <p className="bilingual-sample">
          <span>{sampleTr}</span>
          <em>{sampleEn}</em>
        </p>
      )}

      {recents.length > 0 && (
        <div className="customer-chips" style={{ marginBottom: "0.75rem" }}>
          {recents.map((c) => (
            <button
              key={c}
              type="button"
              onClick={() => {
                setSetCode(c);
                void (async () => {
                  setSetCode(c);
                  setBusy(true);
                  setErr("");
                  try {
                    const [bi, ph] = await Promise.all([
                      api.pipelineBilingualSet(c, scope),
                      api.pipelineResolvePhotos(c, scope),
                    ]);
                    setBilingual(bi);
                    setPhotos(ph);
                  } catch (ex) {
                    setErr(String(ex));
                  } finally {
                    setBusy(false);
                  }
                })();
              }}
            >
              {c}
            </button>
          ))}
        </div>
      )}

      <form className="gap-form" onSubmit={onPreview}>
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
        <label className="grow">
          Ambalaj seti
          <input
            value={setCode}
            onChange={(e) => setSetCode(e.target.value)}
            placeholder="ST-021-STD-03"
            required
          />
        </label>
        <button type="submit" disabled={busy}>
          {busy ? "…" : "Önizle"}
        </button>
        <button type="button" disabled={busy || !setCode.trim()} onClick={() => void onExport()}>
          Adayı dışa aktar
        </button>
      </form>

      {err && <p className="error">{err}</p>}
      {okMsg && <p className="ok">{okMsg}</p>}

      {bilingual && (
        <div className="detail" style={{ marginTop: "1rem" }}>
          <h2>
            Çift dilli liste · {bilingual.set_code}{" "}
            <Link to={`/bom/${encodeURIComponent(bilingual.set_code)}`}>Ambalaj BOM</Link>
            {" · "}
            <Link to={`/dil-foto`}>Dil & Foto</Link>
          </h2>
          <p className="meta">{BILINGUAL_RULE_TR}</p>
          <table className="data-table">
            <thead>
              <tr>
                <th>Kod</th>
                <th>TR / EN</th>
                <th>Adet</th>
                <th>kg</th>
              </tr>
            </thead>
            <tbody>
              {bilingual.lines.map((line, i) => (
                <tr key={`${line.component_code}-${i}`}>
                  <td>
                    <strong>
                      <Link to={`/components?q=${encodeURIComponent(line.component_code)}`}>
                        {line.component_code}
                      </Link>
                    </strong>
                  </td>
                  <td className="bilingual-cell">
                    <span>{line.tr}</span>
                    <em>{line.en}</em>
                  </td>
                  <td>
                    {measureLabel(line.qty, "")} {measureLabel(line.uom, "")}
                  </td>
                  <td>{measureLabel(line.line_weight)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {bilingual.lines.length === 0 && <p className="muted">Bu set için ambalaj satırı yok.</p>}
        </div>
      )}

      {photos && (
        <div className="detail" style={{ marginTop: "1rem" }}>
          <h2>
            Fotoğraf eki ·{" "}
            <span className={photos.matched === photos.bom_line_count ? "act-badge green" : "act-badge purple"}>
              {photos.matched}/{photos.bom_line_count}
            </span>{" "}
            eşleşti
          </h2>
          {photos.missing.length > 0 && (
            <p className="meta">
              Eksik: {photos.missing.map((m) => m.component_code || m.description).join(", ")}
            </p>
          )}
          <div className="photo-grid">
            {photos.photos.map((p) => (
              <figure key={p.rel + p.component_code} className="photo-card">
                {p.exists ? (
                  <img src={PHOTO_URL(p.rel)} alt={p.name_tr} loading="lazy" />
                ) : (
                  <div className="photo-missing">yok</div>
                )}
                <figcaption>
                  <strong>
                    <Link to={`/components?q=${encodeURIComponent(p.component_code)}`}>
                      {p.component_code}
                    </Link>
                  </strong>
                  <span>{p.name_tr}</span>
                  <em>{p.name_en}</em>
                </figcaption>
              </figure>
            ))}
          </div>
          {photos.photos.length === 0 && (
            <p className="muted">Bu set için eşleşen temsilî foto yok.</p>
          )}
        </div>
      )}
    </section>
  );
}
