import { FormEvent, useEffect, useState } from "react";
import { Link, useNavigate, useParams, useSearchParams } from "react-router-dom";
import { api, RecordDetail, ScopeInfo } from "../api";
import FilePairList, { downloadKindLabel } from "../components/FilePairList";
import { useLastDownload } from "../components/useLastDownload";
import { isWebMode } from "../runtime";
import { scopeLabel } from "../labels";

const META: Record<
  string,
  { title: string; blurb: string; tone: string; searchHint: string }
> = {
  starter: {
    title: "STARTER",
    blurb: "Bireysel ambalaj — teknik dosya, beyan, etiket ve sevkiyat",
    tone: "",
    searchHint: "Starter ürün kodu…",
  },
  industrial: {
    title: "INDUSTRIAL",
    blurb: "Endüstriyel paketleme — seri üretim konfigürasyonları",
    tone: "purple",
    searchHint: "Industrial ürün / config kodu…",
  },
  container: {
    title: "CONTAINER",
    blurb: "Konteyner yükleme ve sevkiyat konfigürasyonları",
    tone: "green",
    searchHint: "Container config kodu…",
  },
  component: {
    title: "COMPONENT",
    blurb: "Bileşen ve yedek parça varyantları",
    tone: "amber",
    searchHint: "Component / variant kodu…",
  },
};

const DOC_FAMILY = [
  { code: "01", name: "Technical File", short: "TF" },
  { code: "02", name: "EU DoC", short: "DoC" },
  { code: "03", name: "Packaging Label", short: "Label" },
  { code: "04", name: "Shipment Statement", short: "STM" },
];

function fmt(n: number) {
  return new Intl.NumberFormat("tr-TR").format(n);
}

export default function Scopes() {
  const { scopeKey } = useParams();
  const navigate = useNavigate();
  const [searchParams, setSearchParams] = useSearchParams();
  const [scopes, setScopes] = useState<ScopeInfo[]>([]);
  const [err, setErr] = useState("");
  const [okMsg, setOkMsg] = useState("");
  const [busy, setBusy] = useState(false);
  const { capture, LastDownloadBar } = useLastDownload();
  const [previewKeys, setPreviewKeys] = useState<string[]>([]);
  const [draftQ, setDraftQ] = useState("");
  const [appliedQ, setAppliedQ] = useState("");
  const [totalKeys, setTotalKeys] = useState(0);
  const [activeKey, setActiveKey] = useState<string | null>(null);
  const [detail, setDetail] = useState<RecordDetail | null>(null);
  const [docFilter, setDocFilter] = useState<string | null>(null);
  const [recents, setRecents] = useState<string[]>([]);

  useEffect(() => {
    api
      .scopes()
      .then((r) => setScopes(r.scopes))
      .catch((e) => setErr(String(e)));
  }, []);

  const selected = scopeKey
    ? scopes.find((s) => s.key === scopeKey.toLowerCase())
    : null;
  const meta = selected ? META[selected.key] || META.starter : null;

  useEffect(() => {
    setDraftQ("");
    setAppliedQ("");
    setActiveKey(null);
    setDetail(null);
    const doc = searchParams.get("doc");
    setDocFilter(doc && ["01", "02", "03", "04"].includes(doc) ? doc : null);
    if (!selected?.key) {
      setRecents([]);
      return;
    }
    try {
      const raw = sessionStorage.getItem(`inci_ppwr_recent_${selected.key}`);
      setRecents(raw ? (JSON.parse(raw) as string[]).slice(0, 8) : []);
    } catch {
      setRecents([]);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selected?.key]);

  useEffect(() => {
    const doc = searchParams.get("doc");
    setDocFilter(doc && ["01", "02", "03", "04"].includes(doc) ? doc : null);
  }, [searchParams]);

  function setDoc(code: string | null) {
    setDocFilter(code);
    const next = new URLSearchParams(searchParams);
    if (code) next.set("doc", code);
    else next.delete("doc");
    setSearchParams(next, { replace: true });
  }

  function pushRecent(scope: string, key: string) {
    try {
      const storeKey = `inci_ppwr_recent_${scope}`;
      const prev = JSON.parse(sessionStorage.getItem(storeKey) || "[]") as string[];
      const next = [key, ...prev.filter((x) => x !== key)].slice(0, 8);
      sessionStorage.setItem(storeKey, JSON.stringify(next));
      setRecents(next);
    } catch {
      /* ignore */
    }
  }

  useEffect(() => {
    if (!selected) {
      setPreviewKeys([]);
      setTotalKeys(0);
      return;
    }
    api
      .keys(selected.key, appliedQ, 40)
      .then((r) => {
        setPreviewKeys(r.keys);
        setTotalKeys(r.total);
        if (r.keys[0] && (!activeKey || !r.keys.includes(activeKey))) {
          void loadRecord(selected.key, r.keys[0]);
        }
      })
      .catch(() => {
        setPreviewKeys([]);
        setTotalKeys(0);
      });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selected?.key, appliedQ]);

  async function loadRecord(scope: string, key: string) {
    setErr("");
    setOkMsg("");
    setActiveKey(key);
    try {
      setDetail(await api.record(scope, key));
      pushRecent(scope, key);
    } catch (ex) {
      setDetail(null);
      setErr(String(ex));
    }
  }

  function onLocalSearch(e: FormEvent) {
    e.preventDefault();
    setAppliedQ(draftQ.trim());
  }

  if (scopeKey && scopes.length > 0 && !selected) {
    return (
      <section>
        <p className="eyebrow">Doküman Merkezi</p>
        <h1>Kapsam Bulunamadı</h1>
        <p className="lead">
          <code>{scopeKey}</code> geçerli değil.{" "}
          <Link to="/scopes">Doküman Merkezi’ne dön</Link>
        </p>
      </section>
    );
  }

  if (selected && meta) {
    return (
      <section>
        <p className="eyebrow">
          <Link to="/scopes">Doküman Merkezi</Link> · onaylı arşiv
        </p>
        <div className="scope-switcher">
          {(["starter", "industrial", "container", "component"] as const).map((k) => {
            const sc = scopes.find((s) => s.key === k);
            const m = META[k];
            return (
              <button
                key={k}
                type="button"
                className={`scope-switch ${selected.key === k ? "active" : ""} ${m.tone}`}
                onClick={() => navigate(`/scopes/${k}`)}
              >
                {m.title}
                <em>{sc ? fmt(sc.records) : "—"}</em>
              </button>
            );
          })}
        </div>

        <h1>{meta.title}</h1>
        <p className="lead">{meta.blurb}</p>
        <p className="meta">
          {fmt(selected.records)} kayıt · {selected.exists ? "hazır" : "eksik"}
        </p>

        <div className="doc-strip" style={{ marginTop: "1rem" }}>
          {DOC_FAMILY.map((d) => (
            <button
              key={d.code}
              type="button"
              className={`doc-chip doc-chip-btn ${docFilter === d.code ? "active" : ""}`}
              onClick={() => setDoc(docFilter === d.code ? null : d.code)}
            >
              <strong>
                {d.code} · {d.short}
              </strong>
              <span>{d.name}</span>
            </button>
          ))}
        </div>

        <div className="engine-actions" style={{ marginTop: "1rem" }}>
          {!isWebMode() && (
            <button
              type="button"
              disabled={!selected.launcher || busy}
              onClick={() => {
                setBusy(true);
                setOkMsg("");
                api
                  .openEngine(selected.key)
                  .then(() => setOkMsg("Doküman indeksi açıldı"))
                  .catch((e) => setErr(String(e)))
                  .finally(() => setBusy(false));
              }}
            >
              Doküman indeksini aç
            </button>
          )}
          <button
            type="button"
            onClick={() =>
              navigate(`/search?source=${encodeURIComponent(selected.key)}`)
            }
          >
            Gelişmiş arama
          </button>
          {selected.key === "starter" && (
            <button type="button" onClick={() => navigate("/bom")}>
              Ambalaj BOM
            </button>
          )}
          <button type="button" className="btn-ghost" onClick={() => navigate("/scopes")}>
            Tüm kapsamlar
          </button>
        </div>

        {err && <p className="error">{err}</p>}
        {okMsg && <p className="ok">{okMsg}</p>}
        <LastDownloadBar />

        <div className="section-head" style={{ marginTop: "1.5rem" }}>
          <h2>Doküman Ailesi</h2>
          <span className="muted">{totalKeys} sonuç</span>
        </div>
        <form className="search-bar" onSubmit={onLocalSearch}>
          <input
            value={draftQ}
            onChange={(e) => setDraftQ(e.target.value)}
            placeholder={meta.searchHint}
          />
          <button type="submit">Filtrele</button>
        </form>

        <div className="split">
          <div>
            {recents.length > 0 && (
              <div style={{ marginBottom: "0.75rem" }}>
                <h3 className="section-title">Son Bakılan</h3>
                <div className="customer-chips">
                  {recents.map((k) => (
                    <button key={k} type="button" onClick={() => void loadRecord(selected.key, k)}>
                      {k}
                    </button>
                  ))}
                </div>
              </div>
            )}
            <ul className="key-list">
              {previewKeys.map((k) => (
                <li key={k}>
                  <button
                    type="button"
                    className={activeKey === k ? "picked-rev" : ""}
                    onClick={() => void loadRecord(selected.key, k)}
                  >
                    {k}
                  </button>
                </li>
              ))}
              {previewKeys.length === 0 && (
                <li>
                  <p className="muted" style={{ padding: "0.75rem" }}>
                    Bu filtreyle kayıt yok.
                  </p>
                </li>
              )}
            </ul>
          </div>

          <div className="detail">
            {!detail && <p className="muted">Listeden bir kayıt seçin — Word ve PDF burada açılır.</p>}
            {detail && (
              <>
                <h2>{detail.key}</h2>
                <p className="meta">
                  Kapsam: {scopeLabel(detail.scope)}
                  {docFilter ? ` · filtre ${docFilter}` : ""}
                </p>
                <FilePairList
                  files={detail.files.filter((f) => !docFilter || f.name.startsWith(docFilter))}
                  onOpen={(f) => {
                    setOkMsg("");
                    api
                      .openFile(selected.key, detail.key, f.name)
                      .then((r) => {
                        capture(r.download_url, downloadKindLabel(f.kind, f.name));
                        setOkMsg(`${f.kind === "WORD" ? "WORD" : "PDF"} · ${f.label}`);
                      })
                      .catch((e) => setErr(String(e)));
                  }}
                />
              </>
            )}
          </div>
        </div>
      </section>
    );
  }

  return (
    <section>
      <p className="eyebrow">Doküman Merkezi</p>
      <h1>Kapsamlar</h1>
      <p className="lead">
        STARTER · INDUSTRIAL · CONTAINER · COMPONENT. Bir kapsam seçin, ürün kodunu arayın ve
        Word / PDF dosyalarını açın.
      </p>
      {err && <p className="error">{err}</p>}
      <div className="scope-enter-grid">
        {(["starter", "industrial", "container", "component"] as const).map((k) => {
          const s = scopes.find((x) => x.key === k);
          const m = META[k];
          return (
            <button
              key={k}
              type="button"
              className={`scope-enter-card ${m.tone}`}
              disabled={!s}
              onClick={() => navigate(`/scopes/${k}`)}
            >
              <span className="scope-enter-label">{m.title}</span>
              <strong>{s ? fmt(s.records) : "…"}</strong>
              <span className="muted">{m.blurb}</span>
              <span className="scope-enter-go">İçeri gir ›</span>
            </button>
          );
        })}
      </div>
    </section>
  );
}
