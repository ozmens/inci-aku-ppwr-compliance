import { FormEvent, useEffect, useState } from "react";
import { Link, useSearchParams } from "react-router-dom";
import { api, RecordDetail, WsProduct, WsProductDetail } from "../api";
import FilePairList, { downloadKindLabel } from "../components/FilePairList";
import { useLastDownload } from "../components/useLastDownload";
import { scopeLabel, statusLabel } from "../labels";

const FROZEN = ["starter", "industrial", "container", "component"] as const;
type Source = "workspace" | (typeof FROZEN)[number];

export default function Search() {
  const [params] = useSearchParams();
  const [source, setSource] = useState<Source>("workspace");
  const [q, setQ] = useState(() => params.get("q") || "");
  const [wsHits, setWsHits] = useState<WsProduct[]>([]);
  const [frozenKeys, setFrozenKeys] = useState<string[]>([]);
  const [total, setTotal] = useState(0);
  const [wsDetail, setWsDetail] = useState<WsProductDetail | null>(null);
  const [frozenDetail, setFrozenDetail] = useState<RecordDetail | null>(null);
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const { capture, LastDownloadBar } = useLastDownload();

  async function runSearch(query = q, src: Source = source) {
    setErr("");
    setWsDetail(null);
    setFrozenDetail(null);
    setBusy(true);
    try {
      if (src === "workspace") {
        const r = await api.wsSearch(query);
        setWsHits(r.products);
        setFrozenKeys([]);
        setTotal(r.total);
        if (r.products[0]) {
          void selectWs(r.products[0].product_code);
        }
        // Product code often lives in frozen starter; fall through if empty.
        if (query.trim() && r.total === 0) {
          const fr = await api.keys("starter", query);
          if (fr.total > 0) {
            setSource("starter");
            setFrozenKeys(fr.keys);
            setWsHits([]);
            setTotal(fr.total);
            if (fr.keys[0]) void selectFrozenKey("starter", fr.keys[0]);
          }
        }
      } else {
        const r = await api.keys(src, query);
        setFrozenKeys(r.keys);
        setWsHits([]);
        setTotal(r.total);
        if (r.keys[0]) void selectFrozenKey(src, r.keys[0]);
      }
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  useEffect(() => {
    const fromUrl = params.get("q") || "";
    const srcParam = params.get("source") || "";
    const allowed = ["workspace", "starter", "industrial", "container", "component"] as const;
    if (srcParam && (allowed as readonly string[]).includes(srcParam) && srcParam !== source) {
      setSource(srcParam as Source);
      return;
    }
    if (fromUrl && fromUrl !== q) setQ(fromUrl);
    void runSearch(fromUrl || q, source);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [source, params]);

  async function onSearch(e: FormEvent) {
    e.preventDefault();
    await runSearch(q, source);
  }

  async function selectWs(code: string) {
    setErr("");
    setFrozenDetail(null);
    try {
      setWsDetail(await api.wsProduct(code));
    } catch (ex) {
      setErr(String(ex));
    }
  }

  async function selectFrozenKey(scope: string, key: string) {
    setErr("");
    setWsDetail(null);
    try {
      setFrozenDetail(await api.record(scope as (typeof FROZEN)[number], key));
    } catch (ex) {
      setErr(String(ex));
    }
  }

  async function selectFrozen(key: string) {
    await selectFrozenKey(source as (typeof FROZEN)[number], key);
  }

  function statusClass(status: string) {
    const s = (status || "").toUpperCase();
    if (s === "ISSUED") return "act-badge green";
    if (s === "DRAFT") return "act-badge";
    if (s === "SUPERSEDED") return "act-badge purple";
    return "act-badge";
  }

  return (
    <section>
      <p className="eyebrow">Ürün · doküman</p>
      <h1>Ürün Arama</h1>
      <p className="lead">
        Ürün kodunu yazın; Word ve PDF dosyalarını açın veya indirin. Varsayılan kaynak revizyon
        yönetimi — resmi paketler buradadır.
      </p>
      <form className="search-bar" onSubmit={onSearch}>
        <select value={source} onChange={(e) => setSource(e.target.value as Source)}>
          <option value="workspace">Revizyon</option>
          {FROZEN.map((s) => (
            <option key={s} value={s}>
              {scopeLabel(s)}
            </option>
          ))}
        </select>
        <input
          value={q}
          onChange={(e) => setQ(e.target.value)}
          placeholder={source === "workspace" ? "Kod / açıklama / set…" : "Ürün / config kodu…"}
        />
        <button type="submit" disabled={busy}>
          {busy ? "…" : "Ara"}
        </button>
      </form>
      {err && <p className="error">{err}</p>}
      <LastDownloadBar />
      <p className="meta">
        {total} sonuç · kaynak {source === "workspace" ? "Revizyon" : scopeLabel(source)}
        {source === "workspace" && (
          <>
            {" "}
            · <Link to="/workspace">Revizyon</Link>
          </>
        )}
      </p>
      <div className="split">
        <ul className="key-list">
          {source === "workspace" &&
            wsHits.map((p) => (
              <li key={p.product_code}>
                <button
                  type="button"
                  className={wsDetail?.product.product_code === p.product_code ? "picked-rev" : ""}
                  onClick={() => void selectWs(p.product_code)}
                >
                  {p.product_code}
                  <span className="muted"> · {p.current_revision || "—"}</span>
                  <span className={statusClass(p.status)} style={{ marginLeft: "0.4rem" }}>
                    {statusLabel(p.status)}
                  </span>
                  {!p.complete && (
                    <span className="act-badge purple" style={{ marginLeft: "0.35rem" }}>
                      eksik
                    </span>
                  )}
                </button>
              </li>
            ))}
          {source !== "workspace" &&
            frozenKeys.map((k) => (
              <li key={k}>
                <button
                  type="button"
                  className={frozenDetail?.key === k ? "picked-rev" : ""}
                  onClick={() => void selectFrozen(k)}
                >
                  {k}
                </button>
              </li>
            ))}
          {total === 0 && (
            <li>
              <p className="muted" style={{ padding: "0.75rem" }}>
                Sonuç yok — sorguyu veya kaynağı değiştirin.
              </p>
            </li>
          )}
        </ul>
        <div className="detail">
          {!wsDetail && !frozenDetail && (
            <p className="muted">Listeden bir ürün seçin — Word ve PDF burada açılır.</p>
          )}
          {wsDetail && (
            <>
              <h2>
                {wsDetail.product.product_code}{" "}
                <span className={statusClass(wsDetail.product.status)}>
                  {statusLabel(wsDetail.product.status)}
                </span>
              </h2>
              <p className="meta">{wsDetail.product.description}</p>
              <p className="meta">
                {wsDetail.product.current_revision} · set {wsDetail.product.set_code}
              </p>
              <FilePairList
                files={wsDetail.current_files || []}
                onOpen={(f) =>
                  api
                    .wsOpen(wsDetail.product.product_code, f.name)
                    .then((r) => capture(r.download_url, downloadKindLabel(f.kind, f.name)))
                    .catch((e) => setErr(String(e)))
                }
              />
            </>
          )}
          {frozenDetail && (
            <>
              <h2>{frozenDetail.key}</h2>
              <FilePairList
                files={frozenDetail.files}
                onOpen={(f) =>
                  api
                    .openFile(source, frozenDetail.key, f.name)
                    .then((r) => capture(r.download_url, downloadKindLabel(f.kind, f.name)))
                    .catch((e) => setErr(String(e)))
                }
              />
            </>
          )}
        </div>
      </div>
    </section>
  );
}
