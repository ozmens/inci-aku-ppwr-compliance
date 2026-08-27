import { useEffect, useMemo, useState } from "react";
import { Link, useNavigate } from "react-router-dom";
import { api, ScopeInfo } from "../api";
import {
  actionBadge,
  actionLabel,
  eventCode,
  eventDetail,
  scopeLabel,
  type ActivityEvent,
} from "../labels";

function fmtNum(n: number | null | undefined): string {
  if (n == null || Number.isNaN(n)) return "—";
  return new Intl.NumberFormat("tr-TR").format(n);
}

const SCOPE_META: Record<string, { title: string; blurb: string; tone: string }> = {
  starter: {
    title: "STARTER",
    blurb: "Bireysel ambalaj · onaylı doküman seti",
    tone: "",
  },
  industrial: {
    title: "INDUSTRIAL",
    blurb: "Endüstriyel paketleme · seri üretim",
    tone: "purple",
  },
  container: {
    title: "CONTAINER",
    blurb: "Konteyner yükleme · sevkiyat konfigi",
    tone: "green",
  },
  component: {
    title: "COMPONENT",
    blurb: "Bileşen ve yedek parça varyantları",
    tone: "amber",
  },
};

export default function Home() {
  const navigate = useNavigate();
  const [starterSets, setStarterSets] = useState<number | null>(null);
  const [scopes, setScopes] = useState<ScopeInfo[]>([]);
  const [events, setEvents] = useState<ActivityEvent[]>([]);
  const [activityOpen, setActivityOpen] = useState(false);

  useEffect(() => {
    api
      .mastersSummary()
      .then((s) => {
        setStarterSets(s.starter?.unique_sets ?? null);
      })
      .catch(() => undefined);

    api
      .scopes()
      .then((r) => setScopes(r.scopes))
      .catch(() => undefined);

    api
      .wsActivity(12)
      .then((a) => setEvents(a.events))
      .catch(() => setEvents([]));
  }, []);

  const orderedScopes = useMemo(() => {
    const order = ["starter", "industrial", "container", "component"];
    return order
      .map((k) => scopes.find((s) => s.key === k))
      .filter(Boolean) as ScopeInfo[];
  }, [scopes]);

  return (
    <section>
      <div className="hero-strip" aria-hidden>
        <img className="hero-shot" src="/hero-battery.png" alt="" />
        <img className="hero-shot" src="/hero-industrial-battery.png" alt="" />
        <img className="hero-shot" src="/hero-pack-still.png" alt="" />
      </div>

      <div className="section-head">
        <h2>Kapsamlar</h2>
        <Link to="/scopes">Doküman Merkezi →</Link>
      </div>
      <div className="scope-enter-grid">
        {orderedScopes.map((s) => {
          const m = SCOPE_META[s.key] || {
            title: s.key.toUpperCase(),
            blurb: "",
            tone: "",
          };
          return (
            <button
              key={s.key}
              type="button"
              className={`scope-enter-card ${m.tone}`}
              onClick={() => navigate(`/scopes/${s.key}`)}
            >
              <span className="scope-enter-label">{m.title}</span>
              <strong>{fmtNum(s.records)}</strong>
              <span className="muted">{m.blurb}</span>
              <span className="scope-enter-go">Aç ›</span>
            </button>
          );
        })}
      </div>

      <div className="kpi-row kpi-row-2">
        <button type="button" className="kpi kpi-btn accent-purple" onClick={() => navigate("/bom")}>
          <div className="kpi-label">Ambalaj setleri</div>
          <strong>{fmtNum(starterSets)}</strong>
          <span>Starter fiziksel setler</span>
        </button>
      </div>

      <div className="section-head">
        <h2>Doküman Ailesi</h2>
      </div>
      <div className="doc-strip">
        <Link className="doc-chip" to="/scopes/starter?doc=01">
          <strong>Technical File</strong>
          <span>01 · Teknik dosya</span>
        </Link>
        <Link className="doc-chip" to="/scopes/starter?doc=02">
          <strong>EU DoC</strong>
          <span>02 · Uygunluk beyanı</span>
        </Link>
        <Link className="doc-chip" to="/scopes/starter?doc=03">
          <strong>Packaging Label</strong>
          <span>03 · Ambalaj etiketi</span>
        </Link>
        <Link className="doc-chip" to="/scopes/starter?doc=04">
          <strong>Shipment Statement</strong>
          <span>04 · Sevkiyat beyanı</span>
        </Link>
      </div>

      <div className="panel">
        <div className="section-head">
          <button
            type="button"
            className="section-fold"
            aria-expanded={activityOpen}
            onClick={() => setActivityOpen((v) => !v)}
          >
            <h2>Son İşlemler</h2>
            <span className="fold-arrow" aria-hidden>
              {activityOpen ? "▾" : "▸"}
            </span>
          </button>
          <Link to="/workspace">Revizyon →</Link>
        </div>
        {activityOpen &&
          (events.length === 0 ? (
          <p className="muted">Henüz işlem kaydı yok. Ürün arama veya müşteri paketi ile başlayın.</p>
        ) : (
          <table className="activity-table">
            <thead>
              <tr>
                <th>Kod</th>
                <th>Açıklama</th>
                <th>İşlem</th>
                <th>Kapsam</th>
                <th>Tarih / Saat</th>
              </tr>
            </thead>
            <tbody>
              {events.map((ev, i) => {
                const code = eventCode(ev);
                const detail = eventDetail(ev);
                return (
                  <tr key={`${ev.at}-${i}`}>
                    <td>
                      {ev.product_code ? (
                        <Link to={`/search?q=${encodeURIComponent(String(ev.product_code))}&source=workspace`}>
                          <code>{code}</code>
                        </Link>
                      ) : ev.supplier_id ? (
                        <Link to="/suppliers">
                          <code>{code}</code>
                        </Link>
                      ) : ev.customer_id ? (
                        <Link to="/customers">
                          <code>{code}</code>
                        </Link>
                      ) : (
                        <code>{code}</code>
                      )}
                    </td>
                    <td>{detail}</td>
                    <td>
                      <span className={`act-badge ${actionBadge(String(ev.action))}`}>
                        {actionLabel(String(ev.action))}
                      </span>
                    </td>
                    <td className="muted">{scopeLabel(String(ev.scope || "workspace"))}</td>
                    <td className="muted">
                      {String(ev.at || "").slice(0, 19).replace("T", " ")}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        ))}
      </div>
    </section>
  );
}
