import { FormEvent, useEffect, useState } from "react";
import { Navigate, NavLink, Route, Routes, useLocation, useNavigate } from "react-router-dom";
import AuthGate from "./AuthGate";
import Bom from "./pages/Bom";
import Components from "./pages/Components";
import Customers from "./pages/Customers";
import DesktopDrop from "./pages/DesktopDrop";
import Home from "./pages/Home";
import Login from "./pages/Login";
import Master from "./pages/Master";
import PackBuilder from "./pages/PackBuilder";
import Pipeline from "./pages/Pipeline";
import Scopes from "./pages/Scopes";
import Search from "./pages/Search";
import Suppliers from "./pages/Suppliers";
import Users from "./pages/Users";
import Workspace from "./pages/Workspace";
import { api, AuthUser } from "./api";
import AccountPassword from "./components/AccountPassword";
import ThemeToggle from "./components/ThemeToggle";

const links: { to: string; label: string; ico: string; end?: boolean }[] = [
  { to: "/", label: "Ana Sayfa", ico: "⌂", end: true },
  { to: "/search", label: "Ürün Arama", ico: "⌕" },
  { to: "/workspace", label: "Revizyon", ico: "▣" },
  { to: "/customers", label: "Müşteri Paketi", ico: "▦" },
  { to: "/suppliers", label: "Tedarikçi", ico: "◎" },
  { to: "/components", label: "Bileşen Matrisi", ico: "⬡" },
  { to: "/scopes", label: "Doküman Merkezi", ico: "▤" },
  { to: "/master", label: "Master Veri", ico: "☰" },
  { to: "/bom", label: "Ambalaj BOM", ico: "≡" },
  { to: "/drop", label: "Paket ZIP", ico: "↓" },
  { to: "/dil-foto", label: "Dil & Foto", ico: "⇢" },
  { to: "/packs", label: "Aday Paket", ico: "+" },
  { to: "/users", label: "Kullanıcılar", ico: "◇" },
];

export default function App() {
  const navigate = useNavigate();
  const location = useLocation();
  const [q, setQ] = useState("");
  const [apiOk, setApiOk] = useState(false);
  const [version, setVersion] = useState("");
  const [starterN, setStarterN] = useState<number | null>(null);
  const [industrialN, setIndustrialN] = useState<number | null>(null);
  const [user, setUser] = useState<AuthUser | null>(() => {
    try {
      const raw = localStorage.getItem("inci_ppwr_user");
      return raw ? (JSON.parse(raw) as AuthUser) : null;
    } catch {
      return null;
    }
  });
  const [navCollapsed, setNavCollapsed] = useState(() => {
    try {
      return localStorage.getItem("inci_ppwr_nav_collapsed") === "1";
    } catch {
      return false;
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem("inci_ppwr_nav_collapsed", navCollapsed ? "1" : "0");
    } catch {
      /* ignore */
    }
  }, [navCollapsed]);

  useEffect(() => {
    function onKey(e: KeyboardEvent) {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "b") {
        e.preventDefault();
        setNavCollapsed((v) => !v);
      }
    }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  useEffect(() => {
    if (location.pathname === "/login") return;
    api
      .health()
      .then((h) => {
        setApiOk(h.ok);
        setVersion(h.version || "");
      })
      .catch(() => setApiOk(false));
    api
      .scopes()
      .then((r) => {
        const starter = r.scopes.find((s) => s.key === "starter");
        const industrial = r.scopes.find((s) => s.key === "industrial");
        setStarterN(starter?.records ?? null);
        setIndustrialN(industrial?.records ?? null);
      })
      .catch(() => undefined);
    api
      .authMe()
      .then((m) => setUser(m.user))
      .catch(() => setUser(null));
  }, [location.pathname]);

  function onTopSearch(e: FormEvent) {
    e.preventDefault();
    const query = q.trim();
    navigate(query ? `/search?q=${encodeURIComponent(query)}` : "/search");
  }

  async function onLogout() {
    await api.authLogout();
    setUser(null);
    navigate("/login", { replace: true });
  }

  if (location.pathname === "/login") {
    return <Login />;
  }

  return (
    <AuthGate>
      <div className={`app-shell ${navCollapsed ? "nav-collapsed" : ""}`}>
        <aside className="sidebar">
          <div className="sidebar-brand">
            <img className="sidebar-brand-logo" src="/inci-aku-logo.png" alt="İnci Akü" />
            <span className="sidebar-brand-sub">PPWR Compliance Suite</span>
          </div>
          <nav className="sidebar-nav">
            {links
              .filter((l) => l.to !== "/users" || user?.role === "admin")
              .map((l) => (
                <NavLink
                  key={l.to}
                  to={l.to}
                  end={l.end}
                  title={l.label}
                  className={({ isActive }) => (isActive ? "nav active" : "nav")}
                >
                  <span className="nav-ico" aria-hidden>
                    {l.ico}
                  </span>
                  <span className="nav-label">{l.label}</span>
                </NavLink>
              ))}
          </nav>
          <div className="sidebar-foot">
            <div className="sidebar-foot-status">
              <span className={`dot ${apiOk ? "" : "off"}`} />
              <span className="nav-label">{apiOk ? "Hazır" : "Bağlantı yok"}</span>
            </div>
            <div className="sidebar-foot-meta nav-label">
              {user ? user.username : "—"} · v{version || "—"}
            </div>
            <AccountPassword />
            <button type="button" className="btn-ghost sidebar-logout nav-label" onClick={() => void onLogout()}>
              Çıkış
            </button>
          </div>
        </aside>

        <div className="app-main-col">
          <header className="topbar">
            <div className="topbar-left">
              <button
                type="button"
                className="nav-toggle"
                title={navCollapsed ? "Menüyü aç (Ctrl+B)" : "Menüyü daralt (Ctrl+B)"}
                onClick={() => setNavCollapsed((v) => !v)}
              >
                {navCollapsed ? "☰" : "⟨"}
              </button>
              <div className="topbar-title">
                <img className="topbar-logo" src="/inci-aku-logo.png" alt="İnci Akü" />
                <span className="topbar-suite">PPWR Compliance Suite</span>
              </div>
            </div>
            <form className="top-search" onSubmit={onTopSearch}>
              <input
                value={q}
                onChange={(e) => setQ(e.target.value)}
                placeholder="Ürün kodu veya doküman ara…"
                aria-label="Ürün kodu ara"
              />
              <button type="submit">Ara</button>
              <ThemeToggle variant="bar" />
            </form>
            <div className="topbar-actions">
              <button
                type="button"
                className="scope-pill clickable"
                onClick={() => navigate("/scopes/starter")}
              >
                STARTER <em>{starterN ?? "—"}</em>
              </button>
              <button
                type="button"
                className="scope-pill clickable"
                onClick={() => navigate("/scopes/industrial")}
              >
                INDUSTRIAL <em>{industrialN ?? "—"}</em>
              </button>
              {user?.role === "admin" && (
                <button
                  type="button"
                  className="scope-pill clickable"
                  title="Kullanıcı yönetimi"
                  onClick={() => navigate("/users")}
                >
                  {user.username}
                </button>
              )}
              {user?.role !== "admin" && user && (
                <span className="scope-pill" title={user.display_name || user.username}>
                  {user.username}
                </span>
              )}
              <button type="button" className="scope-pill clickable" onClick={() => void onLogout()}>
                Çıkış
              </button>
            </div>
          </header>

          <main className="main">
            <Routes>
              <Route path="/" element={<Home />} />
              <Route path="/workspace" element={<Workspace />} />
              <Route path="/customers" element={<Customers />} />
              <Route path="/suppliers" element={<Suppliers />} />
              <Route path="/components" element={<Components />} />
              <Route path="/drop" element={<DesktopDrop />} />
              <Route path="/scopes" element={<Scopes />} />
              <Route path="/scopes/:scopeKey" element={<Scopes />} />
              <Route path="/search" element={<Search />} />
              <Route path="/master" element={<Master />} />
              <Route path="/bom" element={<Bom />} />
              <Route path="/bom/:setCode" element={<Bom />} />
              <Route path="/gaps" element={<Navigate to="/" replace />} />
              <Route path="/dil-foto" element={<Pipeline />} />
              <Route path="/pipeline" element={<Navigate to="/dil-foto" replace />} />
              <Route path="/packs" element={<PackBuilder />} />
              <Route
                path="/users"
                element={user?.role === "admin" ? <Users /> : <Navigate to="/" replace />}
              />
            </Routes>
          </main>
        </div>
      </div>
    </AuthGate>
  );
}
