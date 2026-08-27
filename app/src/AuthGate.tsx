import { ReactNode, useEffect, useState } from "react";
import { Navigate, useLocation } from "react-router-dom";
import { api, AuthUser } from "./api";

export default function AuthGate({ children }: { children: ReactNode }) {
  const location = useLocation();
  const [ready, setReady] = useState(false);
  const [required, setRequired] = useState(true);
  const [user, setUser] = useState<AuthUser | null>(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const st = await api.authStatus();
        if (cancelled) return;
        if (!st.auth_required) {
          setRequired(false);
          setReady(true);
          return;
        }
        const token = localStorage.getItem("inci_ppwr_token");
        if (!token) {
          setUser(null);
          setReady(true);
          return;
        }
        const me = await api.authMe();
        if (cancelled) return;
        setUser(me.user);
      } catch {
        if (!cancelled) {
          localStorage.removeItem("inci_ppwr_token");
          setUser(null);
        }
      } finally {
        if (!cancelled) setReady(true);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [location.pathname]);

  if (!ready) {
    return (
      <div className="login-screen">
        <div className="login-card">
          <p className="eyebrow" lang="en">PPWR Compliance Suite</p>
          <h1>Yükleniyor</h1>
          <p className="lead">Oturum kontrol ediliyor…</p>
        </div>
      </div>
    );
  }

  if (required && !user && location.pathname !== "/login") {
    return <Navigate to="/login" replace state={{ from: location.pathname }} />;
  }

  return <>{children}</>;
}
