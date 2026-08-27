import { FormEvent, useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { api } from "../api";
import ThemeToggle from "../components/ThemeToggle";

export default function Login() {
  const navigate = useNavigate();
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    const token = localStorage.getItem("inci_ppwr_token");
    if (!token) return;
    api
      .authMe()
      .then(() => navigate("/", { replace: true }))
      .catch(() => {
        localStorage.removeItem("inci_ppwr_token");
        localStorage.removeItem("inci_ppwr_user");
      });
  }, [navigate]);

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setErr("");
    setBusy(true);
    try {
      await api.authLogin(username.trim(), password);
      navigate("/", { replace: true });
    } catch (ex) {
      const msg = String(ex);
      if (msg.includes("404") || msg.includes("Failed to fetch") || msg.includes("NetworkError")) {
        setErr("Sunucuya bağlanılamıyor. Biraz sonra tekrar deneyin.");
      } else {
        setErr("Giriş başarısız. Kullanıcı adı veya şifreyi kontrol edin.");
      }
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="login-screen">
      <ThemeToggle variant="bar" />
      <form className="login-card" onSubmit={onSubmit}>
        <img className="login-logo" src="/inci-aku-logo.png" alt="İnci Akü" />
        <p className="eyebrow" lang="en">
          PPWR Compliance Suite
        </p>
        <h1>Giriş</h1>
        <p className="lead">Yetkili hesabınızla giriş yapın.</p>
        <label className="block-label">
          Kullanıcı adı
          <input
            value={username}
            onChange={(e) => setUsername(e.target.value)}
            autoComplete="username"
            required
            autoFocus
          />
        </label>
        <label className="block-label">
          Şifre
          <input
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            autoComplete="current-password"
            required
          />
        </label>
        {err && <p className="error">{err}</p>}
        <button type="submit" disabled={busy || !username.trim() || !password}>
          {busy ? "…" : "Oturum aç"}
        </button>
      </form>
    </div>
  );
}
