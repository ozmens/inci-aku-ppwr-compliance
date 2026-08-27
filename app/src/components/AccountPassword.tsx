import { FormEvent, useState } from "react";
import { api } from "../api";

/** Compact self-service password change for any signed-in user. */
export default function AccountPassword() {
  const [open, setOpen] = useState(false);
  const [current, setCurrent] = useState("");
  const [next, setNext] = useState("");
  const [again, setAgain] = useState("");
  const [err, setErr] = useState("");
  const [ok, setOk] = useState("");
  const [busy, setBusy] = useState(false);

  if (!open) {
    return (
      <button type="button" className="btn-ghost sidebar-logout nav-label" onClick={() => setOpen(true)}>
        Şifre Değiştir
      </button>
    );
  }

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setErr("");
    setOk("");
    if (next.length < 6) {
      setErr("Yeni şifre en az 6 karakter");
      return;
    }
    if (next !== again) {
      setErr("Şifreler eşleşmiyor");
      return;
    }
    setBusy(true);
    try {
      await api.authChangePassword(current, next);
      setOk("Şifre güncellendi");
      setCurrent("");
      setNext("");
      setAgain("");
      setTimeout(() => setOpen(false), 900);
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  return (
    <form className="account-pw" onSubmit={onSubmit}>
      <p className="nav-label" style={{ fontWeight: 600, marginBottom: "0.35rem" }}>
        Şifre Değiştir
      </p>
      <input
        type="password"
        placeholder="Mevcut şifre"
        value={current}
        onChange={(e) => setCurrent(e.target.value)}
        required
        autoComplete="current-password"
      />
      <input
        type="password"
        placeholder="Yeni şifre"
        value={next}
        onChange={(e) => setNext(e.target.value)}
        required
        minLength={6}
        autoComplete="new-password"
      />
      <input
        type="password"
        placeholder="Yeni şifre tekrar"
        value={again}
        onChange={(e) => setAgain(e.target.value)}
        required
        minLength={6}
        autoComplete="new-password"
      />
      {err && <p className="error" style={{ fontSize: "0.75rem" }}>{err}</p>}
      {ok && <p className="ok" style={{ fontSize: "0.75rem" }}>{ok}</p>}
      <div className="engine-actions">
        <button type="submit" disabled={busy}>
          {busy ? "…" : "Kaydet"}
        </button>
        <button type="button" className="btn-ghost" onClick={() => setOpen(false)}>
          Kapat
        </button>
      </div>
    </form>
  );
}
