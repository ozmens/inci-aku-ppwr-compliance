import { FormEvent, useEffect, useState } from "react";
import { api, AuthUser } from "../api";

type UserRow = AuthUser & { active?: boolean };

export default function Users() {
  const [users, setUsers] = useState<UserRow[]>([]);
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [displayName, setDisplayName] = useState("");
  const [role, setRole] = useState("user");
  const [err, setErr] = useState("");
  const [okMsg, setOkMsg] = useState("");
  const [busy, setBusy] = useState(false);
  const [resetId, setResetId] = useState<string | null>(null);
  const [resetPw, setResetPw] = useState("");
  const [resetPw2, setResetPw2] = useState("");

  async function refresh() {
    const r = await api.authUsers();
    setUsers(r.users as UserRow[]);
  }

  useEffect(() => {
    refresh().catch((e) => setErr(String(e)));
  }, []);

  async function onCreate(e: FormEvent) {
    e.preventDefault();
    setErr("");
    setOkMsg("");
    setBusy(true);
    try {
      await api.authCreateUser({
        username: username.trim(),
        password,
        display_name: displayName.trim() || username.trim(),
        role,
      });
      setUsername("");
      setPassword("");
      setDisplayName("");
      setRole("user");
      setOkMsg("Kullanıcı oluşturuldu");
      await refresh();
    } catch (ex) {
      setErr(String(ex));
    } finally {
      setBusy(false);
    }
  }

  async function onResetSubmit(e: FormEvent) {
    e.preventDefault();
    if (!resetId) return;
    if (resetPw.length < 6) {
      setErr("Şifre en az 6 karakter olmalı");
      return;
    }
    if (resetPw !== resetPw2) {
      setErr("Şifreler eşleşmiyor");
      return;
    }
    setErr("");
    setOkMsg("");
    try {
      await api.authResetPassword(resetId, resetPw);
      const u = users.find((x) => x.id === resetId);
      setOkMsg(`Şifre güncellendi: ${u?.username || resetId}`);
      setResetId(null);
      setResetPw("");
      setResetPw2("");
    } catch (ex) {
      setErr(String(ex));
    }
  }

  async function onToggle(u: UserRow) {
    setErr("");
    try {
      await api.authSetActive(u.id, !(u.active !== false));
      await refresh();
    } catch (ex) {
      setErr(String(ex));
    }
  }

  return (
    <section>
      <p className="eyebrow">Kullanıcılar</p>
      <h1>Kullanıcı Yönetimi</h1>
      <p className="lead">
        Yeni kullanıcı ekleyin, şifre sıfırlayın veya hesabı pasifleştirin.
      </p>

      {err && <p className="error">{err}</p>}
      {okMsg && <p className="ok">{okMsg}</p>}

      <div className="split">
        <div className="detail">
          <h2 className="section-title">Yeni Kullanıcı</h2>
          <form onSubmit={onCreate}>
            <label className="block-label">
              Kullanıcı adı
              <input value={username} onChange={(e) => setUsername(e.target.value)} required />
            </label>
            <label className="block-label">
              Görünen ad
              <input value={displayName} onChange={(e) => setDisplayName(e.target.value)} />
            </label>
            <label className="block-label">
              Şifre
              <input
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                required
                minLength={6}
              />
            </label>
            <label className="block-label">
              Rol
              <select value={role} onChange={(e) => setRole(e.target.value)}>
                <option value="user">Kullanıcı</option>
                <option value="admin">Yönetici</option>
              </select>
            </label>
            <button type="submit" disabled={busy}>
              {busy ? "…" : "Oluştur"}
            </button>
          </form>

          {resetId && (
            <form onSubmit={onResetSubmit} style={{ marginTop: "1.25rem" }}>
              <h2 className="section-title">Şifre Sıfırla</h2>
              <p className="meta">{users.find((u) => u.id === resetId)?.username}</p>
              <label className="block-label">
                Yeni şifre
                <input
                  type="password"
                  value={resetPw}
                  onChange={(e) => setResetPw(e.target.value)}
                  required
                  minLength={6}
                  autoFocus
                />
              </label>
              <label className="block-label">
                Tekrar
                <input
                  type="password"
                  value={resetPw2}
                  onChange={(e) => setResetPw2(e.target.value)}
                  required
                  minLength={6}
                />
              </label>
              <div className="engine-actions">
                <button type="submit">Kaydet</button>
                <button
                  type="button"
                  className="btn-ghost"
                  onClick={() => {
                    setResetId(null);
                    setResetPw("");
                    setResetPw2("");
                  }}
                >
                  İptal
                </button>
              </div>
            </form>
          )}
        </div>

        <div>
          <h2 className="section-title">Kullanıcılar ({users.length})</h2>
          <table className="data-table">
            <thead>
              <tr>
                <th>Kullanıcı</th>
                <th>Rol</th>
                <th>Durum</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {users.map((u) => (
                <tr key={u.id}>
                  <td>
                    <strong>{u.username}</strong>
                    <div className="muted">{u.display_name || ""}</div>
                  </td>
                  <td>
                    <span className={u.role === "admin" ? "act-badge green" : "act-badge"}>
                      {u.role === "admin" ? "yönetici" : "kullanıcı"}
                    </span>
                  </td>
                  <td>{u.active === false ? "kapalı" : "aktif"}</td>
                  <td>
                    <button
                      type="button"
                      className="btn-ghost"
                      onClick={() => {
                        setResetId(u.id);
                        setResetPw("");
                        setResetPw2("");
                        setErr("");
                      }}
                    >
                      Şifre
                    </button>{" "}
                    {u.username !== "admin" && (
                      <button type="button" className="btn-ghost" onClick={() => void onToggle(u)}>
                        {u.active === false ? "Aç" : "Kapat"}
                      </button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </section>
  );
}
