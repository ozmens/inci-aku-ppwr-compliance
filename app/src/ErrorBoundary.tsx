import { Component, ReactNode } from "react";

type Props = { children: ReactNode };
type State = { error: string | null };

export default class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null };

  static getDerivedStateFromError(err: unknown) {
    return { error: err instanceof Error ? err.message : String(err) };
  }

  render() {
    if (!this.state.error) return this.props.children;
    return (
      <div className="login-screen">
        <div className="login-card">
          <p className="eyebrow" lang="en">PPWR Compliance Suite</p>
          <h1>Sayfa Yüklenemedi</h1>
          <p className="lead">Sayfayı yenileyin. Sorun sürerse uygulamayı yeniden başlatın.</p>
          <p className="error">{this.state.error}</p>
          <button type="button" onClick={() => window.location.assign("/login")}>
            Girişe dön
          </button>
        </div>
      </div>
    );
  }
}
