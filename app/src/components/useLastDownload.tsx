import { useState } from "react";
import DownloadLink from "./DownloadLink";

/** Inline “last file ready” strip after WORD / PDF. */
export function useLastDownload() {
  const [last, setLast] = useState<{ href: string; label: string } | null>(null);

  function capture(href?: string | null, label = "Dosyayı indir") {
    if (href) setLast({ href, label });
  }

  function LastDownloadBar() {
    if (!last) return null;
    return (
      <p className="last-download">
        <span className="muted">Dosya hazır — </span>
        <DownloadLink href={last.href} label={last.label} />
      </p>
    );
  }

  return { capture, LastDownloadBar, clear: () => setLast(null) };
}
