/** Append session token so downloads / <img> work without relying on cookies alone. */
export function withAccessToken(url: string): string {
  if (!url) return url;
  const token = localStorage.getItem("inci_ppwr_token");
  if (!token) return url;
  const sep = url.includes("?") ? "&" : "?";
  return `${url}${sep}access_token=${encodeURIComponent(token)}`;
}

/** Trigger browser download; returns the final URL for an explicit “İndir” link. */
export function triggerDownload(url: string, filename?: string): string {
  const href = withAccessToken(url);
  const a = document.createElement("a");
  a.href = href;
  a.rel = "noopener noreferrer";
  if (filename) a.download = filename;
  else a.target = "_blank";
  document.body.appendChild(a);
  a.click();
  a.remove();
  return href;
}
