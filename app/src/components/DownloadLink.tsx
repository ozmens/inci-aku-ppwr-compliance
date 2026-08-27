import { withAccessToken } from "../download";

/** Visible download fallback when popup blockers swallow auto-open. */
export default function DownloadLink({
  href,
  label = "İndir",
  filename,
}: {
  href: string;
  label?: string;
  filename?: string;
}) {
  if (!href) return null;
  const safe = withAccessToken(href);
  return (
    <a className="download-cta" href={safe} download={filename || undefined} rel="noopener noreferrer">
      {label}
    </a>
  );
}
