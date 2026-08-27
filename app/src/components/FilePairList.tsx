export type PairableFile = {
  stem: string;
  kind: string;
  name: string;
  exists: boolean;
  label?: string;
};

const STEM_LABEL: Record<string, string> = {
  "01_Technical_File": "Technical File",
  "02_EU_DoC": "EU DoC",
  "03_Label": "Packaging Label",
  "04_Shipment_Statement": "Shipment Statement",
};

export function isWordKind(kind: string, name?: string): boolean {
  const k = (kind || "").toUpperCase();
  if (k === "WORD" || k === "DOCX") return true;
  return Boolean(name && /\.docx$/i.test(name));
}

export function downloadKindLabel(kind: string, name?: string): string {
  return isWordKind(kind, name) ? "WORD indir" : "PDF indir";
}

function friendlyLabel(file: PairableFile): string {
  if (file.label) return file.label;
  return STEM_LABEL[file.stem] || file.stem.replace(/^\d+_/, "").replace(/_/g, " ");
}

type Pair = {
  stem: string;
  label: string;
  word?: PairableFile;
  pdf?: PairableFile;
};

function pairFiles(files: PairableFile[]): Pair[] {
  const order: string[] = [];
  const map = new Map<string, Pair>();
  for (const file of files) {
    const stem = file.stem || file.name.replace(/\.(docx|pdf)$/i, "");
    let row = map.get(stem);
    if (!row) {
      row = { stem, label: friendlyLabel(file) };
      map.set(stem, row);
      order.push(stem);
    } else if (file.label) {
      row.label = file.label;
    }
    if (isWordKind(file.kind, file.name)) row.word = file;
    else row.pdf = file;
  }
  return order.map((stem) => map.get(stem)!);
}

export default function FilePairList({
  files,
  onOpen,
}: {
  files: PairableFile[];
  onOpen: (file: PairableFile) => void;
}) {
  const rows = pairFiles(files);
  if (rows.length === 0) return null;
  return (
    <ul className="file-list">
      {rows.map((row) => (
        <li key={row.stem}>
          <span className="file-doc-name">{row.label}</span>
          <span className="file-pair-actions">
            <button
              type="button"
              disabled={!row.word?.exists}
              onClick={() => row.word && onOpen(row.word)}
            >
              WORD
            </button>
            <button
              type="button"
              disabled={!row.pdf?.exists}
              onClick={() => row.pdf && onOpen(row.pdf)}
            >
              PDF
            </button>
          </span>
        </li>
      ))}
    </ul>
  );
}
