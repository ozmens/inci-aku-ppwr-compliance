"""Document path resolution — single authority for relative DOCX links."""

from __future__ import annotations

from dataclasses import dataclass, field

from models.document import Document, DocumentLink

FAMILY_FOLDER = {
    "STARTER": "01_STARTER",
    "INDUSTRIAL": "02_INDUSTRIAL",
    "CONTAINER": "03_CONTAINER",
}

DOC_KIND_FILE = {
    "TF": "01_Technical_File.docx",
    "TECHNICAL_FILE": "01_Technical_File.docx",
    "DOC": "02_EU_DoC.docx",
    "EU_DOC": "02_EU_DoC.docx",
    "LABEL": "03_Label.docx",
    "STM": "04_Shipment_Statement.docx",
    "STATEMENT": "04_Shipment_Statement.docx",
}


def normalize_relative_path(path: str | None) -> str:
    if not path:
        return ""
    t = str(path).replace("\\", "/").strip().lstrip("./")
    return t


def family_folder(family: str | None) -> str:
    key = (family or "").strip().upper()
    if key in FAMILY_FOLDER:
        return FAMILY_FOLDER[key]
    if key.startswith("ST") or "STARTER" in key:
        return FAMILY_FOLDER["STARTER"]
    if key.startswith("IND") or "INDUSTRIAL" in key:
        return FAMILY_FOLDER["INDUSTRIAL"]
    if key.startswith("CNT") or "CONTAINER" in key:
        return FAMILY_FOLDER["CONTAINER"]
    # packaging set code prefixes
    sc = key
    if sc.startswith("ST-"):
        return FAMILY_FOLDER["STARTER"]
    if sc.startswith("IND-"):
        return FAMILY_FOLDER["INDUSTRIAL"]
    if sc.startswith("CNT-"):
        return FAMILY_FOLDER["CONTAINER"]
    return FAMILY_FOLDER["STARTER"]


@dataclass(slots=True)
class DocumentLinkService:
    """Path resolution + optional in-memory document registry."""

    documents: list[Document] = field(default_factory=list)
    links: list[DocumentLink] = field(default_factory=list)
    _uri_by_code: dict[str, str] = field(default_factory=dict)

    def add_document(self, document: Document) -> None:
        self.documents.append(document)

    def link(self, document_link: DocumentLink) -> None:
        self.links.append(document_link)

    def register_uri(self, document_code: str, file_uri: str) -> None:
        code = (document_code or "").strip()
        uri = normalize_relative_path(file_uri)
        if code and uri:
            self._uri_by_code[code] = uri

    def load_from_document_library_rows(
        self, rows: list[tuple[str, str]]
    ) -> None:
        """rows: (document_code, file_uri)."""
        for code, uri in rows:
            self.register_uri(str(code or ""), str(uri or ""))

    def relative_path(
        self,
        *,
        packaging_set_code: str,
        doc_kind: str,
        family: str | None = None,
        document_code: str | None = None,
    ) -> str:
        """Resolve relative DOCX path. Prefer DOCUMENT_LIBRARY URI when known."""
        if document_code:
            known = self._uri_by_code.get(str(document_code).strip())
            if known:
                return known
        set_code = (packaging_set_code or "").strip()
        if not set_code:
            return ""
        folder = family_folder(family or set_code)
        fname = DOC_KIND_FILE.get((doc_kind or "").strip().upper())
        if not fname:
            return ""
        return f"{folder}/{set_code}/{fname}"

    def pack_paths(
        self,
        *,
        packaging_set_code: str,
        family: str | None = None,
        tf_id: str | None = None,
        doc_id: str | None = None,
        label_id: str | None = None,
        statement_id: str | None = None,
    ) -> dict[str, str]:
        return {
            "tf": self.relative_path(
                packaging_set_code=packaging_set_code,
                doc_kind="TF",
                family=family,
                document_code=tf_id,
            ),
            "doc": self.relative_path(
                packaging_set_code=packaging_set_code,
                doc_kind="DOC",
                family=family,
                document_code=doc_id,
            ),
            "label": self.relative_path(
                packaging_set_code=packaging_set_code,
                doc_kind="LABEL",
                family=family,
                document_code=label_id,
            ),
            "statement": self.relative_path(
                packaging_set_code=packaging_set_code,
                doc_kind="STM",
                family=family,
                document_code=statement_id,
            ),
        }

    def links_for_configuration(self, packaging_configuration_id: str) -> list[DocumentLink]:
        pid = packaging_configuration_id.strip().upper()
        return [
            ln
            for ln in self.links
            if (ln.packaging_configuration_id or "").upper() == pid
        ]

    def broken_links(self) -> list[DocumentLink]:
        known = {d.document_id for d in self.documents}
        return [ln for ln in self.links if ln.document_id not in known]
