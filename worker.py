import re
import time
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional

import fitz
import pandas as pd
from PySide6 import QtCore

from heuristics import (
    normalize_spaces,
    h_find_fecha,
    h_find_tipo_examen,
    h_find_nombre_y_doc,
    h_find_cargo,
    h_find_concepto,
    h_find_examenes,
    h_find_reco_medicas,
    h_find_reco_ocup,
    h_find_habitos,
)


def extract_text_from_page(page: fitz.Page) -> str:
    try:
        return page.get_text("text") or ""
    except Exception:
        return ""


@dataclass
class FieldRule:
    name: str
    pattern: str  # expresión regular (opcional si se usa heurística de plantilla)

    def to_dict(self):
        return asdict(self)

    @staticmethod
    def from_dict(d):
        return FieldRule(name=d.get("name", ""), pattern=d.get("pattern", ""))


DEFAULT_TEMPLATE: List[FieldRule] = [
    FieldRule("FECHA DE REALIZACIÓN DEL EXÁMEN", r""),
    FieldRule("TIPO DE EXÁMEN MÉDICO OCUPACIONAL", r""),
    FieldRule("Apellidos y Nombres", r""),
    FieldRule("Documento de Identificación", r""),
    FieldRule("Cargo", r""),
    FieldRule("CONCEPTO DE APTITUD OCUPACIONAL", r""),
    FieldRule("Exámenes practicados (base del concepto)", r""),
    FieldRule("RECOMENDACIONES MÉDICAS", r""),
    FieldRule("RECOMENDACIONES OCUPACIONALES", r""),
    FieldRule("HÁBITOS Y ESTILO DE VIDA SALUDABLES", r""),
]


class ExtractWorker(QtCore.QThread):
    progress = QtCore.Signal(int, str)
    finished = QtCore.Signal(str)
    error = QtCore.Signal(str)

    def __init__(self, pdf_path: str, out_xlsx: str, fields: List[FieldRule], sheet_name: str,
                 use_template_heuristics: bool, max_pages: int, include_pdf_page: bool):
        super().__init__()
        self.pdf_path = pdf_path
        self.out_xlsx = out_xlsx
        self.fields = fields
        self.sheet_name = sheet_name or "Datos"
        self.use_template_heuristics = use_template_heuristics
        self.max_pages = max_pages
        self.include_pdf_page = include_pdf_page
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        t0 = time.time()
        try:
            doc = fitz.open(self.pdf_path)
        except Exception as e:
            self.error.emit(f"No se pudo abrir el PDF: {e}")
            return

        total_pages = min(len(doc), self.max_pages if self.max_pages > 0 else len(doc))
        data_rows: List[Dict[str, str]] = []

        compiled: Dict[str, Optional[re.Pattern]] = {}
        for fr in self.fields:
            try:
                compiled[fr.name] = re.compile(fr.pattern, re.I | re.S) if fr.pattern else None
            except re.error as rex:
                self.error.emit(f"Regex inválida para '{fr.name}': {rex}")
                return

        textless_pages = 0
        last_emit = 0

        for idx in range(total_pages):
            if self._cancel:
                self.error.emit("Proceso cancelado por el usuario.")
                return

            page = doc[idx]
            text = extract_text_from_page(page)
            if not text.strip():
                textless_pages += 1

            row: Dict[str, str] = {}

            for fr in self.fields:
                val = ""
                pat = compiled.get(fr.name)
                if pat is not None:
                    m = pat.search(text)
                    if m:
                        val = m.group(1) if m.groups() else m.group(0)
                        val = normalize_spaces(val)
                elif self.use_template_heuristics:
                    key = fr.name.strip().lower()
                    if key.startswith("fecha de realiz"):
                        val = h_find_fecha(text)
                    elif key.startswith("tipo de ex"):
                        val = h_find_tipo_examen(text)
                    elif key == "apellidos y nombres":
                        n, _ = h_find_nombre_y_doc(text)
                        val = n
                    elif key.startswith("documento"):
                        _, d = h_find_nombre_y_doc(text)
                        val = d
                    elif key == "cargo":
                        val = h_find_cargo(text)
                    elif key.startswith("concepto de aptitud"):
                        val = h_find_concepto(text)
                    elif key.startswith("exámenes practicados"):
                        val = h_find_examenes(text)
                    elif key.startswith("recomendaciones m"):
                        val = h_find_reco_medicas(text)
                    elif key.startswith("recomendaciones o"):
                        val = h_find_reco_ocup(text)
                    elif key.startswith("hábitos") or key.startswith("habitos"):
                        val = h_find_habitos(text)

                row[fr.name] = val

            if self.include_pdf_page:
                row["Página PDF"] = idx + 1

            data_rows.append(row)

            pct = int((idx + 1) * 100 / total_pages)
            if pct != last_emit:
                last_emit = pct
                self.progress.emit(pct, f"Procesando página {idx + 1}/{total_pages}")

        if textless_pages > max(3, int(total_pages * 0.6)):
            self.progress.emit(100, f"Advertencia: {textless_pages}/{total_pages} páginas sin texto (PDF escaneado)")

        try:
            df = pd.DataFrame(data_rows)
            with pd.ExcelWriter(self.out_xlsx, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name=self.sheet_name)
            dt = time.time() - t0
            self.finished.emit(self.out_xlsx)
            self.progress.emit(100, f"Completado en {dt:0.1f}s. Filas: {len(data_rows)}")
        except Exception as e:
            self.error.emit(f"Error al escribir Excel: {e}")