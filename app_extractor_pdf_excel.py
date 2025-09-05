# -*- coding: utf-8 -*-
"""
Extractor PDF → Excel (GUI de escritorio, moderno)
=================================================

✔ Interfaz original (tabla de campos, presets, opciones, progreso y cancelar)
✔ Extrae SOLO los nombres desde el bloque **DATOS DEL TRABAJADOR / ASPIRANTE**
✔ No confunde CARGO con el nombre
✔ Documento (CC/TI/CE/PT) tomado cerca del nombre dentro del mismo bloque
✔ Hasta 2000 páginas por corrida

Requisitos rápidos:
    pip install "PySide6>=6.9.2,<7" "PyMuPDF>=1.24.10,<1.26" pandas XlsxWriter

Ejecución:
    python app_extractor_pdf_excel.py
"""

import json
import os
import re
import sys
import time
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional, Tuple

import fitz  # PyMuPDF
import pandas as pd
from PySide6 import QtCore, QtGui, QtWidgets

APP_TITLE = "Extractor PDF → Excel"
VERSION = "1.1.0"  # Reinstala interfaz original + fix nombres/documentos

# -----------------------------
# Utilidades de extracción
# -----------------------------

def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()


def extract_text_from_page(page: fitz.Page) -> str:
    try:
        return page.get_text("text") or ""
    except Exception:
        return ""


# -----------------------------
# Heurísticas específicas (plantilla ocupacional)
# -----------------------------

def h_find_fecha(text: str) -> str:
    # Prioriza patrón DÍA MES AÑO; si no, primer dd mm yyyy que no sea pie de página
    m = re.search(r"D[ÍI]A\s+MES\s+A[ÑN]O\s+(\d{2})\s+(\d{2})\s+(\d{4})", text)
    if m:
        d, mth, y = m.groups()
        return f"{y}-{mth}-{d}"
    for m3 in re.finditer(r"(\d{2})\s+(\d{2})\s+(\d{4})", text):
        d, mth, y = m3.groups()
        pre = text[max(0, m3.start()-60):m3.start()]
        if "Impreso el" not in pre:
            return f"{y}-{mth}-{d}"
    return ""


def h_find_tipo_examen(text: str) -> str:
    m = re.search(r"(EVALUACI[ÓO]N?\s*M[ÉE]DICO\s*OCUPACIONAL.*)$", text, re.I | re.M)
    return normalize_spaces(m.group(1)) if m else ""


def h_find_nombre_y_doc(text: str) -> Tuple[str, str]:
    """
    Extrae (Apellidos y Nombres, Documento) **exclusivamente** del bloque
    "DATOS DEL TRABAJADOR / ASPIRANTE" (o variantes), para no mezclar CARGO.

    Estrategia:
      1) Delimitar bloque de datos personales: inicia en 'DATOS DEL TRABAJADOR…'
         y termina al encontrar algún encabezado (CARGO, CONCEPTO, EPS, ARL, AFP, NIT, DIRECCIÓN, RECOMENDACIONES,...).
      2) Dentro del bloque:
         A) Si hay rótulo 'Apellidos y Nombres', tomar la siguiente línea no vacía como nombre.
         B) Si no hay rótulo, buscar línea en MAYÚSCULAS (≥2 palabras) que no contenga tokens típicos de CARGO.
         C) Documento (CC/TI/CE/PT) se busca en la MISMA línea o ±6 líneas del nombre.
      3) Si no hay nombre convincente, devolver vacío (mejor vacío que cargo falso).
    """
    raw_lines = text.splitlines()
    lines = [re.sub(r"\s{2,}", " ", l.strip()) for l in raw_lines]

    start_markers = [
        "DATOS DEL TRABAJADOR / ASPIRANTE",
        "DATOS DEL TRABAJADOR",
        "DATOS DEL TRABAJADOR/ASPIRANTE",
        "DATOS DEL TRABAJADOR O ASPIRANTE",
    ]
    stop_markers = [
        "CONCEPTO DE APTITUD", "CONCEPTO MÉDICO OCUPACIONAL", "CONCEPTO MEDICO OCUPACIONAL",
        "CARGO", "EPS", "ARL", "AFP", "NIT", "DIRECCIÓN", "DIRECCION",
        "EXÁMENES", "EXAMENES", "OBSERVACIONES", "RECOMENDACIONES",
    ]

    # Documento flexible (CC/C.C./TI/CE/PT) + número
    doc_re = re.compile(r"\b(C\.?C\.?|TI|CE|PT)\s+([0-9A-Z.\- ]{5,})\b", re.I)

    # Palabras típicas de cargo — si aparecen, NO es nombre
    cargo_tokens = {
        "GENERADOR", "GENERADORA", "OPERARIO", "OPERARIA", "AUXILIAR", "ASEO",
        "OFICIOS", "VARIOS", "CONDUCTOR", "VENDEDOR", "VENDEDORA", "MENSAJERO",
        "JEFE", "SUPERVISOR", "SUPERVISORA", "APRENDIZ", "COORDINADOR", "COORDINADORA",
        "PROFESIONAL", "TECNICO", "TÉCNICO", "AYUDANTE", "MANTENIMIENTO"
    }

    def clean(s: str) -> str:
        return re.sub(r"\s{2,}", " ", s).strip()

    def titlecase(s: str) -> str:
        return " ".join(w.capitalize() for w in s.split())

    def looks_like_upper_name(s: str) -> bool:
        su = s.upper()
        if len(su) < 6:
            return False
        if not re.fullmatch(r"[A-ZÁÉÍÓÚÑ ]{6,}", su):
            return False
        toks = set(t for t in su.split() if len(t) > 2)
        if toks & cargo_tokens:
            return False
        return len([t for t in su.split() if len(t) > 1]) >= 2

    # 1) Delimitar bloque
    start_idx = -1
    for i, ln in enumerate(lines):
        if any(m in ln.upper() for m in start_markers):
            start_idx = i
            break
    if start_idx == -1:
        return "", ""  # no hay bloque claro → evitar falsos positivos

    end_idx = len(lines)
    for j in range(start_idx + 1, len(lines)):
        if any(m in lines[j].upper() for m in stop_markers):
            end_idx = j
            break

    block = lines[start_idx:end_idx]
    if not block:
        return "", ""

    # Helper: buscar doc cerca de índice
    def find_doc_near(idx_base: int) -> str:
        for off in range(0, 7):  # mismo, ±6
            for j in (idx_base - off, idx_base + off):
                if 0 <= j < len(block):
                    m = doc_re.search(block[j])
                    if m:
                        return f"{m.group(1).upper().replace('.', '')} {clean(m.group(2))}"
        return ""

    # A) Rótulo explícito
    for i, ln in enumerate(block):
        if "APELLIDOS Y NOMBRES" in ln.upper():
            for k in range(i + 1, min(i + 8, len(block))):
                cand = block[k].strip()
                if cand and looks_like_upper_name(cand):
                    name = titlecase(cand)
                    doc = find_doc_near(k)
                    return name, doc

    # B) Nombre + doc misma línea
    for i, ln in enumerate(block):
        m = re.search(r"^([A-ZÁÉÍÓÚÑ ]{6,}).*?\b(C\.?C\.?|TI|CE|PT)\s+([0-9A-Z.\- ]{5,})", ln)
        if m and looks_like_upper_name(m.group(1)):
            name = titlecase(clean(m.group(1)))
            doc = f"{m.group(2).upper().replace('.', '')} {clean(m.group(3))}"
            return name, doc

    # C) Doc en una línea → buscar nombre cerca
    for i, ln in enumerate(block):
        m = doc_re.search(ln)
        if not m:
            continue
        doc = f"{m.group(1).upper().replace('.', '')} {clean(m.group(2))}"
        for up in range(1, 7):
            j = i - up
            if j < 0:
                break
            cand = block[j].strip()
            if looks_like_upper_name(cand):
                return titlecase(clean(cand)), doc
        for dn in range(1, 5):
            j = i + dn
            if j >= len(block):
                break
            cand = block[j].strip()
            if looks_like_upper_name(cand):
                return titlecase(clean(cand)), doc
        # Si hay doc pero no nombre convincente, devolver doc al menos
        return "", doc

    # D) Si no se encontró nada confiable dentro del bloque, no forzar.
    return "", ""


def h_find_cargo(text: str) -> str:
    # Cargo en línea propia o etiqueta explícita
    m = re.search(r"^Cargo\s*\n([A-Za-zÁÉÍÓÚÑ ]{3,})", text, re.M)
    if m:
        return normalize_spaces(m.group(1)).title()
    m2 = re.search(r"\n([A-ZÁÉÍÓÚÑ ]{3,})\nNIT\b", text)
    if m2:
        return normalize_spaces(m2.group(1)).title()
    return ""


def h_find_concepto(text: str) -> str:
    m = re.search(r"CONCEPTO DE APTITUD OCUPACIONAL\s*\n([A-ZÁÉÍÓÚÑ .\-]+)", text)
    return normalize_spaces(m.group(1)).title() if m else ""


def h_find_block(text: str, start_key: str, stop_keys: List[str]) -> str:
    idx = text.find(start_key)
    if idx == -1:
        return ""
    sub = text[idx + len(start_key): idx + len(start_key) + 1200]
    for sk in stop_keys:
        cut = sub.find(sk)
        if cut != -1:
            sub = sub[:cut]
    lines = [normalize_spaces(x.strip(" -•\t")) for x in sub.splitlines()]
    lines = [l for l in lines if l]
    return "; ".join(lines)


def h_find_examenes(text: str) -> str:
    return h_find_block(
        text,
        "El concepto de Aptitud se definió a part",
        [
            "RECOMENDACIONES MÉDICAS",
            "RECOMENDACIONES OCUPACIONALES",
            "HABITOS Y ESTILO",
            "OTRAS OBSERVACIONES",
            "Consentimiento informado",
        ],
    )


def h_find_reco_medicas(text: str) -> str:
    return h_find_block(
        text,
        "RECOMENDACIONES MÉDICAS",
        [
            "RECOMENDACIONES OCUPACIONALES",
            "HABITOS Y ESTILO",
            "OTRAS OBSERVACIONES",
            "Consentimiento informado",
        ],
    )


def h_find_reco_ocup(text: str) -> str:
    return h_find_block(
        text,
        "RECOMENDACIONES OCUPACIONALES",
        [
            "HABITOS Y ESTILO",
            "OTRAS OBSERVACIONES",
            "Consentimiento informado",
        ],
    )


def h_find_habitos(text: str) -> str:
    return h_find_block(
        text,
        "HABITOS Y ESTILO DE VIDA SALUDABLES",
        ["OTRAS OBSERVACIONES", "Consentimiento informado"],
    )


# -----------------------------
# Estructuras de datos
# -----------------------------

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


# -----------------------------
# Worker en hilo
# -----------------------------

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

        # Precompilar regex del usuario
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


# -----------------------------
# Interfaz (ORIGINAL: tabla de campos + presets)
# -----------------------------

class FieldTable(QtWidgets.QTableWidget):
    COL_NAME = 0
    COL_PATTERN = 1

    def __init__(self, parent=None):
        super().__init__(0, 2, parent)
        self.setHorizontalHeaderLabels(["Campo", "Patrón (regex opcional)"])
        self.horizontalHeader().setStretchLastSection(True)
        self.verticalHeader().setVisible(False)
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.setEditTriggers(QtWidgets.QAbstractItemView.DoubleClicked | QtWidgets.QAbstractItemView.SelectedClicked | QtWidgets.QAbstractItemView.EditKeyPressed)

    def add_row(self, name: str = "", pattern: str = ""):
        r = self.rowCount()
        self.insertRow(r)
        self.setItem(r, self.COL_NAME, QtWidgets.QTableWidgetItem(name))
        self.setItem(r, self.COL_PATTERN, QtWidgets.QTableWidgetItem(pattern))

    def remove_selected(self):
        rows = sorted({i.row() for i in self.selectedIndexes()}, reverse=True)
        for r in rows:
            self.removeRow(r)

    def to_rules(self) -> List[FieldRule]:
        rules = []
        for r in range(self.rowCount()):
            name_item = self.item(r, self.COL_NAME)
            pat_item = self.item(r, self.COL_PATTERN)
            name = name_item.text().strip() if name_item else ""
            pat = pat_item.text().strip() if pat_item else ""
            if name:
                rules.append(FieldRule(name=name, pattern=pat))
        return rules

    def load_rules(self, rules: List[FieldRule]):
        self.setRowCount(0)
        for fr in rules:
            self.add_row(fr.name, fr.pattern)


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_TITLE} · v{VERSION}")
        self.resize(1050, 720)
        self.worker: Optional[ExtractWorker] = None
        self._setup_ui()
        self._apply_dark_palette()

    def _setup_ui(self):
        w = QtWidgets.QWidget()
        self.setCentralWidget(w)
        v = QtWidgets.QVBoxLayout(w)
        v.setContentsMargins(12, 12, 12, 12)
        v.setSpacing(10)

        # Archivos
        box_paths = QtWidgets.QGroupBox("Archivo de entrada y salida")
        v_paths = QtWidgets.QGridLayout(box_paths)
        self.ed_pdf = QtWidgets.QLineEdit()
        self.btn_pdf = QtWidgets.QPushButton("Seleccionar PDF…")
        self.btn_pdf.clicked.connect(self._choose_pdf)
        self.ed_out = QtWidgets.QLineEdit()
        self.btn_out = QtWidgets.QPushButton("Guardar Excel como…")
        self.btn_out.clicked.connect(self._choose_out)
        v_paths.addWidget(QtWidgets.QLabel("PDF:"), 0, 0)
        v_paths.addWidget(self.ed_pdf, 0, 1)
        v_paths.addWidget(self.btn_pdf, 0, 2)
        v_paths.addWidget(QtWidgets.QLabel("Excel (.xlsx):"), 1, 0)
        v_paths.addWidget(self.ed_out, 1, 1)
        v_paths.addWidget(self.btn_out, 1, 2)
        v.addWidget(box_paths)

        # Campos
        box_fields = QtWidgets.QGroupBox("Campos a extraer (agrega/edita los que necesites)")
        v_fields = QtWidgets.QVBoxLayout(box_fields)
        self.tbl = FieldTable()
        v_fields.addWidget(self.tbl)
        btns = QtWidgets.QHBoxLayout()
        self.btn_add = QtWidgets.QPushButton("+ Añadir campo")
        self.btn_del = QtWidgets.QPushButton("– Quitar seleccionado")
        self.btn_preset = QtWidgets.QPushButton("Plantilla Ocupacional")
        self.btn_load_json = QtWidgets.QPushButton("Cargar preset JSON…")
        self.btn_save_json = QtWidgets.QPushButton("Guardar preset JSON…")
        self.btn_add.clicked.connect(lambda: self.tbl.add_row("Nuevo campo", ""))
        self.btn_del.clicked.connect(self.tbl.remove_selected)
        self.btn_preset.clicked.connect(self._load_default_template)
        self.btn_load_json.clicked.connect(self._load_json)
        self.btn_save_json.clicked.connect(self._save_json)
        for b in (self.btn_add, self.btn_del, self.btn_preset, self.btn_load_json, self.btn_save_json):
            btns.addWidget(b)
        btns.addStretch(1)
        v_fields.addLayout(btns)
        v.addWidget(box_fields)

        # Opciones
        box_opts = QtWidgets.QGroupBox("Opciones")
        h_opts = QtWidgets.QGridLayout(box_opts)
        self.chk_template = QtWidgets.QCheckBox("Usar heurísticas de plantilla para campos vacíos")
        self.chk_template.setChecked(True)
        self.spin_max = QtWidgets.QSpinBox()
        self.spin_max.setRange(1, 2000)
        self.spin_max.setValue(2000)
        self.ed_sheet = QtWidgets.QLineEdit("Datos")
        self.chk_page = QtWidgets.QCheckBox("Agregar columna ‘Página PDF’")
        self.chk_page.setChecked(True)
        h_opts.addWidget(self.chk_template, 0, 0, 1, 2)
        h_opts.addWidget(QtWidgets.QLabel("Máx. páginas a procesar:"), 1, 0)
        h_opts.addWidget(self.spin_max, 1, 1)
        h_opts.addWidget(QtWidgets.QLabel("Nombre de hoja Excel:"), 2, 0)
        h_opts.addWidget(self.ed_sheet, 2, 1)
        h_opts.addWidget(self.chk_page, 3, 0, 1, 2)
        v.addWidget(box_opts)

        # Proceso
        box_run = QtWidgets.QGroupBox("Ejecución")
        v_run = QtWidgets.QVBoxLayout(box_run)
        self.progress = QtWidgets.QProgressBar()
        self.progress.setRange(0, 100)
        self.lbl_status = QtWidgets.QLabel("Listo.")
        h_run_btns = QtWidgets.QHBoxLayout()
        self.btn_start = QtWidgets.QPushButton("Analizar y Exportar")
        self.btn_cancel = QtWidgets.QPushButton("Cancelar")
        self.btn_cancel.setEnabled(False)
        self.btn_start.clicked.connect(self._start)
        self.btn_cancel.clicked.connect(self._cancel)
        h_run_btns.addWidget(self.btn_start)
        h_run_btns.addWidget(self.btn_cancel)
        h_run_btns.addStretch(1)
        self.log = QtWidgets.QPlainTextEdit(); self.log.setReadOnly(True); self.log.setMaximumBlockCount(2000)
        v_run.addWidget(self.progress)
        v_run.addWidget(self.lbl_status)
        v_run.addLayout(h_run_btns)
        v_run.addWidget(self.log)
        v.addWidget(box_run)

        self._load_default_template()

    def _apply_dark_palette(self):
        dark = QtGui.QPalette()
        dark.setColor(QtGui.QPalette.Window, QtGui.QColor(37, 37, 38))
        dark.setColor(QtGui.QPalette.WindowText, QtCore.Qt.white)
        dark.setColor(QtGui.QPalette.Base, QtGui.QColor(30, 30, 30))
        dark.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor(45, 45, 48))
        dark.setColor(QtGui.QPalette.ToolTipBase, QtCore.Qt.white)
        dark.setColor(QtGui.QPalette.ToolTipText, QtCore.Qt.white)
        dark.setColor(QtGui.QPalette.Text, QtCore.Qt.white)
        dark.setColor(QtGui.QPalette.Button, QtGui.QColor(45, 45, 48))
        dark.setColor(QtGui.QPalette.ButtonText, QtCore.Qt.white)
        dark.setColor(QtGui.QPalette.BrightText, QtCore.Qt.red)
        dark.setColor(QtGui.QPalette.Highlight, QtGui.QColor(38, 79, 120))
        dark.setColor(QtGui.QPalette.HighlightedText, QtCore.Qt.white)
        self.setPalette(dark)

    # ---- Callbacks ----
    def _choose_pdf(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Selecciona PDF", "", "PDF (*.pdf)")
        if path:
            self.ed_pdf.setText(path)
            base = os.path.splitext(os.path.basename(path))[0]
            out = os.path.join(os.path.dirname(path), f"{base}_extract.xlsx")
            self.ed_out.setText(out)

    def _choose_out(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Guardar Excel como", self.ed_out.text() or "salida.xlsx", "Excel (*.xlsx)")
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.ed_out.setText(path)

    def _load_default_template(self):
        self.tbl.load_rules(DEFAULT_TEMPLATE)
        self._log("Plantilla ocupacional cargada (puedes editarla o añadir campos).")

    def _load_json(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Cargar preset JSON", "", "JSON (*.json)")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            rules = [FieldRule.from_dict(x) for x in data.get("fields", [])]
            self.tbl.load_rules(rules)
            self._log(f"Preset cargado: {os.path.basename(path)}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, APP_TITLE, f"No se pudo cargar el preset: {e}")

    def _save_json(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Guardar preset JSON", "preset.json", "JSON (*.json)")
        if not path:
            return
        try:
            rules = self.tbl.to_rules()
            data = {"fields": [r.to_dict() for r in rules]}
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self._log(f"Preset guardado en {path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, APP_TITLE, f"No se pudo guardar el preset: {e}")

    def _start(self):
        pdf_path = self.ed_pdf.text().strip()
        out_xlsx = self.ed_out.text().strip()
        if not pdf_path or not os.path.isfile(pdf_path):
            QtWidgets.QMessageBox.warning(self, APP_TITLE, "Selecciona un archivo PDF válido.")
            return
        if not out_xlsx:
            QtWidgets.QMessageBox.warning(self, APP_TITLE, "Indica la ruta de salida del Excel.")
            return

        fields = self.tbl.to_rules()
        if not fields:
            QtWidgets.QMessageBox.warning(self, APP_TITLE, "Agrega al menos un campo de extracción.")
            return

        self.worker = ExtractWorker(
            pdf_path=pdf_path,
            out_xlsx=out_xlsx,
            fields=fields,
            sheet_name=self.ed_sheet.text().strip() or "Datos",
            use_template_heuristics=self.chk_template.isChecked(),
            max_pages=self.spin_max.value(),
            include_pdf_page=self.chk_page.isChecked(),
        )
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_finished)
        self.worker.error.connect(self._on_error)

        self.btn_start.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.progress.setValue(0)
        self._log("Iniciando procesamiento…")
        self.worker.start()

    def _cancel(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self._log("Cancelando…")

    def _on_progress(self, pct: int, msg: str):
        self.progress.setValue(pct)
        self.lbl_status.setText(msg)
        if msg:
            self._log(msg)

    def _on_finished(self, path: str):
        self.btn_start.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self._log(f"✅ Exportado: {path}")
        QtWidgets.QMessageBox.information(self, APP_TITLE, f"Exportación completada:\n{path}")

    def _on_error(self, err: str):
        self.btn_start.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self._log(f"❌ {err}")
        QtWidgets.QMessageBox.critical(self, APP_TITLE, err)

    def _log(self, s: str):
        ts = time.strftime("%H:%M:%S")
        self.log.appendPlainText(f"[{ts}] {s}")


def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    mw = MainWindow()
    mw.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
