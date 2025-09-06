import json
import os
import time
from typing import List, Optional

from PySide6 import QtCore, QtGui, QtWidgets

from worker import FieldRule, ExtractWorker, DEFAULT_TEMPLATE

APP_TITLE = "Extractor PDF → Excel"
VERSION = "1.1.0"  # Reinstala interfaz original + fix nombres/documentos

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

