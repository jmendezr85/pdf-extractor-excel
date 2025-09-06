import re
from typing import List, Tuple

def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()



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
