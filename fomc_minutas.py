#!/usr/bin/env python3
"""
fomc_minutas.py
Comparativo de minutas del FOMC: evolución de la percepción de inflación,
mercado laboral, actividad, estabilidad financiera y política monetaria.

Flujo:
  1. Desc
  2. Parsea cada minuta en secciones y párrafos (la estructura de la Fed es fija).
  3. Parte en oraciones y cláusulas; asigna tema(s), cuantificador ("a few",
     "most", ...) y dirección hawkish/dovish con un diccionario por tema.
  4. Extrae variables duras: rango objetivo, votos, disidentes y su dirección,
     PCE / subyacente / desempleo citados por el staff, balance de riesgos.
  5. Genera: Excel con varias hojas, CSVs, gráficas PNG, un HTML con el diff
     palabra por palabra entre comunicados consecutivos y la lámina resumen
     interactiva (panorama_fed_minutas.html, Plotly, funciona sin internet).

Uso:
  pip install requests beautifulsoup4 pandas matplotlib openpyxl plotly
  python fomc_minutas.py                       # 2025 y 2026, detecta fechas solo
  python fomc_minutas.py --anio 2026           # solo un año (o varios: --anio 2024 2025 2026)
  python fomc_minutas.py --fechas 20260128 20260318
  python fomc_minutas.py --solo-local          # no descarga, usa ./minutas_html
  python fomc_minutas.py --llm                 # capa opcional con Claude (API)
  python fomc_minutas.py --desde-excel salida_fomc/comparativo_minutas_fomc.xlsx
                                               # regenera tablas y gráficas sin descargar

Notas:
  - El sentiment por diccionario es transparente pero aproximado. Revisa la hoja
    "oraciones" (trae la evidencia) antes de publicar conclusiones.
  - El mapeo de cuantificadores a proporción del comité es una convención
    aproximada (la Fed no publica una tabla oficial). Se puede ajustar abajo.
"""

from __future__ import annotations

import argparse
import difflib
import html as html_lib
import json
import os
import re
import sys
import time
from pathlib import Path

import pandas as pd

# ----------------------------------------------------------------------------
# Configuración
# ----------------------------------------------------------------------------
URL_MINUTA = "https://www.federalreserve.gov/monetarypolicy/fomcminutes{fecha}.htm"
URL_CALENDARIO = "https://www.federalreserve.gov/monetarypolicy/fomccalendars.htm"
HEADERS = {"User-Agent": "Mozilla/5.0 (analisis macro; uso academico)"}

N_PARTICIPANTES = 19   # gobernadores + presidentes de bancos de reserva
N_VOTANTES = 12

SECCIONES = [
    ("organizacion", r"annual organizational matters"),
    ("mercados_desk", r"developments in financial markets and open market operations"),
    ("staff_economia", r"staff review of the economic situation"),
    ("staff_financiero", r"staff review of the financial situation"),
    ("staff_financiero", r"staff review and policymakers' discussion of financial stability"),
    ("staff_perspectivas", r"staff economic outlook"),
    ("participantes", r"participants.? views on current conditions and the economic outlook"),
    ("participantes", r"policymakers' discussion of current conditions, the economic outlook, and monetary policy considerations"),
    ("acciones_politica", r"committee policy actions"),
    ("votacion_notacion", r"notation votes?"),
    ("asistencia", r"attendance"),
]

NIVEL_POR_SECCION = {
    "mercados_desk": "mercado",
    "staff_economia": "staff",
    "staff_financiero": "staff",
    "staff_perspectivas": "staff",
    "participantes": "participantes",
    "acciones_politica": "comite",
}

# Cuantificadores -> proporción aproximada del comité
CUANTIFICADORES = {
    "all": 1.00, "almost all": 0.90, "nearly all": 0.90,
    "the vast majority of": 0.85, "vast majority of": 0.85,
    "generally": 0.85,          # "participants generally ..."
    "most": 0.70,
    "a majority of": 0.55, "the majority of": 0.55,
    "many": 0.45,
    "several": 0.30,
    "some": 0.25, "various": 0.25, "a number of": 0.25,
    "a few": 0.15,
    "a couple of": 0.10,
}
NUMEROS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
           "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11, "twelve": 12}

_q_alts = sorted(list(CUANTIFICADORES) + list(NUMEROS), key=len, reverse=True)
_q_alts = [q for q in _q_alts if q != "generally"]
RE_CUANT = re.compile(
    r"\b(?P<q>" + "|".join(re.escape(q) for q in _q_alts) + r")\s+"
    r"(?:of\s+)?(?:these\s+|those\s+|the\s+)?(?:other\s+)?"
    r"(?P<actor>participants?|members?)\b",
    re.I,
)
RE_GENERALLY = re.compile(r"\b(participants|members)\s+generally\b", re.I)
RE_BARE = re.compile(r"^(participants|members)\b", re.I)          # "Participants noted..."
RE_ELIPTICO = re.compile(
    r"^(?:but|while|whereas|and|however,)?\s*(?P<q>a few|a couple|several|many|most|some)\b"
    r"(?!\s+(?:of|participants|members|other participants))\s+(?:other\s+)?[a-z]+(?:ed|s)\b", re.I)
RE_HEREDA = re.compile(r"^(these|they)\b|\bthese participants\b", re.I)
RE_ANIDADO = re.compile(r"\bof\s+these\s+(?:participants|members)\b", re.I)
PESO_SIN_CUANTIFICADOR = 0.5  # peso de una cláusula de participantes sin cuantificador identificable
K_SUAVIZADO = 3.0             # sube este valor para suavizar más el índice
MIN_CLAUSULAS = 5             # celdas con menos cláusulas se atenúan en las gráficas

# Motores de la narrativa: menciones por cada 10 mil palabras de la minuta
MOTORES = {
    "Aranceles": r"(?i)\btariffs?\b",
    "IA": r"\bAI\b|(?i:artificial intelligence)",
    "Medio Oriente / energía": r"Middle East|Strait of Hormuz|(?i:\boil prices?\b|energy prices?)",
    "Crédito privado": r"(?i)private credit",
    "Cierre de gobierno": r"(?i)shutdown",
    "Inmigración": r"(?i)immigration",
}

# Temas: patrones que indican de qué habla una cláusula
TEMAS = {
    "inflacion": [r"\binflation", r"\bprices?\b", r"\bPCE\b", r"\bCPI\b", r"\btariffs?\b",
                  r"disinflation", r"price pressures?", r"cost pressures?", r"input costs?",
                  r"price stability", r"price increases?"],
    "mercado_laboral": [r"labor market", r"unemployment", r"payrolls?", r"\bjobs?\b", r"hiring",
                        r"layoffs?", r"\bwages?\b", r"labor demand", r"labor force", r"\bworkers\b",
                        r"(?<!maximum[ -])\bemployment\b"],
    "actividad": [r"\bGDP\b", r"economic activity", r"consumer spending", r"household spending",
                  r"business (?:fixed )?investment", r"(?:economic|GDP|output) growth",
                  r"pace of (?:real GDP )?growth", r"\boutput\b", r"productivity",
                  r"aggregate demand", r"economic outlook"],
    "estabilidad_financiera": [r"financial stability", r"vulnerabilit", r"valuations?", r"leverage",
                               r"private credit", r"hedge funds?", r"credit spreads?",
                               r"cybersecurity", r"financial system"],
    "politica_monetaria": [r"federal funds rate", r"target range", r"policy rate",
                           r"stance of (?:monetary )?policy", r"policy stance", r"monetary policy",
                           r"rate (?:cut|hike|increase|reduction)s?", r"easing bias",
                           r"postmeeting statement", r"policy firming", r"\btightening\b",
                           r"\bfirming\b", r"\beasing\b", r"restrictive",
                           r"interest rate decisions", r"two-sided"],
}

# Dirección por tema. +1 = hawkish (más presión a subir tasas), -1 = dovish.
# Se buscan todos los patrones y gana la coincidencia más larga cuando se traslapan
# (así "downward pressure" no cuenta también como "pressure").
DIRECCION = {
    "inflacion": {
        +1: [r"elevated", r"higher(?!-income)", r"\bhigh\b", r"\bros[e]\b|\brising\b|\brisen\b",
             r"increas\w*", r"upside", r"persisten\w*|\bpersist\w*", r"broad[- ]based",
             r"above (?:the Committee's )?2 percent|above[- ]target", r"upward pressures?",
             r"price pressures?|cost pressures?", r"moved (?:up|higher)|picked up|stepped up",
             r"de-?anchor\w*", r"entrenched", r"accelerat\w*", r"more persistent",
             r"slower (?:and more uneven )?than (?:generally |previously )?expected",
             r"take longer to return", r"delay(?:d)? the (?:anticipated|expected) (?:decline|return)"],
        -1: [r"declin\w*", r"\beas(?:e|ed|ing)\b", r"moderat\w*(?!-income)", r"lower(?!-income)",
             r"downward pressures?", r"disinflation\w*", r"\bslow(?:ed|ing)?\b", r"diminish\w*",
             r"\bwan(?:e|ed|ing)\b", r"\bfad(?:e|ed|ing)\b", r"(?:well )?anchored",
             r"step(?:ped)? down|move(?:d)? down|edge(?:d)? down",
             r"(?:return|move) (?:to|toward) (?:the Committee's )?2 percent",
             r"largely complete", r"absorbing", r"resist further price increases",
             r"limited to select categories"],
    },
    "mercado_laboral": {
        +1: [r"strong\w*", r"\bsolid\b", r"stabl\w*|stabiliz\w*", r"\btight\w*", r"robust",
             r"downside risks? (?:to (?:the )?(?:labor market|employment) )?(?:had )?(?:moderated|diminished|eased)",
             r"concerns? .{0,60}(?:had )?eased", r"momentum", r"notable increases in their wages",
             r"upside risks to wage growth"],
        -1: [r"soft\w*", r"\bweak\w*", r"deteriorat\w*", r"downside", r"fragil\w*", r"vulnerab\w*",
             r"(?:remained|were|been) low\b", r"low (?:rate|pace)s? of job", r"cooling",
             r"sharply higher", r"reduce hiring|delay or reduce hiring|implementing layoffs",
             r"low dynamism|relatively low dynamism", r"elevated long-term unemployment",
             r"low job-finding rate", r"caution in hiring"],
    },
    "actividad": {
        +1: [r"\bsolid\b", r"strong\w*", r"robust", r"resilien\w*", r"expan\w*", r"picked up",
             r"outpace\w*", r"supported? by", r"supportive"],
        -1: [r"\bslow\w*", r"\bweak\w*", r"\bsoft\w*", r"contract\w*", r"downside", r"headwinds?",
             r"strains?", r"\bdrag\b", r"disappoint\w*", r"repricing"],
    },
    "estabilidad_financiera": {
        +1: [r"resilien\w*", r"well capitalized", r"\bsolid\b", r"stable", r"orderly"],
        -1: [r"elevated", r"vulnerabilit\w*", r"notable", r"\brisks?\b", r"concerns?",
             r"spill\s?over", r"disruption", r"stress", r"near all-time highs", r"record levels"],
    },
    "politica_monetaria": {
        +1: [r"(?:raise|raising|increase in|increase|increasing) (?:the )?target range",
             r"an increase (?:of \d+ basis points )?in the target range", r"\d+ basis point hike",
             r"rate (?:hike|increase)s?", r"\bhikes?\b", r"\bfirming\b", r"\btighten\w*",
             r"more restrictive", r"not (?:currently )?sufficiently restrictive",
             r"did not see the current policy stance as restrictive",
             r"(?:remov\w*|not (?:to )?repeat|not support inclusion of) .{0,80}easing bias",
             r"two-sided", r"above the current target range",
             r"for longer than previously anticipated", r"will deliver price stability"],
        -1: [r"(?:lower|lowering|reduce|reduction in|cut) (?:the )?target range",
             r"rate (?:cut|reduction)s?", r"\bcuts?\b", r"(?<!removing the language that suggested an )\beasing\b",
             r"downward adjustments?", r"(?:overly|meaningfully|slightly|remained) restrictive",
             r"(?:within or )?(?:slightly )?below the current target range",
             r"further (?:policy )?easing", r"easing bias"],
    },
}

# ----------------------------------------------------------------------------
# Utilidades de texto
# ----------------------------------------------------------------------------
def normalizar(t: str) -> str:
    t = html_lib.unescape(t)
    reemplazos = {"\u00a0": " ", "\u2011": "-", "\u2010": "-", "\u00ad": "",
                  "\u2019": "'", "\u2018": "'", "\u201c": '"', "\u201d": '"',
                  }
    for k, v in reemplazos.items():
        t = t.replace(k, v)
    for k, v in {"½": "1/2", "¼": "1/4", "¾": "3/4"}.items():
        t = re.sub(r"(\d)\s?" + k, r"\1-" + v, t)   # "3½" -> "3-1/2"
        t = t.replace(k, v)                         # "by ¼" -> "by 1/4"
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def identificar_seccion(titulo: str):
    t = normalizar(titulo).lower().strip(" :")
    for clave, patron in SECCIONES:
        if re.fullmatch(patron, t):
            return clave
    return None


_ABREV = ["U.S.", "a.m.", "p.m.", "e.g.", "i.e.", "Jr.", "St.", "No.", "vs."]


def partir_oraciones(texto: str) -> list[str]:
    t = texto
    for a in _ABREV:
        t = t.replace(a, a.replace(".", "§"))
    partes = re.split(r'(?<=[.!?])(?<!\s[A-Z]\.)\s+(?=["(A-Z])', t)
    return [p.replace("§", ".").strip() for p in partes if p.strip()]


def partir_clausulas(oracion: str) -> list[str]:
    """Separa oraciones con contraste (while/but/however/whereas/although/;)."""
    partes = re.split(r";\s+|,\s+(?=(?:while|but|whereas|although|though|however|by contrast)\b)",
                      oracion)
    return [p.strip() for p in partes if len(p.strip()) > 3]


def a_numero(frac: str) -> float:
    """'3-1/2' -> 3.5 ; '3' -> 3.0"""
    m = re.fullmatch(r"(\d+)(?:-(\d)/(\d))?", frac.strip())
    if not m:
        return float("nan")
    base = float(m.group(1))
    if m.group(2):
        base += int(m.group(2)) / int(m.group(3))
    return base


# ----------------------------------------------------------------------------
# 1. Descarga
# ----------------------------------------------------------------------------
def descubrir_fechas(anio: int) -> list[str]:
    import requests
    r = requests.get(URL_CALENDARIO, headers=HEADERS, timeout=30)
    r.raise_for_status()
    fechas = sorted(set(re.findall(r"fomcminutes(\d{8})\.htm", r.text)))
    return [f for f in fechas if f.startswith(str(anio))]


def descargar(fechas: list[str], carpeta: Path) -> dict[str, Path]:
    import requests
    carpeta.mkdir(parents=True, exist_ok=True)
    rutas = {}
    for f in fechas:
        ruta = carpeta / f"fomcminutes{f}.htm"
        if ruta.exists() and ruta.stat().st_size > 20_000:
            rutas[f] = ruta
            continue
        r = requests.get(URL_MINUTA.format(fecha=f), headers=HEADERS, timeout=30)
        if r.status_code == 404:
            print(f"[aviso] {f}: la minuta aún no está publicada (404)")
            continue
        r.raise_for_status()
        ruta.write_bytes(r.content)
        rutas[f] = ruta
        print(f"[ok] descargada {f}")
        time.sleep(1)
    return rutas


# ----------------------------------------------------------------------------
# 2. Parseo -> lista de párrafos {fecha, seccion, n, tipo, texto}
# ----------------------------------------------------------------------------
def parsear_html(html: str, fecha: str) -> list[dict]:
    from bs4 import BeautifulSoup, NavigableString

    soup = BeautifulSoup(html, "html.parser")
    cont = soup.find("div", id="article") or soup.find("div", id="content") or soup.body or soup
    for el in cont.select('a[href^="#fn"], sup, script, style'):
        el.decompose()

    parrafos, seccion, n = [], "preambulo", 0
    for el in cont.find_all(["p", "li", "h2", "h3", "h4", "h5", "h6", "blockquote"]):
        if el.name in ("p", "li") and el.find_parent("blockquote"):
            continue  # el blockquote se procesa completo

        if el.name.startswith("h"):
            clave = identificar_seccion(el.get_text(" "))
            if clave:
                seccion = clave
            if seccion == "asistencia":
                break
            continue

        if el.name == "blockquote":
            if seccion == "preambulo":
                continue
            txt = normalizar(el.get_text(" "))
            tipo = "directiva" if "directs the Desk" in txt else "comunicado"
            n += 1
            parrafos.append(dict(fecha=fecha, seccion=seccion, n=n, tipo=tipo, texto=txt))
            continue

        # <p>: ¿empieza con <strong>Título de sección</strong>?
        hijos = [c for c in el.children if not (isinstance(c, NavigableString) and not c.strip())]
        if hijos and getattr(hijos[0], "name", None) in ("strong", "b"):
            clave = identificar_seccion(hijos[0].get_text(" "))
            if clave:
                seccion_anterior = seccion
                seccion = clave
                hijos[0].extract()
                if seccion == "asistencia":
                    if seccion_anterior == "preambulo":
                        seccion = "preambulo"
                        continue
                    break
        txt = normalizar(el.get_text(" "))
        if not txt or seccion == "preambulo":
            continue
        n += 1
        parrafos.append(dict(fecha=fecha, seccion=seccion, n=n, tipo="parrafo", texto=txt))
    return parrafos


def parsear_texto(texto: str, fecha: str) -> list[dict]:
    """Respaldo para texto extraído de PDF (pdftotext). Menos preciso que el HTML:
    detecta fin de párrafo cuando una línea termina en punto y es más corta."""
    lineas = [l.rstrip() for l in texto.replace("\f", "\n").split("\n")]
    ancho = max((len(l) for l in lineas), default=100)
    parrafos, seccion, n, buffer = [], "preambulo", 0, []

    def volcar():
        nonlocal n, buffer
        if buffer and seccion != "preambulo":
            txt = normalizar(" ".join(buffer).replace("- ", "-"))
            if txt:
                tipo = "parrafo"
                if txt.startswith('"Effective') or "directs the Desk" in txt[:120]:
                    tipo = "directiva"
                n += 1
                parrafos.append(dict(fecha=fecha, seccion=seccion, n=n, tipo=tipo, texto=txt))
        buffer = []

    for l in lineas:
        s = l.strip()
        if not s:
            continue
        if re.fullmatch(r"\d{1,2}|\d+ [A-Z][a-z]+ \d.*\d{4}|Minutes of the Federal Open Market Committee \d+", s):
            continue  # encabezados / números de página
        if re.match(r"^\d \S", s) and ("referenced as" in s or "Attended" in s):
            continue  # notas al pie
        clave = identificar_seccion(s)
        if clave:
            volcar()
            seccion = clave
            if seccion == "asistencia":
                break
            continue
        buffer.append(s)
        if re.search(r'[.:"\u201d]$', s) and not re.search(r"\b[A-Z]\.$", s) and len(s) < 0.85 * ancho:
            volcar()
    volcar()
    return _fusionar_citas_texto(parrafos)


def _fusionar_citas_texto(parrafos: list[dict]) -> list[dict]:
    """En texto de PDF, la directiva y el comunicado salen en varios párrafos: se unen."""
    out, modo, acumulado = [], None, None
    for p in parrafos:
        t = p["texto"]
        if p["seccion"] == "acciones_politica":
            if "directs the Desk" in t[:150]:
                modo, acumulado = "directiva", dict(p, tipo="directiva")
                out.append(acumulado)
                continue
            if "approved the following statement" in t or (modo == "directiva" and t.startswith("The vote also")):
                modo = "espera_comunicado" if "approved the following" not in t else "comunicado"
                if modo == "comunicado":
                    acumulado = dict(p, tipo="comunicado", texto=t)
                    out.append(acumulado)
                else:
                    out.append(p)
                continue
            if modo == "espera_comunicado" and "approved the following statement" in t:
                modo = "comunicado"
            if t.startswith("Voting for") or t.startswith("Voting against"):
                modo = None
            elif modo in ("directiva", "comunicado"):
                acumulado["texto"] += " " + t
                continue
        out.append(p)
    return out


# ----------------------------------------------------------------------------
# 3. Oraciones, cláusulas, temas, cuantificadores y dirección
# ----------------------------------------------------------------------------
_TEMAS_RE = {t: [re.compile(p, re.I) for p in ps] for t, ps in TEMAS.items()}
_DIR_RE = {t: [(s, re.compile(p, re.I)) for s, ps in d.items() for p in ps]
           for t, d in DIRECCION.items()}


def temas_de(texto: str) -> list[str]:
    return [t for t, pats in _TEMAS_RE.items() if any(p.search(texto) for p in pats)]


def direccion(texto: str, tema: str):
    """Devuelve (signo, términos_hawk, términos_dove) con coincidencias no traslapadas."""
    hallazgos = []
    for signo, pat in _DIR_RE[tema]:
        for m in pat.finditer(texto):
            hallazgos.append((m.start(), m.end(), signo, m.group(0)))
    hallazgos.sort(key=lambda h: (-(h[1] - h[0]), h[0]))
    ocupado, hawk, dove = set(), [], []
    for ini, fin, signo, txt in hallazgos:
        if any(i in ocupado for i in range(ini, fin)):
            continue
        ocupado.update(range(ini, fin))
        (hawk if signo > 0 else dove).append(txt.lower())
    neto = len(hawk) - len(dove)
    return (1 if neto > 0 else -1 if neto < 0 else 0), hawk, dove


def cuantificador(texto: str):
    """Devuelve (etiqueta, proporción) o (None, None) si no hay."""
    m = RE_CUANT.search(texto)
    if m:
        q, actor = m.group("q").lower(), m.group("actor").lower()
        if q in NUMEROS:
            total = N_VOTANTES if actor.startswith("member") else N_PARTICIPANTES
            return f"{q} {actor}", round(min(NUMEROS[q] / total, 1), 2)
        return f"{q} {actor}", CUANTIFICADORES[q]
    m = RE_GENERALLY.search(texto)
    if m:
        return f"{m.group(1).lower()} generally", CUANTIFICADORES["generally"]
    m = RE_BARE.search(texto)
    if m:
        return f"{m.group(1).lower()} (sin calificar)", 1.0
    m = RE_ELIPTICO.search(texto)
    if m:
        q = m.group("q").lower()
        q = "a couple of" if q == "a couple" else q
        return f"{q} (elíptico)", CUANTIFICADORES[q]
    return None, None


def resolver_cuantificador(texto: str, padre):
    """Cuantificador de un texto. Si es anidado ("almost all of these participants"),
    se multiplica por la proporción del grupo padre."""
    etq, prop = cuantificador(texto)
    if etq and RE_ANIDADO.search(texto) and padre[1]:
        return f"{etq} [de {padre[0]}]", round(prop * padre[1], 2), True
    return etq, prop, False


def analizar(parrafos: list[dict]) -> pd.DataFrame:
    filas = []
    for p in parrafos:
        nivel = NIVEL_POR_SECCION.get(p["seccion"])
        if p["tipo"] == "comunicado":
            nivel = "comunicado"
        if nivel is None or p["tipo"] == "directiva":
            continue
        heredado = (None, None)
        for i_o, oracion in enumerate(partir_oraciones(p["texto"])):
            etq, prop, anidado = resolver_cuantificador(oracion, heredado)
            q_or = (etq, prop)
            if etq is None and RE_HEREDA.search(oracion):
                q_or = heredado
            if etq is not None and not anidado:
                heredado = q_or
            for i_c, cl in enumerate(partir_clausulas(oracion)):
                etq_c, prop_c, _ = resolver_cuantificador(cl, heredado)
                q_cl = (etq_c, prop_c) if etq_c else q_or
                if nivel in ("participantes", "comite"):
                    peso = q_cl[1] if q_cl[1] else PESO_SIN_CUANTIFICADOR
                else:
                    peso = 1.0
                for tema in temas_de(cl):
                    signo, hawk, dove = direccion(cl, tema)
                    filas.append(dict(
                        fecha=p["fecha"], seccion=p["seccion"], nivel=nivel, parrafo=p["n"],
                        oracion_n=i_o, clausula_n=i_c, tema=tema,
                        cuantificador=q_cl[0], proporcion=q_cl[1] if nivel in ("participantes", "comite") else None,
                        signo=signo, peso=peso, puntaje=signo * peso,
                        terminos_hawk=", ".join(hawk), terminos_dove=", ".join(dove),
                        clausula=cl, oracion=oracion,
                    ))
    return pd.DataFrame(filas)


def panel_por_tema(df: pd.DataFrame) -> pd.DataFrame:
    """Índice neto por fecha x nivel x tema, en [-1, 1].
    indice_neto = sum(signo*peso) / sum(|signo|*peso)   (solo cláusulas con dirección)
    tono_medio  = sum(signo*peso) / n_clausulas         (incluye neutrales)"""
    g = df.assign(abs_peso=df["signo"].abs() * df["peso"])
    out = (g.groupby(["fecha", "nivel", "tema"])
             .agg(n_clausulas=("signo", "size"),
                  n_hawk=("signo", lambda s: int((s > 0).sum())),
                  n_dove=("signo", lambda s: int((s < 0).sum())),
                  suma=("puntaje", "sum"), base=("abs_peso", "sum"))
             .reset_index())
    out["indice_neto"] = (out["suma"] / out["base"]).where(out["base"] > 0).round(3)
    # Suavizado: una sola frase ya no lleva el índice a ±1
    out["indice_suavizado"] = (out["suma"] / (out["base"] + K_SUAVIZADO)).round(3)
    out["tono_medio"] = (out["suma"] / out["n_clausulas"]).round(3)
    return out.drop(columns=["suma", "base"])


# ----------------------------------------------------------------------------
# 4. Posturas de política (cuantificador asociado a cada postura)
# ----------------------------------------------------------------------------
POSTURAS = {
    "alza_en_esta_reunion": r"(?:favored|preferred|case for|supported|in favor of) (?:an increase|raising|to raise|increasing)",
    "recorte_en_esta_reunion": r"(?:favored|preferred|supported|in favor of) (?:to lower|lowering|reducing|a reduction|a cut|a decrease"
                               r"|a \d+ basis point (?:cut|reduction)|a \d/\d percentage point (?:cut|reduction))",
    # Solo se grafica cuando el comité recortó: es el lado halcón de esas reuniones
    "mantener_en_esta_reunion": r"(?:favored|preferred|supported|merit in) (?:to keep|keeping|maintaining|no change|leaving)",
    "alza_condicional": r"(?:firming|tightening|rate (?:increase|hike)s?|upward adjustments?)\b.{0,90}\b(?:appropriate|warranted|necessary)"
                        r"|(?:appropriate|warranted|necessary)\b.{0,40}\b(?:firming|tightening)",
    "recorte_condicional": r"(?:lower the target range|rate (?:cut|reduction)s?|downward adjustments?|further (?:policy )?easing|rate reductions)\b.{0,90}\b(?:appropriate|warranted)"
                           r"|(?:appropriate|warranted)\b.{0,40}\b(?:to lower|lower the target range|rate (?:cut|reduction)s?|reduc)",
    "quitar_sesgo_recortes": r"(?:remov\w*|not (?:to )?repeat|not support inclusion of|two-sided)",
}


def posturas(df: pd.DataFrame) -> pd.DataFrame:
    base = df[(df["nivel"] == "participantes") & (df["tema"] == "politica_monetaria")]
    base = base.drop_duplicates(subset=["fecha", "parrafo", "oracion_n", "clausula_n"])
    filas = []
    for fecha, g in base.groupby("fecha"):
        for postura, patron in POSTURAS.items():
            hit = g[g["clausula"].str.contains(patron, case=False, regex=True)]
            if postura == "quitar_sesgo_recortes":
                hit = hit[hit["clausula"].str.contains("easing bias|two-sided|upward adjustments", case=False)]
            if hit.empty:
                filas.append(dict(fecha=fecha, postura=postura, proporcion_max=0.0, cuantificador=None, evidencia=None))
                continue
            top = hit.sort_values("proporcion", ascending=False).iloc[0]
            filas.append(dict(fecha=fecha, postura=postura, proporcion_max=top["proporcion"],
                              cuantificador=top["cuantificador"], evidencia=top["oracion"]))
    return pd.DataFrame(filas)


# ----------------------------------------------------------------------------
# 5. Variables duras por reunión
# ----------------------------------------------------------------------------
RE_NOMBRE = re.compile(r"\b([A-Z][a-z]+(?:\s[A-Z]\.)?(?:\s[A-Z][a-zA-Z'\-]+)+)")
NO_NOMBRE = {"Federal", "Reserve", "Committee", "Board", "Open", "Market", "Voting", "None",
             "System", "Governors", "Consistent", "Chair", "Vice", "Governor", "President",
             "Chairman", "Bank"}


def _nombres(texto: str) -> list[str]:
    vistos = []
    for m in RE_NOMBRE.finditer(texto):
        nom = m.group(1)
        if set(nom.split()) & NO_NOMBRE:
            continue
        if nom not in vistos:
            vistos.append(nom)
    return vistos


def _direccion_disidente(segmento: str) -> str:
    """La guía/sesgo tiene prioridad; si no, gana la primera acción mencionada
    ("preferred to lower ..." -> recorte)."""
    s = segmento.lower()
    if re.search(r"easing bias|two-sided|characterization", s):
        return "sesgo/guía"
    if re.search(r"securities holdings|balance sheet|runoff|redemption", s):
        return "balance"
    opciones = [("alza", r"\braise|\bincrease|firming|tighten"),
                ("recorte", r"\blower|reduc|\bcut\b"),
                ("pausa", r"no change|maintain|unchanged|leave")]
    hallados = [(m.start(), etq) for etq, pat in opciones for m in [re.search(pat, s)] if m]
    return min(hallados)[1] if hallados else "?"


def _disidentes(acc: list[dict]) -> list[tuple[str, str]]:
    i = next((k for k, p in enumerate(acc) if p["texto"].startswith("Voting against this action")), None)
    if i is None:
        return []
    primero = acc[i]["texto"].split(":", 1)[1].strip()
    if re.match(r"none\b", primero, re.I):
        return []
    nombres = _nombres(primero)
    texto = primero
    # Párrafos siguientes solo si explican el voto de los disidentes ("X preferred...")
    for p in acc[i + 1:]:
        t = p["texto"]
        if t.startswith("Consistent with") or "voted as an alternate" in t:
            break
        apellidos = [n.split()[-1] for n in nombres]
        if any(a in t for a in apellidos) and re.search(r"preferred|dissent|voted against|supported", t):
            texto += " " + t
    salida = []
    for nom in nombres:
        pos = texto.find(nom) + len(nom)
        m = re.search(r"\b(who|preferred|supported|would have)\b", texto[pos:])
        if not m:
            salida.append((nom, "?"))
            continue
        ini = pos + m.start()
        sig = [texto.find(n2, ini) for n2 in nombres if texto.find(n2, ini) > ini]
        salida.append((nom, _direccion_disidente(texto[ini:min(sig) if sig else len(texto)])))
    return salida


def _rango(acc: list[dict]):
    directiva = " ".join(p["texto"] for p in acc if p["tipo"] == "directiva" or "directs the Desk" in p["texto"])
    todo = " ".join(p["texto"] for p in acc)
    for txt in (directiva, todo):
        m = re.search(r"target range (?:for the federal funds rate )?(?:[^.]{0,60}?)(?:at|of|to) "
                      r"(\d+(?:-\d/\d)?) to (\d+(?:-\d/\d)?) percent", txt)
        if m:
            return a_numero(m.group(1)), a_numero(m.group(2))
    return None, None


def _dato(texto: str, patron: str):
    m = re.search(patron, texto)
    return (float(m.group(1)), m.group(2) if m.lastindex and m.lastindex >= 2 else None) if m else (None, None)


def extraer_decision(parrafos: list[dict], fecha: str) -> dict:
    todo = [p for p in parrafos if p["fecha"] == fecha]
    acc = [p for p in todo if p["seccion"] == "acciones_politica"]
    d = dict(fecha=fecha)
    d["rango_inf"], d["rango_sup"] = _rango(acc)

    i_favor = next((i for i, p in enumerate(acc) if p["texto"].startswith("Voting for this action")), None)
    a_favor = _nombres(acc[i_favor]["texto"].split(":", 1)[1]) if i_favor is not None else []
    d["presidente"] = a_favor[0] if a_favor else None
    dis = _disidentes(acc)
    d["votos_a_favor"] = len(a_favor)
    d["votos_en_contra"] = len(dis)
    d["disidentes"] = "; ".join(f"{n} ({x})" for n, x in dis) if dis else "ninguno"
    for tipo in ("alza", "recorte", "pausa", "sesgo/guía", "balance"):
        d[f"disenso_{tipo.split('/')[0]}"] = sum(1 for _, x in dis if x == tipo)

    # Datos del staff (varias redacciones entre años)
    eco = " ".join(p["texto"] for p in todo if p["seccion"] == "staff_economia")
    MES = r"(?:in|over the 12 months ending in) (\w+)"
    d["pce_total_ultimo_dato"], d["pce_mes"] = _dato(eco, r"Total consumer price inflation.{0,220}?(\d+\.\d) percent " + MES)
    if d["pce_total_ultimo_dato"] is None:
        d["pce_total_ultimo_dato"], d["pce_mes"] = _dato(
            eco, r"total personal consumption expenditures \(PCE\)[^.;]{0,180}?"
                r"(?:rose|increased|edged up) to (\d+\.\d) percent in (\w+)")
    d["pce_subyacente_ultimo_dato"], _ = _dato(eco.split("staff estimated")[0], r"(?i)core PCE price inflation.{0,160}?(\d+\.\d) percent")
    if d["pce_subyacente_ultimo_dato"] is None:
        d["pce_subyacente_ultimo_dato"], _ = _dato(
            eco, r"(?i)core PCE price inflation[^.;]{0,180}?"
                r"(?:rose|increased|edged up) to (\d+\.\d) percent in (\w+)")
    m = re.search(r"staff estimated that total PCE price inflation.{0,60}?(\d+\.\d) percent " + MES + r"(.{0,250})", eco)
    if m:
        d["pce_total_estimado"], d["pce_estimado_mes"] = float(m.group(1)), m.group(2)
        mc = re.search(r"core PCE price inflation\D{0,60}?(\d+\.\d) percent", m.group(3))
        d["pce_subyacente_estimado"] = float(mc.group(1)) if mc else None
    else:
        d["pce_total_estimado"] = d["pce_estimado_mes"] = d["pce_subyacente_estimado"] = None
    # Serie para gráficas: el dato más reciente que cita el staff
    d["pce_total"] = d["pce_total_estimado"] if d["pce_total_estimado"] else d["pce_total_ultimo_dato"]
    d["pce_subyacente"] = d["pce_subyacente_estimado"] if d["pce_subyacente_estimado"] else d["pce_subyacente_ultimo_dato"]
    d["desempleo"], d["desempleo_mes"] = _dato(eco, r"unemployment rate\D{0,30}?(\d+\.\d) percent in (\w+)")

    persp = [o for p in todo if p["seccion"] == "staff_perspectivas" for o in partir_oraciones(p["texto"])]

    def frase_riesgo(tema_pat):
        cand = [o for o in persp if re.search(tema_pat, o, re.I)]
        return (next((o for o in cand if re.search(r"skewed|tilted", o)), None)
                or next((o for o in cand if re.search(r"balanced", o)), None))
    d["staff_riesgo_inflacion"] = frase_riesgo(r"inflation")
    d["staff_riesgo_empleo"] = frase_riesgo(r"employment|GDP|real activity|economic activity")
    d["escala_riesgo_inflacion"] = escala_riesgo(d["staff_riesgo_inflacion"])
    d["escala_riesgo_empleo"] = escala_riesgo(d["staff_riesgo_empleo"])

    desk = [o for p in todo if p["seccion"] == "mercados_desk" for o in partir_oraciones(p["texto"])]
    d["desk_expectativas"] = " ".join(o for o in desk if re.search(
        r"rate (?:cut|hike|reduction)s?|priced|modal path|probability", o))[:1200] or None

    com = next((p["texto"] for p in todo if p["tipo"] == "comunicado"), None)
    if not com:  # comunicado sin <blockquote>: párrafos entre "approval of the statement" y "Voting for"
        i0 = next((i for i, p in enumerate(acc) if "approval of the statement" in p["texto"]), None)
        if i0 is not None:
            partes = []
            for p in acc[i0 + 1:]:
                if p["texto"].startswith("Voting for"):
                    break
                partes.append(p["texto"])
            com = " ".join(partes) or None
    if com:
        com = re.sub(r'^"?The Federal Open Market Committee approved the following statement.*?vote:\s*', "", com)
        com = com.strip('" ')
    d["comunicado"] = com

    # Menciones de los "motores" de la narrativa, por cada 10 mil palabras
    texto_total = " ".join(p["texto"] for p in todo)
    n_pal = max(len(texto_total.split()), 1)
    for etiqueta, patron in MOTORES.items():
        d[f"motor_{etiqueta}"] = round(len(re.findall(patron, texto_total)) / n_pal * 10_000, 1)
    return d


def escala_riesgo(frase):
    """Convierte la redacción del staff en escala ordinal: + al alza, − a la baja.
    ±0.5 'somewhat/slightly', ±1 'skewed/tilted', ±1.5 'more skewed', 0 'balanced'."""
    if not frase:
        return None
    f = frase.lower()
    if "balanced" in f and not re.search(r"skewed|tilted", f):
        return 0.0
    signo = 1 if "upside" in f else -1 if "downside" in f else 0
    if re.search(r"more skewed|more tilted|significantly|considerably", f):
        mag = 1.5
    elif re.search(r"somewhat|slightly|a little", f):
        mag = 0.5
    else:
        mag = 1.0
    return signo * mag


def agregar_decision_tomada(dec: pd.DataFrame) -> pd.DataFrame:
    dec = dec.sort_values("fecha").reset_index(drop=True)
    cambio = dec["rango_sup"].diff()
    dec["decision"] = ["—" if pd.isna(c) else "alza" if c > 0 else "recorte" if c < 0 else "pausa" for c in cambio]
    dec["cambio_pb"] = (cambio * 100).round(0)
    return dec


# ----------------------------------------------------------------------------
# 6. Salidas: diff de comunicados, gráficas, Excel
# ----------------------------------------------------------------------------
def diff_comunicados(dec: pd.DataFrame, ruta: Path):
    filas = dec.sort_values("fecha").to_dict("records")
    bloques = []
    for prev, cur in zip(filas, filas[1:]):
        txt = lambda v: v if isinstance(v, str) else ""
        a, b = txt(prev["comunicado"]).split(), txt(cur["comunicado"]).split()
        out = []
        for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b).get_opcodes():
            if op == "equal":
                out.append(html_lib.escape(" ".join(a[i1:i2])))
            if op in ("delete", "replace"):
                out.append(f"<del>{html_lib.escape(' '.join(a[i1:i2]))}</del>")
            if op in ("insert", "replace"):
                out.append(f"<ins>{html_lib.escape(' '.join(b[j1:j2]))}</ins>")
        bloques.append(f"<h2>{prev['fecha']} &rarr; {cur['fecha']}</h2><p>{' '.join(out)}</p>")
    ruta.write_text(
        "<html><head><meta charset='utf-8'><title>Cambios en comunicados FOMC</title><style>"
        "body{font-family:Georgia,serif;max-width:900px;margin:2rem auto;line-height:1.6;padding:0 1rem}"
        "del{background:#fdd;color:#900}ins{background:#dfd;color:#060;text-decoration:none}"
        "h2{font-family:Arial,sans-serif;font-size:1rem;border-bottom:1px solid #ccc}</style></head>"
        "<body><h1>Cambios palabra por palabra entre comunicados</h1>" + "".join(bloques) + "</body></html>",
        encoding="utf-8")


# ----------------------------------------------------------------------------
# Gráficas
# ----------------------------------------------------------------------------
# Hitos para anotar en la gráfica de tasa. Edítalos libremente: (fecha AAAAMMDD, texto)
HITOS = [
    ("20250917", "Reanuda recortes"),
    ("20260318", "Conflicto en\nMedio Oriente"),
    ("20260617", "Sale el sesgo\na recortes"),
]

# Frases del comunicado a rastrear: (etiqueta, regex)
FRASES_COMUNICADO = [
    ("Inflación 'somewhat elevated'", r"somewhat elevated"),
    ("Inflación 'elevated' (sin matiz)", r"[Ii]nflation[^.]{0,40}(?<!somewhat )elevated"),
    ("Atento a riesgos de ambos lados", r"both sides of its dual mandate"),
    ("Riesgos a la baja para el empleo", r"downside risks to employment"),
    ("Sesgo a recortes ('additional adjustments')", r"extent and timing of additional adjustments"),
    ("Incertidumbre elevada", r"[Uu]ncertainty[^.]{0,60}(?:elevated|high|increased)|high level of uncertainty"),
    ("Menciona Medio Oriente", r"Middle East"),
    ("Empleo: 'job gains ... low/slowed'", r"[Jj]ob gains have (?:slowed|remained low)"),
    ("'Will deliver price stability'", r"will deliver price stability"),
]

TEMAS_ETIQUETA = {
    "inflacion": "Inflación",
    "mercado_laboral": "Mercado laboral",
    "actividad": "Actividad",
    "politica_monetaria": "Política monetaria",
    "estabilidad_financiera": "Estab. financiera*",
}
MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
C_HALCON, C_HALCON_CLARO = "#b2182b", "#f4a582"
C_PALOMA, C_PALOMA_CLARO = "#2166ac", "#92c5de"
C_GRIS, C_TEXTO = "#8c8c8c", "#333333"
FUENTE = "Fuente: minutas del FOMC (Federal Reserve). Elaboración propia con análisis de texto."


def _etq(f: str) -> str:
    return f"{MESES[int(f[4:6]) - 1]}-{f[2:4]}"


def _estilo():
    import matplotlib.pyplot as plt
    plt.rcParams.update({
        "font.family": "DejaVu Sans", "font.size": 9, "axes.titlesize": 11, "axes.titleweight": "bold",
        "axes.titlelocation": "left", "axes.spines.top": False, "axes.spines.right": False,
        "axes.edgecolor": "#999999", "axes.labelcolor": C_TEXTO, "xtick.color": C_TEXTO,
        "ytick.color": C_TEXTO, "legend.frameon": False, "figure.dpi": 100,
    })


def _xt(ax, fechas):
    rot = 45 if len(fechas) > 10 else 0
    ax.set_xticks(range(len(fechas)), [_etq(f) for f in fechas], rotation=rot,
                  ha="right" if rot else "center", rotation_mode="anchor", fontsize=8)


def _pie(fig):
    fig.text(0.01, 0.005, FUENTE, fontsize=7, color=C_GRIS, ha="left", va="bottom")


def g_tasa(ax, dec: pd.DataFrame):
    dec = dec.sort_values("fecha").reset_index(drop=True)
    x = list(range(len(dec)))
    inf, sup = dec["rango_inf"].astype(float), dec["rango_sup"].astype(float)
    ax.fill_between(x, inf, sup, step="mid", color="#4d4d4d", alpha=0.25, lw=0)
    ax.step(x, (inf + sup) / 2, where="mid", color="#4d4d4d", lw=1.8, label="Rango objetivo (punto medio)")
    marcas = {"alza": ("^", C_HALCON, 1), "pausa": ("o", "#e08214", 1), "sesgo": ("D", "#e08214", 1),
              "balance": ("x", C_GRIS, 1), "recorte": ("v", C_PALOMA, -1)}
    etiquetas = {"alza": "Disenso: alza", "pausa": "Disenso: mantener", "sesgo": "Disenso: guía/sesgo",
                 "balance": "Disenso: balance", "recorte": "Disenso: recorte"}
    usados = set()
    for i, r in dec.iterrows():
        arriba = abajo = 0
        for tipo, (mk, col, lado) in marcas.items():
            for _ in range(int(r.get(f"disenso_{tipo}", 0) or 0)):
                if lado > 0:
                    y = r["rango_sup"] + 0.07 + 0.09 * arriba; arriba += 1
                else:
                    y = r["rango_inf"] - 0.07 - 0.09 * abajo; abajo += 1
                ax.scatter(i, y, marker=mk, color=col, s=34, zorder=3,
                           label=etiquetas[tipo] if tipo not in usados else None)
                usados.add(tipo)
        if r.get("decision") in ("alza", "recorte") and not pd.isna(r.get("cambio_pb")):
            ax.annotate(f"{int(r['cambio_pb']):+d} pb", (i, r["rango_sup"]), xytext=(0, 3),
                        textcoords="offset points", ha="center", fontsize=7, color=C_TEXTO)
    # Cambio de presidente
    pres = dec["presidente"].fillna("").tolist()
    for i in range(1, len(pres)):
        if pres[i] and pres[i - 1] and pres[i] != pres[i - 1]:
            ax.axvline(i - 0.5, color=C_GRIS, ls=":", lw=1)
            ax.text(i - 0.45, ax.get_ylim()[1], f"Preside {pres[i].split()[-1]}", fontsize=7,
                    color=C_GRIS, va="top")
    fechas = dec["fecha"].astype(str).tolist()
    ymin = inf.min() - 0.55
    for f, txt in HITOS:
        if f in fechas:
            i = fechas.index(f)
            ax.annotate(txt, (i, ymin + 0.05), fontsize=7, color=C_TEXTO, ha="center", va="bottom")
    ax.set_ylim(ymin, sup.max() + 0.45)
    _xt(ax, fechas)
    ax.set_ylabel("%")
    ax.set_title("Tasa de fondos federales y votos disidentes")
    ax.legend(fontsize=7, loc="upper center", bbox_to_anchor=(0.5, -0.16), ncol=3)


def g_balance(ax, pos: pd.DataFrame, fechas: list[str], dec: pd.DataFrame | None = None):
    piv = pos.pivot(index="fecha", columns="postura", values="proporcion_max").reindex(fechas).fillna(0)
    x = list(range(len(fechas)))
    g = lambda c: piv[c] if c in piv else pd.Series(0, index=piv.index)
    recorto = pd.Series(False, index=piv.index)
    if dec is not None:
        recorto = (dec.set_index(dec["fecha"].astype(str))["decision"].reindex(fechas) == "recorte").values
    mantener = g("mantener_en_esta_reunion").where(recorto, 0)
    ax.bar(x, g("alza_condicional"), color=C_HALCON_CLARO, width=0.7, label="Alza probable si la inflación persiste")
    ax.bar(x, g("alza_en_esta_reunion"), color=C_HALCON, width=0.4, label="Favorecen alza en la reunión")
    if mantener.max() > 0:
        ax.bar(x, mantener, color="#e08214", width=0.4, label="Preferían no recortar (reuniones con recorte)")
    ax.bar(x, -g("recorte_condicional"), color=C_PALOMA_CLARO, width=0.7, label="Recortes probables más adelante")
    ax.bar(x, -g("recorte_en_esta_reunion"), color=C_PALOMA, width=0.4, label="Favorecen recorte en la reunión")
    ax.axhline(0, color=C_TEXTO, lw=0.8)
    ax.set_ylim(-1.05, 1.05)
    ax.set_yticks([-1, -0.5, 0, 0.5, 1], ["100%", "50%", "0", "50%", "100%"])
    ax.text(-0.6, 1.0, "HALCONES ▲", fontsize=7, color=C_HALCON, va="top", weight="bold")
    ax.text(-0.6, -1.0, "PALOMAS ▼", fontsize=7, color=C_PALOMA, va="bottom", weight="bold")
    _xt(ax, fechas)
    ax.set_ylabel("Proporción aprox. de participantes")
    ax.set_title("Balance de posturas en el comité")
    ax.legend(fontsize=7, loc="upper center", bbox_to_anchor=(0.5, -0.16), ncol=2)


def g_heatmap(ax, panel: pd.DataFrame, fechas: list[str], nivel: str, temas: list[str], titulo: str):
    import numpy as np
    sub = panel[panel["nivel"] == nivel]
    val = sub.pivot(index="tema", columns="fecha", values="indice_suavizado").reindex(index=temas, columns=fechas)
    n = sub.pivot(index="tema", columns="fecha", values="n_clausulas").reindex(index=temas, columns=fechas).fillna(0)
    im = ax.imshow(val.values.astype(float), cmap="RdBu_r", vmin=-0.8, vmax=0.8, aspect="auto")
    for i in range(len(temas)):
        for j in range(len(fechas)):
            v, k = val.iat[i, j], n.iat[i, j]
            if pd.isna(v) or k < MIN_CLAUSULAS:
                ax.add_patch(__import__("matplotlib").patches.Rectangle((j - 0.5, i - 0.5), 1, 1, color="#eeeeee"))
                ax.text(j, i, "·" if k == 0 else f"n={int(k)}", ha="center", va="center", fontsize=6, color=C_GRIS)
            else:
                ax.text(j, i, f"{v:+.2f}", ha="center", va="center", fontsize=7,
                        color="white" if abs(v) > 0.45 else C_TEXTO)
    _xt(ax, fechas)
    ax.set_yticks(range(len(temas)), [TEMAS_ETIQUETA[t] for t in temas])
    ax.set_title(titulo)
    for s in ax.spines.values():
        s.set_visible(False)
    return im


def g_datos_vs_tono(axs, dec: pd.DataFrame, panel: pd.DataFrame, fechas: list[str]):
    x = list(range(len(fechas)))
    dec = dec.set_index(dec["fecha"].astype(str)).reindex(fechas)
    tono = lambda nivel, tema: (panel[(panel.nivel == nivel) & (panel.tema == tema)]
                                .set_index("fecha")["indice_suavizado"].reindex(fechas))

    def gemela(ax, serie, etiqueta, color):
        a2 = ax.twinx()
        a2.fill_between(x, 0, serie.values.astype(float), color=color, alpha=0.12, lw=0)
        a2.plot(x, serie.values, color=color, lw=1.4, ls="--", marker=".", label=etiqueta)
        a2.set_ylim(-1, 1)
        a2.axhline(0, color=color, lw=0.5, alpha=0.5)
        a2.set_ylabel(etiqueta, color=color, fontsize=8)
        a2.spines["right"].set_visible(True)
        a2.tick_params(axis="y", colors=color, labelsize=7)
        return a2

    ax = axs[0]
    gemela(ax, tono("participantes", "inflacion"), "Tono participantes (+ hawkish)", C_HALCON)
    ax.plot(x, dec["pce_total"].astype(float), marker="o", color="#4d4d4d", lw=1.8, label="PCE total")
    ax.plot(x, dec["pce_subyacente"].astype(float), marker="s", color="#e08214", lw=1.8, label="PCE subyacente")
    ax.axhline(2, color=C_GRIS, ls="--", lw=1)
    ax.text(len(x) - 0.5, 2.03, "Meta 2%", fontsize=7, color=C_GRIS, ha="right", va="bottom")
    ax.set_zorder(2); ax.patch.set_visible(False)
    ax.set_ylabel("% anual (último dato citado por el staff)")
    ax.set_title("Inflación: datos vs tono del comité")
    ax.legend(fontsize=7, loc="upper left")

    if len(axs) > 1:
        ax = axs[1]
        gemela(ax, tono("participantes", "mercado_laboral"), "Tono participantes (+ fuerte)", C_PALOMA)
        u = dec["desempleo"].astype(float).reset_index(drop=True).dropna()
        ax.plot(u.index, u.values, marker="o", color="#4d4d4d", lw=1.8, label="Tasa de desempleo (sin dato = omitido)")
        ax.set_zorder(2); ax.patch.set_visible(False)
        ax.set_ylabel("%")
        ax.set_title("Empleo: datos vs tono del comité")
        ax.legend(fontsize=7, loc="upper left")

    if len(axs) > 2:
        ax = axs[2]
        ax.plot(x, tono("participantes", "actividad"), marker="o", color="#1b7837", lw=1.8, label="Participantes")
        ax.plot(x, tono("staff", "actividad"), marker="s", color="#7fbf7b", lw=1.5, ls="--", label="Staff")
        ax.axhline(0, color=C_GRIS, lw=0.8)
        ax.set_ylim(-1, 1)
        ax.set_ylabel("Tono (+ fuerte / − débil)")
        ax.set_title("Actividad: tono de participantes y staff")
        ax.legend(fontsize=7, loc="lower left")
    for a in axs:
        _xt(a, fechas)


def g_riesgos(ax, dec: pd.DataFrame, fechas: list[str]):
    dec = dec.set_index(dec["fecha"].astype(str)).reindex(fechas)
    x = list(range(len(fechas)))
    ax.axhspan(0, 1.7, color=C_HALCON, alpha=0.05)
    ax.axhspan(-1.7, 0, color=C_PALOMA, alpha=0.05)
    ax.step(x, dec["escala_riesgo_inflacion"].astype(float), where="mid", color=C_HALCON, lw=2, marker="o",
            label="Inflación")
    ax.step(x, dec["escala_riesgo_empleo"].astype(float), where="mid", color=C_PALOMA, lw=2, marker="s",
            label="Empleo y PIB")
    ax.axhline(0, color=C_TEXTO, lw=0.8)
    ax.set_ylim(-1.7, 1.7)
    ax.set_yticks([-1.5, -1, -0.5, 0, 0.5, 1, 1.5],
                  ["Más a la baja", "A la baja", "Algo a la baja", "Balanceados",
                   "Algo al alza", "Al alza", "Más al alza"])
    _xt(ax, fechas)
    ax.set_title("Balance de riesgos según el staff de la Fed")
    ax.legend(fontsize=8, loc="upper left")


def g_motores(ax, dec: pd.DataFrame, fechas: list[str]):
    dec = dec.set_index(dec["fecha"].astype(str)).reindex(fechas)
    x = list(range(len(fechas)))
    colores = ["#8c510a", "#762a83", "#b2182b", "#35978f", "#8c8c8c", "#dfc27d"]
    for (etq, _), col in zip(MOTORES.items(), colores):
        serie = dec.get(f"motor_{etq}")
        if serie is None or serie.fillna(0).max() == 0:
            continue
        ax.plot(x, serie.astype(float), marker="o", ms=4, lw=1.8, color=col, label=etq)
    _xt(ax, fechas)
    ax.set_ylabel("Menciones por cada 10 mil palabras")
    ax.set_title("¿De qué habla la Fed? Motores de la narrativa")
    ax.legend(fontsize=8, loc="upper left", ncol=2)


def g_lenguaje(ax, dec: pd.DataFrame, fechas: list[str]):
    dec = dec.set_index(dec["fecha"].astype(str)).reindex(fechas)
    for i, (etq, patron) in enumerate(FRASES_COMUNICADO):
        for j, f in enumerate(fechas):
            com = dec.at[f, "comunicado"] if f in dec.index else None
            presente = bool(com) and isinstance(com, str) and re.search(patron, com)
            ax.scatter(j, i, s=110, marker="s", color="#4d4d4d" if presente else "#eeeeee",
                       edgecolor="#bbbbbb", lw=0.5)
    _xt(ax, fechas)
    ax.set_yticks(range(len(FRASES_COMUNICADO)), [e for e, _ in FRASES_COMUNICADO])
    ax.invert_yaxis()
    ax.set_xlim(-0.6, len(fechas) - 0.4)
    ax.set_title("Lenguaje del comunicado: qué frases aparecen en cada reunión (■ = presente)")
    for s in ax.spines.values():
        s.set_visible(False)
    ax.tick_params(length=0)


def graficas(panel: pd.DataFrame, pos: pd.DataFrame, dec: pd.DataFrame, carpeta: Path):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    _estilo()
    fechas = sorted(dec["fecha"].astype(str))
    panel = panel.assign(fecha=panel["fecha"].astype(str))
    pos = pos.assign(fecha=pos["fecha"].astype(str))
    temas_p = ["inflacion", "mercado_laboral", "actividad", "politica_monetaria", "estabilidad_financiera"]
    temas_s = ["inflacion", "mercado_laboral", "actividad"]
    ancho = max(10, 0.75 * len(fechas))

    def guardar(fig, nombre):
        _pie(fig)
        fig.tight_layout(rect=(0, 0.02, 1, 1))
        fig.savefig(carpeta / nombre, dpi=170)
        plt.close(fig)

    fig, ax = plt.subplots(figsize=(ancho, 4.8)); g_tasa(ax, dec); guardar(fig, "01_tasa_y_votos.png")
    fig, ax = plt.subplots(figsize=(ancho, 4.8)); g_balance(ax, pos, fechas, dec); guardar(fig, "02_balance_posturas.png")

    fig, axs = plt.subplots(2, 1, figsize=(ancho, 6.2), gridspec_kw={"height_ratios": [5, 3]})
    im = g_heatmap(axs[0], panel, fechas, "participantes", temas_p, "Tono de los participantes por tema")
    g_heatmap(axs[1], panel, fechas, "staff", temas_s, "Tono del staff por tema")
    cb = fig.colorbar(im, ax=axs, shrink=0.6, pad=0.01)
    cb.set_label("− dovish / débil     + hawkish / fuerte", fontsize=8)
    fig.text(0.01, 0.03, "* Estabilidad financiera: + calma, − preocupación. Celdas grises: menos de "
             f"{MIN_CLAUSULAS} cláusulas.", fontsize=7, color=C_GRIS)
    fig.tight_layout(rect=(0, 0.05, 1, 1)); _pie(fig)
    fig.savefig(carpeta / "03_mapa_calor_temas.png", dpi=170, bbox_inches="tight"); plt.close(fig)

    fig, axs = plt.subplots(3, 1, figsize=(ancho, 10), sharex=False)
    g_datos_vs_tono(axs, dec, panel, fechas); guardar(fig, "04_datos_vs_tono.png")

    fig, ax = plt.subplots(figsize=(ancho, 4.2)); g_riesgos(ax, dec, fechas); guardar(fig, "05_riesgos_staff.png")
    fig, ax = plt.subplots(figsize=(ancho, 4.5)); g_motores(ax, dec, fechas); guardar(fig, "06_motores_narrativa.png")
    fig, ax = plt.subplots(figsize=(ancho, 0.42 * len(FRASES_COMUNICADO) + 1.4))
    g_lenguaje(ax, dec, fechas); guardar(fig, "07_lenguaje_comunicado.png")

    fig = plt.figure(figsize=(16, 10.5))
    gs = fig.add_gridspec(2, 2, hspace=0.55, wspace=0.22)
    g_tasa(fig.add_subplot(gs[0, 0]), dec)
    g_balance(fig.add_subplot(gs[0, 1]), pos, fechas, dec)
    g_motores(fig.add_subplot(gs[1, 0]), dec, fechas)
    g_datos_vs_tono([fig.add_subplot(gs[1, 1])], dec, panel, fechas)
    fig.suptitle("Panorama de la Reserva Federal según las minutas del FOMC", fontsize=15,
                 fontweight="bold", x=0.01, ha="left")
    fig.text(0.01, 0.945, f"{_etq(fechas[0])} a {_etq(fechas[-1])}  ·  política monetaria, inflación, empleo y actividad",
             fontsize=10, color=C_GRIS)
    _pie(fig)
    fig.savefig(carpeta / "00_lamina_resumen.png", dpi=170, bbox_inches="tight")
    plt.close(fig)


# ----------------------------------------------------------------------------
# Lámina resumen interactiva (Plotly, un solo HTML que abre sin internet)
# ----------------------------------------------------------------------------
# Canales de impacto macroeconómico de la IA (gráfica 09). La lectura
# (0 = oportunidad, 1 = mixto/condicional, 2 = riesgo) es una clasificación
# analítica de la minuta indicada en CANALES_IA_FECHA; el patrón sirve para
# localizar en la minuta la oración que se muestra como resumen en el hover.
CANALES_IA_FECHA = "20261028"
CANALES_IA = [
    ("Productividad", 0, r"productiv\w*|potential output"),
    ("Mercado laboral", 1, r"labor market|employment|hiring|\bjobs?\b|\bworkers\b"),
    ("Crecimiento / inversión", 1, r"invest\w*|data center|\bgrowth\b|spending|\bGDP\b|manufacturing"),
    ("Inflación / costos", 2, r"inflation\w*|\bprices?\b|price pressures?|\bcosts?\b"),
    ("Finanzas / estabilidad", 2, r"financ\w*|\bdebt\b|valuations?|credit|leverag\w*|\breturns?\b|asset price"),
]
NIVELES_IA = {0: ("Oportunidad", "#1b7837"), 1: ("Mixto / condicional", "#e08214"), 2: ("Riesgo", C_HALCON)}
RE_IA = re.compile(r"\bAI\b|(?i:artificial intelligence)")


def _envolver(texto: str, ancho: int = 60) -> str:
    """Escapa el texto y lo parte con <br> cada ~ancho caracteres."""
    import textwrap
    texto = html_lib.escape(re.sub(r"\s+", " ", str(texto)).strip(), quote=False)
    return "<br>".join(textwrap.wrap(texto, width=ancho, break_long_words=False, break_on_hyphens=False))


def _hover(campos: list[tuple[str, str]]) -> str:
    """[(etiqueta, valor)] -> '<b>Etiqueta:</b> valor' con saltos cada ~60 caracteres."""
    lineas = []
    for etq, val in campos:
        envuelto = _envolver(f"{etq}: {val}")
        lineas.append(envuelto.replace(html_lib.escape(etq) + ":", f"<b>{html_lib.escape(etq)}:</b>", 1))
    return "<br>".join(lineas)


def _recortar(o: str, maximo: int = 380) -> str:
    return o if len(o) <= maximo else o[:maximo].rsplit(" ", 1)[0] + " …"


def _oraciones_relevantes(parrafos: list[dict], patron, max_oraciones: int = 2,
                          max_chars: int = 340, filtro=None) -> str | None:
    """Resumen extractivo: la(s) oración(es) de la minuta con más coincidencias del
    patrón (se prefieren participantes y staff). Se devuelven tal cual, en el orden
    de la minuta."""
    pat = re.compile(patron) if isinstance(patron, str) else patron
    bono = {"participantes": 0.5, "staff_perspectivas": 0.3, "staff_economia": 0.3,
            "staff_financiero": 0.3, "mercados_desk": 0.1}
    cand = []
    for k, p in enumerate(parrafos):
        if p.get("tipo") == "directiva":
            continue
        ors = partir_oraciones(str(p.get("texto") or ""))
        for j, o in enumerate(ors):
            if filtro and not filtro(o, ors[j - 1] if j else ""):
                continue
            n = len(pat.findall(o))
            if n:
                puntaje = n + bono.get(p.get("seccion"), 0)
                puntaje -= 0.4 if p.get("tipo") == "comunicado" else 0
                cand.append((puntaje, k, j, o))
    if not cand:
        return None
    cand.sort(key=lambda c: (-c[0], c[1], c[2]))
    elegidas = [cand[0]]
    for c in cand[1:max_oraciones]:
        if sum(len(e[3]) for e in elegidas) + len(c[3]) <= max_chars:
            elegidas.append(c)
    elegidas.sort(key=lambda c: (c[1], c[2]))
    return " ".join(_recortar(c[3]) for c in elegidas)


def _pb_propuestos(texto: str):
    m = re.search(r"by (\d+) basis points?", texto)
    if m:
        return int(m.group(1))
    m = re.search(r"by (1/4|1/2|3/4) (?:of a )?percentage point", texto)
    return {"1/4": 25, "1/2": 50, "3/4": 75}[m.group(1)] if m else None


def _clausula(frase: str, nom: str, nombres: list[str]) -> str:
    """En "A, who preferred X, and B, who preferred Y" devuelve la parte de B.
    Conserva los grupos ("A, B, and C, who preferred ...")."""
    if sum(n in frase for n in nombres) < 2 or nom not in frase:
        return frase
    pos = sorted((frase.find(n), n) for n in nombres if n in frase)
    k = next(i for i, (_, n) in enumerate(pos) if n == nom)
    ini = pos[k][0]
    while k > 0 and re.fullmatch(r"[,\s]*(?:and\s*)?", frase[pos[k - 1][0] + len(pos[k - 1][1]):pos[k][0]]):
        k -= 1
        ini = pos[k][0]
    m = re.search(r"\b(who|preferred|supported|would have)\b", frase[frase.find(nom):])
    fin = len(frase)
    if m:
        desde = frase.find(nom) + m.start()
        sig = [p for p, _ in pos if p > desde]
        fin = min(sig) if sig else fin
    return frase[ini:fin].strip(" ;,").removesuffix(" and").strip(" ;,.") + "."


def _disidentes_detalle(acc: list[dict], cambio_pb) -> list[dict]:
    """Por disidente: nombre, qué propuso (etiqueta + oración literal de la minuta) y
    clasificación del disenso respecto de la decisión del comité."""
    i = next((k for k, p in enumerate(acc) if str(p["texto"]).startswith("Voting against this action")), None)
    if i is None:
        return []
    texto = str(acc[i]["texto"]).split(":", 1)[1].strip()
    if re.match(r"none\b", texto, re.I):
        return []
    nombres = _nombres(texto)
    apellidos = [n.split()[-1] for n in nombres]
    for p in acc[i + 1:]:
        t = str(p["texto"])
        if t.startswith("Consistent with") or "voted as an alternate" in t:
            break
        if any(a in t for a in apellidos) and re.search(r"preferred|dissent|voted against|supported", t):
            texto += " " + t
    oraciones = partir_oraciones(texto)
    decision = 0 if cambio_pb is None or pd.isna(cambio_pb) else int(cambio_pb)
    salida = []
    for nom, ape, (_, tipo) in zip(nombres, apellidos, _disidentes(acc)):
        verbo = r"preferred|supported|favored|dissented|would have"
        cand = [o for o in oraciones if ape in o and re.search(verbo, o)]
        solo = [o for o in cand if not any(a in o for a in apellidos if a != ape)]
        if not cand:   # nombres agrupados: "A, B y C, who preferred ..."
            cand = [o for o in oraciones if re.search(verbo, o)]
        frase = (cand[0] if cand else texto)
        # Si hay una oración solo de este disidente con más detalle (monto/motivo), se usa
        if solo and (_pb_propuestos(solo[0]) or not _pb_propuestos(frase)):
            frase = solo[0]
        frase = _clausula(frase, nom, nombres)
        pb = _pb_propuestos(frase) or _pb_propuestos(" ".join(cand))
        seg = frase.lower()
        if tipo == "alza":
            prop = None if pb is None else pb
            etiqueta = f"Subir la tasa {pb} pb" if pb else "Subir la tasa"
        elif tipo == "recorte":
            prop = None if pb is None else -pb
            etiqueta = f"Recortar la tasa {pb} pb" if pb else "Recortar la tasa"
        elif tipo == "pausa":
            prop, etiqueta = 0, "Mantener la tasa sin cambio"
        elif tipo == "sesgo/guía":
            prop = 0
            etiqueta = ("Mantener la tasa, pero sin sesgo a recortes en el comunicado"
                        if "easing bias" in seg else "Cambiar la guía / el lenguaje del comunicado")
        elif tipo == "balance":
            prop = 0
            etiqueta = ("Mantener la tasa y el ritmo de reducción de los valores en balance"
                        if re.search(r"continu\w* the pace of decline", seg) else "Cambio en el manejo del balance")
        else:
            prop, etiqueta = None, "No identificado"

        # Clasificación respecto de lo que decidió el comité
        if tipo == "sesgo/guía":
            clase = "restrictivo" if re.search(r"easing bias|not support|remov", seg) else "otro"
            motivo = "pidió una guía menos inclinada a recortes que la aprobada"
        elif tipo == "balance":
            restr = re.search(r"continu\w* the pace of decline|faster|larger", seg)
            clase = "restrictivo" if restr else "otro"
            motivo = ("pidió seguir reduciendo el balance al ritmo previo (el comité lo moderó)"
                      if restr else "desacuerdo sobre el balance")
        elif prop is None:
            clase, motivo = "otro", "la minuta no permite comparar con la decisión"
        else:
            clase = "restrictivo" if prop > decision else "expansivo" if prop < decision else "otro"
            motivo = f"propuso {prop:+d} pb; el comité decidió {decision:+d} pb"
        salida.append(dict(nombre=nom, tipo=tipo, propuso=etiqueta, prop_pb=prop, clase=clase,
                           motivo=motivo, frase=_recortar(frase, 420)))
    return salida


def generar_panorama_interactivo(parrafos, dec: pd.DataFrame, pos: pd.DataFrame, carpeta: Path,
                                 nombre: str = "panorama_fed_minutas.html"):
    """Versión interactiva de 00_lamina_resumen.png: tasa y disidentes, balance de
    posturas, motores de la narrativa y canales de impacto de la IA."""
    try:
        import plotly.graph_objects as go
        from plotly.subplots import make_subplots
    except ImportError:
        print("[aviso] falta plotly (pip install plotly): no se generó la lámina interactiva")
        return
    parrafos = [dict(p, fecha=str(p["fecha"])) for p in (parrafos.to_dict("records")
                if isinstance(parrafos, pd.DataFrame) else parrafos)]
    por_fecha: dict[str, list[dict]] = {}
    for p in parrafos:
        por_fecha.setdefault(p["fecha"], []).append(p)
    dec = dec.assign(fecha=dec["fecha"].astype(str)).sort_values("fecha").reset_index(drop=True)
    fechas = dec["fecha"].tolist()
    x = list(range(len(fechas)))
    etiquetas_x = [_etq(f) for f in fechas]
    tit_canales = "Canales de impacto macroeconómico de la IA"
    if CANALES_IA_FECHA in fechas:
        tit_canales += f" · {_etq(CANALES_IA_FECHA)}"

    fig = make_subplots(
        rows=2, cols=2, vertical_spacing=0.27, horizontal_spacing=0.13,
        subplot_titles=("Tasa de fondos federales y votos disidentes", "Balance de posturas en el comité",
                        "¿De qué habla la Fed? Motores de la narrativa", tit_canales))

    # --- 1. Tasa y disidentes --------------------------------------------------
    inf, sup = dec["rango_inf"].astype(float), dec["rango_sup"].astype(float)
    fig.add_trace(go.Scatter(x=x, y=inf, mode="lines", line=dict(width=0, shape="hvh"),
                             hoverinfo="skip", showlegend=False), 1, 1)
    fig.add_trace(go.Scatter(x=x, y=sup, mode="lines", line=dict(width=0, shape="hvh"), fill="tonexty",
                             fillcolor="rgba(77,77,77,0.25)", hoverinfo="skip", showlegend=False), 1, 1)
    hov_tasa = []
    for _, r in dec.iterrows():
        cambio = "" if pd.isna(r["cambio_pb"]) else f" ({int(r['cambio_pb']):+d} pb)"
        hov_tasa.append(_hover([("Reunión", _etq(r["fecha"])),
                                ("Rango objetivo", f"{r['rango_inf']:.2f}–{r['rango_sup']:.2f}%"),
                                ("Decisión", f"{r['decision']}{cambio}")]))
    fig.add_trace(go.Scatter(x=x, y=(inf + sup) / 2, mode="lines", line=dict(color="#4d4d4d", width=2.2, shape="hvh"),
                             name="Rango objetivo (punto medio)", legend="legend", legendrank=1, customdata=hov_tasa,
                             hovertemplate="%{customdata}<extra></extra>"), 1, 1)

    simbolos = {"alza": "triangle-up", "recorte": "triangle-down", "pausa": "circle",
                "sesgo/guía": "diamond", "balance": "x", "?": "square"}
    nombre_tipo = {"alza": "alza", "recorte": "recorte", "pausa": "mantener", "sesgo/guía": "guía/sesgo",
                   "balance": "balance", "?": "no identificado"}
    color_clase = {"restrictivo": C_HALCON, "expansivo": C_PALOMA, "otro": C_GRIS}
    nombre_clase = {"restrictivo": "Disenso restrictivo", "expansivo": "Disenso expansivo", "otro": "Disenso"}
    puntos: dict[tuple, dict] = {}
    for i, r in dec.iterrows():
        acc = [p for p in por_fecha.get(r["fecha"], []) if p["seccion"] == "acciones_politica"]
        arriba = 1.2 if r["decision"] in ("alza", "recorte") else 0   # deja espacio a la etiqueta "±xx pb"
        abajo = 0
        cambio = "sin cambio (primera reunión)" if pd.isna(r["cambio_pb"]) else \
            "sin cambio" if r["cambio_pb"] == 0 else f"{int(r['cambio_pb']):+d} pb"
        for d in _disidentes_detalle(acc, r["cambio_pb"]):
            if d["clase"] == "expansivo":
                y = r["rango_inf"] - 0.07 - 0.09 * abajo; abajo += 1
            else:
                y = r["rango_sup"] + 0.07 + 0.09 * arriba; arriba += 1
            clave = (d["clase"], d["tipo"])
            pt = puntos.setdefault(clave, dict(x=[], y=[], h=[]))
            pt["x"].append(i); pt["y"].append(y)
            pt["h"].append(_hover([
                ("Disidente", d["nombre"]), ("Reunión", _etq(r["fecha"])),
                ("Qué propuso", d["propuso"]), ("Decisión del comité", cambio),
                ("Clasificación", f"{nombre_clase[d['clase']].lower()} — {d['motivo']}"),
                ("Minuta", f"“{d['frase']}”")]))
    orden = {"restrictivo": 0, "expansivo": 1, "otro": 2}
    for (clase, tipo), pt in sorted(puntos.items(), key=lambda kv: (orden[kv[0][0]], kv[0][1])):
        fig.add_trace(go.Scatter(
            x=pt["x"], y=pt["y"], mode="markers", legend="legend",
            name=f"{nombre_clase[clase]}: " + ("recorte menor que el decidido"
                                               if (clase, tipo) == ("restrictivo", "recorte") else nombre_tipo[tipo]),
            marker=dict(symbol=simbolos.get(tipo, "square"), size=10, color=color_clase[clase],
                        line=dict(width=1.5, color=color_clase[clase]) if tipo == "balance" else dict(width=0)),
            customdata=pt["h"], hovertemplate="%{customdata}<extra></extra>"), 1, 1)

    for i, r in dec.iterrows():
        if r["decision"] in ("alza", "recorte") and not pd.isna(r["cambio_pb"]):
            fig.add_annotation(x=i, y=r["rango_sup"], text=f"{int(r['cambio_pb']):+d} pb", showarrow=False,
                               yshift=9, font=dict(size=9, color=C_TEXTO), row=1, col=1)
    ymin, ymax = inf.min() - 0.55, sup.max() + 0.45
    pres = dec["presidente"].fillna("").tolist()
    for i in range(1, len(pres)):
        if pres[i] and pres[i - 1] and pres[i] != pres[i - 1]:
            fig.add_shape(type="line", x0=i - 0.5, x1=i - 0.5, y0=ymin, y1=ymax, row=1, col=1,
                          line=dict(color=C_GRIS, dash="dot", width=1))
            fig.add_annotation(x=i - 0.45, y=ymax, text=f"Preside {pres[i].split()[-1]}", showarrow=False,
                               xanchor="left", yanchor="top", font=dict(size=9, color=C_GRIS), row=1, col=1)
    for f, txt in HITOS:
        if f in fechas:
            fig.add_annotation(x=fechas.index(f), y=ymin + 0.05, text=txt.replace("\n", "<br>"), showarrow=False,
                               yanchor="bottom", font=dict(size=9, color=C_TEXTO), row=1, col=1)
    fig.update_yaxes(title_text="%", range=[ymin, ymax], row=1, col=1)

    # --- 2. Balance de posturas (igual que la versión estática) ---------------
    pos = pos.assign(fecha=pos["fecha"].astype(str))
    piv = pos.pivot(index="fecha", columns="postura", values="proporcion_max").reindex(fechas).fillna(0)
    cuan = pos.pivot(index="fecha", columns="postura", values="cuantificador").reindex(fechas)
    g = lambda c: piv[c] if c in piv else pd.Series(0.0, index=piv.index)
    recorto = (dec.set_index("fecha")["decision"].reindex(fechas) == "recorte").values
    barras = [("alza_condicional", 1, C_HALCON_CLARO, 0.7, "Alza probable si la inflación persiste", None),
              ("alza_en_esta_reunion", 1, C_HALCON, 0.4, "Favorecen alza en la reunión", None),
              ("mantener_en_esta_reunion", 1, "#e08214", 0.4, "Preferían no recortar (reuniones con recorte)", recorto),
              ("recorte_condicional", -1, C_PALOMA_CLARO, 0.7, "Recortes probables más adelante", None),
              ("recorte_en_esta_reunion", -1, C_PALOMA, 0.4, "Favorecen recorte en la reunión", None)]
    for col, signo, color, ancho, etiqueta, mascara in barras:
        v = g(col).astype(float)
        if mascara is not None:
            v = v.where(mascara, 0)
            if v.max() == 0:
                continue
        q = cuan[col] if col in cuan else pd.Series(None, index=cuan.index)
        hov = [_hover([("Reunión", _etq(f)), ("Postura", etiqueta), ("Proporción aprox.", f"{val:.0%}"),
                       ("Cuantificador", qq if isinstance(qq, str) else "—")])
               for f, val, qq in zip(fechas, v, q)]
        fig.add_trace(go.Bar(x=x, y=signo * v, width=ancho, marker_color=color, name=etiqueta, legend="legend2",
                             customdata=hov, hovertemplate="%{customdata}<extra></extra>"), 1, 2)
    fig.add_hline(y=0, line=dict(color=C_TEXTO, width=0.8), row=1, col=2)
    fig.update_yaxes(title_text="Proporción aprox. de participantes", range=[-1.05, 1.05],
                     tickvals=[-1, -0.5, 0, 0.5, 1], ticktext=["100%", "50%", "0", "50%", "100%"], row=1, col=2)
    fig.add_annotation(x=-0.6, y=1.0, text="<b>HALCONES ▲</b>", showarrow=False, xanchor="left", yanchor="top",
                       font=dict(size=9, color=C_HALCON), row=1, col=2)
    fig.add_annotation(x=-0.6, y=-1.0, text="<b>PALOMAS ▼</b>", showarrow=False, xanchor="left",
                       yanchor="bottom", font=dict(size=9, color=C_PALOMA), row=1, col=2)

    # --- 3. Motores de la narrativa -------------------------------------------
    colores = ["#8c510a", "#762a83", "#b2182b", "#35978f", "#8c8c8c", "#dfc27d"]
    for (tema, patron), color in zip(MOTORES.items(), colores):
        serie = dec.get(f"motor_{tema}")
        if serie is None or serie.fillna(0).max() == 0:
            continue
        pat = re.compile(patron)
        hov = []
        for f, val in zip(fechas, serie.astype(float)):
            ps = por_fecha.get(f, [])
            n = len(pat.findall(" ".join(str(p["texto"]) for p in ps)))
            resumen = _oraciones_relevantes(ps, pat) if n else None
            hov.append(_hover([("Tema", tema), ("Reunión", _etq(f)),
                               ("Menciones", f"{n} ({val:.1f} por cada 10 mil palabras)"),
                               ("Resumen", f"“{resumen}”" if resumen else "Sin menciones")]))
        fig.add_trace(go.Scatter(x=x, y=serie.astype(float), mode="lines+markers", name=tema, legend="legend3",
                                 line=dict(color=color, width=2), marker=dict(size=6, color=color),
                                 customdata=hov, hovertemplate="%{customdata}<extra></extra>"), 2, 1)
    fig.update_yaxes(title_text="Menciones por cada 10 mil palabras", rangemode="tozero", row=2, col=1)

    # --- 4. Canales de impacto macroeconómico (IA) ----------------------------
    ps_ia = [p for p in por_fecha.get(CANALES_IA_FECHA, []) if RE_IA.search(str(p["texto"]))]
    nombres = [c[0] for c in CANALES_IA]
    for canal, nivel, patron in CANALES_IA:
        fig.add_trace(go.Scatter(x=[0, 2], y=[canal, canal], mode="lines", hoverinfo="skip", showlegend=False,
                                 line=dict(color="#e6e6e6", width=4)), 2, 2)
    puntos_ia = {}
    for canal, nivel, patron in CANALES_IA:
        pat = re.compile(patron, re.I)
        n = sum(1 for p in ps_ia if pat.search(str(p["texto"])))
        # Solo oraciones que mencionan IA o que continúan ("These participants ...") una que la menciona
        resumen = _oraciones_relevantes(ps_ia, pat, filtro=lambda o, previa: bool(
            RE_IA.search(o) or (RE_IA.search(previa) and re.match(r"(These|They|This|Those)\b", o)))) if n else None
        etiqueta, _ = NIVELES_IA[nivel]
        pt = puntos_ia.setdefault(nivel, dict(x=[], y=[], h=[]))
        pt["x"].append(nivel); pt["y"].append(canal)
        pt["h"].append(_hover([("Canal", canal), ("Reunión", _etq(CANALES_IA_FECHA)), ("Lectura", etiqueta),
                               ("Párrafos con IA y el canal", str(n)),
                               ("Resumen", f"“{resumen}”" if resumen else "Sin menciones")]))
    for nivel, pt in sorted(puntos_ia.items()):
        etiqueta, color = NIVELES_IA[nivel]
        fig.add_trace(go.Scatter(x=pt["x"], y=pt["y"], mode="markers", name=etiqueta, showlegend=False,
                                 marker=dict(size=14, color=color), customdata=pt["h"],
                                 hovertemplate="%{customdata}<extra></extra>"), 2, 2)
    fig.update_xaxes(range=[-0.45, 2.45], tickvals=[0, 1, 2], showgrid=True, gridcolor="#e6e6e6",
                     ticktext=[NIVELES_IA[k][0] for k in (0, 1, 2)], zeroline=False, row=2, col=2)
    fig.update_yaxes(categoryorder="array", categoryarray=nombres[::-1], showgrid=False, row=2, col=2)

    # --- Ejes, leyendas y formato ---------------------------------------------
    for rr, cc in ((1, 1), (1, 2), (2, 1)):
        fig.update_xaxes(tickmode="array", tickvals=x, ticktext=etiquetas_x, tickangle=-45 if len(x) > 10 else 0,
                         range=[-0.7, len(x) - 0.3], showgrid=False, row=rr, col=cc)
    fig.update_xaxes(showline=True, linecolor="#999999", ticks="outside", tickfont=dict(size=10))
    fig.update_yaxes(showline=True, linecolor="#999999", ticks="outside", zeroline=False, tickfont=dict(size=10))
    fig.update_yaxes(showgrid=False, row=1, col=1)
    fig.update_yaxes(showgrid=False, row=1, col=2)
    fig.update_yaxes(showgrid=False, row=2, col=1)

    leyenda = dict(orientation="v", xanchor="left", yanchor="top", font=dict(size=10),
                   bgcolor="rgba(255,255,255,0.75)", tracegroupgap=2, traceorder="normal")
    d1, d2 = fig.layout.xaxis.domain, fig.layout.xaxis2.domain
    y_ley = fig.layout.yaxis.domain[0] - 0.065
    fig.update_layout(
        title=dict(text="<b>Panorama de la Reserva Federal según las minutas del FOMC</b><br>"
                        f"<span style='font-size:13px;color:{C_GRIS}'>{_etq(fechas[0])} a {_etq(fechas[-1])}  ·  "
                        "política monetaria, inflación, empleo y actividad</span>",
                   x=0.01, xanchor="left", font=dict(size=20, color="black")),
        template="simple_white", height=1300, barmode="overlay", bargap=0,
        font=dict(family="DejaVu Sans, Verdana, Arial, sans-serif", size=11, color=C_TEXTO),
        hoverlabel=dict(bgcolor="white", font=dict(size=11, color=C_TEXTO), bordercolor="#999999", align="left"),
        margin=dict(l=80, r=40, t=110, b=110),
        legend=dict(leyenda, x=d1[0], y=y_ley),
        legend2=dict(leyenda, x=d2[0], y=y_ley),
        legend3=dict(leyenda, x=fig.layout.xaxis3.domain[0] + 0.01, y=fig.layout.yaxis3.domain[1]),
    )
    for a in fig.layout.annotations[:4]:   # títulos de los paneles: negrita, a la izquierda
        dom = {"Tasa": d1, "Balance": d2}.get(a.text.split()[0], None)
        a.update(text=f"<b>{a.text}</b>", font=dict(size=14, color="black"), xanchor="left",
                 x=(dom or (fig.layout.xaxis3.domain if a.text.startswith("¿") else fig.layout.xaxis4.domain))[0])
    fig.add_annotation(text="Clasificación analítica del contenido; no es una previsión cuantitativa. "
                            "Pasa el cursor sobre cada punto para ver la oración de la minuta.",
                       xref="paper", yref="paper", x=fig.layout.xaxis4.domain[0], y=fig.layout.yaxis4.domain[0] - 0.045,
                       xanchor="left", yanchor="top", showarrow=False, font=dict(size=9, color=C_GRIS))
    fig.add_annotation(text=FUENTE, xref="paper", yref="paper", x=0, y=-0.07, xanchor="left", yanchor="top",
                       showarrow=False, font=dict(size=9, color=C_GRIS))
    ruta = carpeta / nombre
    fig.write_html(ruta, include_plotlyjs=True, full_html=True,
                   config={"displaylogo": False, "locale": "es", "responsive": True})
    print(f"[ok] lámina interactiva: {ruta.resolve()}")


def exportar(parrafos, df, panel, pos, dec, carpeta: Path):
    carpeta.mkdir(parents=True, exist_ok=True)
    dec_x = dec.drop(columns=["comunicado"])
    with pd.ExcelWriter(carpeta / "comparativo_minutas_fomc.xlsx", engine="openpyxl") as xw:
        dec_x.to_excel(xw, sheet_name="decisiones", index=False)
        panel.pivot_table(index=["nivel", "tema"], columns="fecha", values="indice_suavizado").to_excel(xw, sheet_name="indice_por_tema")
        panel.to_excel(xw, sheet_name="panel_detalle", index=False)
        pos.to_excel(xw, sheet_name="posturas", index=False)
        dec[["fecha", "comunicado"]].to_excel(xw, sheet_name="comunicados", index=False)
        df.to_excel(xw, sheet_name="oraciones", index=False)
        pd.DataFrame(parrafos).to_excel(xw, sheet_name="parrafos", index=False)
    df.to_csv(carpeta / "oraciones.csv", index=False)
    panel.to_csv(carpeta / "panel_temas.csv", index=False)
    dec.to_csv(carpeta / "decisiones.csv", index=False)
    diff_comunicados(dec, carpeta / "comunicados_diff.html")
    graficas(panel, pos, dec, carpeta)
    generar_panorama_interactivo(parrafos, dec, pos, carpeta)


# ----------------------------------------------------------------------------
# 7. Capa opcional con LLM (Claude) — más fina que el diccionario
# ----------------------------------------------------------------------------
RUBRICA = """Eres analista de bancos centrales. Recibirás un párrafo de una minuta del FOMC.
Para CADA oración devuelve un objeto JSON con:
  "oracion": índice (0,1,2...),
  "tema": uno de ["inflacion","mercado_laboral","actividad","estabilidad_financiera","politica_monetaria","otro"],
  "direccion": "hawkish" | "dovish" | "neutral"  (desde la óptica de presión sobre la tasa de fondos federales),
  "intensidad": 0, 1 o 2,
  "cuantificador": el cuantificador literal si existe ("a few participants", "most participants"...) o null,
  "horizonte": "actual" | "perspectiva" | "riesgo".
Responde SOLO con un arreglo JSON, sin texto adicional ni backticks."""


def clasificar_llm(parrafos: list[dict], modelo: str) -> pd.DataFrame:
    """Requiere: pip install anthropic  y  la variable ANTHROPIC_API_KEY."""
    import anthropic
    cliente = anthropic.Anthropic()
    filas = []
    for p in parrafos:
        if p["seccion"] not in ("participantes", "staff_perspectivas", "acciones_politica") or p["tipo"] != "parrafo":
            continue
        oraciones = partir_oraciones(p["texto"])
        entrada = "\n".join(f"[{i}] {o}" for i, o in enumerate(oraciones))
        try:
            resp = cliente.messages.create(model=modelo, max_tokens=2000, system=RUBRICA,
                                           messages=[{"role": "user", "content": entrada}])
            texto = "".join(b.text for b in resp.content if b.type == "text")
            datos = json.loads(re.sub(r"```(?:json)?", "", texto).strip())
        except Exception as e:  # noqa: BLE001
            print(f"[aviso] LLM falló en {p['fecha']} párrafo {p['n']}: {e}")
            continue
        for item in datos:
            i = item.get("oracion", 0)
            filas.append(dict(fecha=p["fecha"], seccion=p["seccion"], parrafo=p["n"], oracion_n=i,
                              oracion=oraciones[i] if i < len(oraciones) else None, **{
                                  k: item.get(k) for k in ("tema", "direccion", "intensidad", "cuantificador", "horizonte")}))
    return pd.DataFrame(filas)


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def diagnostico(parrafos: list[dict]):
    conteo = pd.DataFrame(parrafos).groupby(["fecha", "seccion"]).size().unstack(fill_value=0)
    print("\nPárrafos por sección (si 'participantes' sale en 0, cambió el HTML):")
    print(conteo.to_string())
    faltan = [f for f in conteo.index if conteo.loc[f].get("participantes", 0) == 0]
    if faltan:
        print(f"[ALERTA] sin sección de participantes en: {faltan}")


def main():
    ap = argparse.ArgumentParser(description="Comparativo y sentiment de minutas del FOMC")
    ap.add_argument("--anio", type=int, nargs="+", default=[2025, 2026],
                    help="Uno o varios años, p. ej. --anio 2025 2026")
    ap.add_argument("--fechas", nargs="*", help="Fechas AAAAMMDD (segundo día de la reunión)")
    ap.add_argument("--html-dir", default="minutas_html")
    ap.add_argument("--salida", default="salida_fomc")
    ap.add_argument("--solo-local", action="store_true", help="No descargar; usar lo que haya en --html-dir")
    ap.add_argument("--desde-excel", help="Regenerar todo desde la hoja 'parrafos' de un Excel previo (sin descargar)")
    ap.add_argument("--llm", action="store_true", help="Agregar clasificación con Claude (API)")
    ap.add_argument("--modelo", default=os.environ.get("FOMC_LLM_MODEL", "claude-sonnet-4-6"))
    args = ap.parse_args()

    html_dir, salida = Path(args.html_dir), Path(args.salida)
    if args.desde_excel:
        pf = pd.read_excel(args.desde_excel, sheet_name="parrafos", dtype={"fecha": str})
        parrafos = pf.to_dict("records")
        rutas = {f: None for f in sorted(pf["fecha"].unique())}
    elif args.solo_local:
        rutas = {re.search(r"(\d{8})", p.name).group(1): p for p in sorted(html_dir.glob("*.htm*"))}
    else:
        fechas = args.fechas or sorted(f for a in args.anio for f in descubrir_fechas(a))
        print(f"Fechas: {fechas}")
        rutas = descargar(fechas, html_dir)
    if not rutas:
        sys.exit("No hay minutas para procesar.")

    if not args.desde_excel:
        parrafos = []
        for fecha, ruta in sorted(rutas.items()):
            parrafos += parsear_html(ruta.read_text(encoding="utf-8", errors="ignore"), fecha)
    diagnostico(parrafos)

    df = analizar(parrafos)
    panel = panel_por_tema(df)
    pos = posturas(df)
    dec = agregar_decision_tomada(pd.DataFrame([extraer_decision(parrafos, f) for f in sorted(rutas)]))
    exportar(parrafos, df, panel, pos, dec, salida)

    if args.llm:
        llm = clasificar_llm(parrafos, args.modelo)
        llm.to_csv(salida / "clasificacion_llm.csv", index=False)
        print(f"[ok] clasificación LLM: {len(llm)} oraciones")

    print("\nÍndice neto por tema (participantes):")
    print(panel[panel.nivel == "participantes"].pivot(index="tema", columns="fecha", values="indice_suavizado").to_string())
    print("\nDecisiones:")
    print(dec[["fecha", "presidente", "rango_inf", "rango_sup", "decision", "votos_a_favor", "votos_en_contra", "disidentes"]].to_string(index=False))
    print(f"\nListo. Resultados en: {salida.resolve()}")


if __name__ == "__main__":
    main()
