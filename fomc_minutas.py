#!/usr/bin/env python3
"""
fomc_minutas.py
Comparativo de minutas del FOMC: evolución de la percepción de inflación,
mercado laboral, actividad, estabilidad financiera y política monetaria.

Flujo:
  1. Descarga (o lee de disco) las minutas en HTML desde federalreserve.gov.
  2. Parsea cada minuta en secciones y párrafos (la estructura de la Fed es fija).
  3. Parte en oraciones y cláusulas; asigna tema(s), cuantificador ("a few",
     "most", ...) y dirección hawkish/dovish con un diccionario por tema.
  4. Extrae variables duras: rango objetivo, votos, disidentes y su dirección,
     PCE / subyacente / desempleo citados por el staff, balance de riesgos.
  5. Genera: Excel con varias hojas, CSVs, gráficas PNG y un HTML con el diff
     palabra por palabra entre comunicados consecutivos.

Uso:
  pip install requests beautifulsoup4 pandas matplotlib openpyxl
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
    ("staff_perspectivas", r"staff economic outlook"),
    ("participantes", r"participants.? views on current conditions and the economic outlook"),
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
# "det" = these/those marca un cuantificador anidado ("a few of these participants"):
# se decide con el propio match, no buscando "of these" en toda la oración.
RE_CUANT = re.compile(
    r"\b(?P<q>" + "|".join(re.escape(q) for q in _q_alts) + r")\s+"
    r"(?:of\s+)?(?P<det>these\s+|those\s+|the\s+)?(?:other\s+)?"
    r"(?P<actor>participants?|members?)\b",
    re.I,
)
RE_GENERALLY = re.compile(r"\b(participants|members)\s+generally\b", re.I)
RE_BARE = re.compile(r"^(participants|members)\b", re.I)          # "Participants noted..."
RE_ELIPTICO = re.compile(
    r"^(?:but|while|whereas|and|however,)?\s*(?P<q>a few|a couple|several|many|most|some)\b"
    r"(?!\s+(?:of|participants|members|other participants))\s+(?:other\s+)?[a-z]+(?:ed|s)\b", re.I)
RE_HEREDA = re.compile(r"^(these|they)\b|\bthese participants\b", re.I)
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
             r"take longer to return", r"delay(?:ed)? the (?:anticipated|expected) (?:decline|return)"],
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


# Error impedía uso completo de 2024 porque cambió el preambulo
_SECCIONES_ANALISIS = {"participantes", "acciones_politica"}
_SIN_TEXTO = ("preambulo", "asistencia")


def fin_de_minuta(seccion: str, vistas: set) -> bool:
    return seccion == "asistencia" and bool(vistas & _SECCIONES_ANALISIS)


def tipo_blockquote(txt: str, seccion: str) -> str:
    """La directiva cambió de redacción ("directs the Desk" desde 2024-07; antes
    "Undertake open market operations ..."). Solo es comunicado lo que está en
    "Committee Policy Actions" y no es la directiva."""
    if "directs the Desk" in txt or re.match(r'"?Undertake open market operations', txt):
        return "directiva"
    return "comunicado" if seccion == "acciones_politica" else "cita"


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

    parrafos, seccion, n, vistas = [], "preambulo", 0, set()
    for el in cont.find_all(["p", "li", "h2", "h3", "h4", "h5", "h6", "blockquote"]):
        if el.name in ("p", "li") and el.find_parent("blockquote"):
            continue  # el blockquote se procesa completo

        if el.name.startswith("h"):
            clave = identificar_seccion(el.get_text(" "))
            if clave:
                if fin_de_minuta(clave, vistas):
                    break
                seccion = clave
                vistas.add(clave)
            continue

        if el.name == "blockquote":
            if seccion in _SIN_TEXTO:
                continue
            txt = normalizar(el.get_text(" "))
            n += 1
            parrafos.append(dict(fecha=fecha, seccion=seccion, n=n, tipo=tipo_blockquote(txt, seccion), texto=txt))
            continue

        # <p>: ¿empieza con <strong>Título de sección</strong>?
        hijos = [c for c in el.children if not (isinstance(c, NavigableString) and not c.strip())]
        if hijos and getattr(hijos[0], "name", None) in ("strong", "b"):
            clave = identificar_seccion(hijos[0].get_text(" "))
            if clave:
                if fin_de_minuta(clave, vistas):
                    break
                seccion = clave
                vistas.add(clave)
                hijos[0].extract()
        txt = normalizar(el.get_text(" "))
        if not txt or seccion in _SIN_TEXTO:
            continue
        n += 1
        parrafos.append(dict(fecha=fecha, seccion=seccion, n=n, tipo="parrafo", texto=txt))
    return parrafos


def parsear_texto(texto: str, fecha: str) -> list[dict]:
    """Respaldo para texto extraído de PDF (pdftotext). Menos preciso que el HTML:
    detecta fin de párrafo cuando una línea termina en punto y es más corta."""
    lineas = [l.rstrip() for l in texto.replace("\f", "\n").split("\n")]
    ancho = max((len(l) for l in lineas), default=100)
    parrafos, seccion, n, buffer, vistas = [], "preambulo", 0, [], set()

    def volcar():
        nonlocal n, buffer
        if buffer and seccion not in _SIN_TEXTO:
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
            if fin_de_minuta(clave, vistas):
                break
            seccion = clave
            vistas.add(clave)
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
    """Devuelve (etiqueta, proporción, anidado) del primer cuantificador del texto,
    o (None, None, False) si no hay. 'anidado' = "... of these/those participants"."""
    m = RE_CUANT.search(texto)
    if m:
        q, actor = m.group("q").lower(), m.group("actor").lower()
        anidado = (m.group("det") or "").strip().lower() in ("these", "those")
        if q in NUMEROS:
            total = N_VOTANTES if actor.startswith("member") else N_PARTICIPANTES
            return f"{q} {actor}", round(min(NUMEROS[q] / total, 1), 2), anidado
        return f"{q} {actor}", CUANTIFICADORES[q], anidado
    m = RE_GENERALLY.search(texto)
    if m:
        return f"{m.group(1).lower()} generally", CUANTIFICADORES["generally"], False
    m = RE_BARE.search(texto)
    if m:
        return f"{m.group(1).lower()} (sin calificar)", 1.0, False
    m = RE_ELIPTICO.search(texto)
    if m:
        q = m.group("q").lower()
        q = "a couple of" if q == "a couple" else q
        return f"{q} (elíptico)", CUANTIFICADORES[q], False
    return None, None, False


def resolver_cuantificador(texto: str, padre):
    """Cuantificador de un texto. Si es anidado ("almost all of these participants"),
    se multiplica por la proporción del grupo padre (el último grupo no anidado)."""
    etq, prop, anidado = cuantificador(texto)
    if etq and anidado and padre[1]:
        return f"{etq} [de {padre[0]}]", round(prop * padre[1], 2), True
    return etq, prop, anidado


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
            padre = heredado  # "of these participants" en una cláusula: el último grupo
            for i_c, cl in enumerate(partir_clausulas(oracion)):  # no anidado de esta oración
                etq_c, prop_c, anidado_c = resolver_cuantificador(cl, padre)
                if etq_c is not None and not anidado_c:
                    padre = (etq_c, prop_c)
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
    d["pce_subyacente_ultimo_dato"], _ = _dato(eco.split("staff estimated")[0], r"(?i)core PCE price inflation.{0,160}?(\d+\.\d) percent")
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
    PAT_INFLACION, PAT_EMPLEO = r"inflation", r"employment|GDP|real activity|economic activity"
    d["staff_riesgo_inflacion"] = frase_riesgo(PAT_INFLACION)
    d["staff_riesgo_empleo"] = frase_riesgo(PAT_EMPLEO)
    d["escala_riesgo_inflacion"] = escala_riesgo(d["staff_riesgo_inflacion"], PAT_INFLACION)
    d["escala_riesgo_empleo"] = escala_riesgo(d["staff_riesgo_empleo"], PAT_EMPLEO)

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


RE_SESGO_RIESGO = re.compile(
    r"(?P<pre>(?:[\w-]+\s+){0,3})(?:skewed|tilted)\s+(?P<post>(?:[\w-]+\s+){0,2}?)"
    r"(?:to|toward)\s+the\s+(?P<dir>upside|downside)")


def escala_riesgo(frase, tema_pat: str | None = None):
    """Convierte la redacción del staff en escala ordinal: + al alza, − a la baja.
    0 'balanced', ±0.5 'somewhat/slightly skewed', ±1 'skewed/tilted',
    ±1.25 'a little more skewed', ±1.5 'more skewed'.
    El modificador se lee solo junto a 'skewed/tilted' (no en el resto de la oración:
    "inflation ran significantly above 2 percent" no hace más sesgado el riesgo)."""
    if not frase:
        return None
    f = frase.lower()
    # Oración con ambos temas ("... inflation ... upside and those to employment ... downside"):
    # quedarse con el tramo que habla del tema pedido.
    if tema_pat and "upside" in f and "downside" in f:
        tramos = re.split(r";|,?\s+and those\b|,?\s+while\b|,?\s+whereas\b|,?\s+but\b", f)
        propios = [t for t in tramos if re.search(tema_pat, t, re.I) and re.search(r"upside|downside", t)]
        if propios:
            f = propios[0]
    m = RE_SESGO_RIESGO.search(f)
    if not m:
        if "balanced" in f:
            return 0.0
        signo = 1 if "upside" in f else -1 if "downside" in f else 0
        return float(signo) if signo else None
    signo = 1 if m.group("dir") == "upside" else -1
    mods = m.group("pre") + m.group("post")
    if re.search(r"\b(?:a little|a bit|slightly|somewhat) more\b", mods):
        mag = 1.25
    elif re.search(r"\bmore\b|significantly|considerably|substantially", mods):
        mag = 1.5
    elif re.search(r"somewhat|slightly|a little|modestly", mods):
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

    # hspace deja lugar a las fechas rotadas del panel de arriba (si no, tapan el título del staff)
    fig, axs = plt.subplots(2, 1, figsize=(ancho, 7.0), gridspec_kw={"height_ratios": [5, 3], "hspace": 0.45})
    im = g_heatmap(axs[0], panel, fechas, "participantes", temas_p, "Tono de los participantes por tema")
    g_heatmap(axs[1], panel, fechas, "staff", temas_s, "Tono del staff por tema")
    cb = fig.colorbar(im, ax=axs, shrink=0.6, pad=0.01)
    cb.set_label("− dovish / débil     + hawkish / fuerte", fontsize=8)
    fig.text(0.01, 0.03, "* Estabilidad financiera: + calma, − preocupación. Celdas grises: menos de "
             f"{MIN_CLAUSULAS} cláusulas.", fontsize=7, color=C_GRIS)
    _pie(fig); fig.savefig(carpeta / "03_mapa_calor_temas.png", dpi=170, bbox_inches="tight"); plt.close(fig)

    fig, axs = plt.subplots(3, 1, figsize=(ancho, 10), sharex=False)
    g_datos_vs_tono(axs, dec, panel, fechas); guardar(fig, "04_datos_vs_tono.png")

    fig, ax = plt.subplots(figsize=(ancho, 4.2)); g_riesgos(ax, dec, fechas); guardar(fig, "05_riesgos_staff.png")
    fig, ax = plt.subplots(figsize=(ancho, 4.5)); g_motores(ax, dec, fechas); guardar(fig, "06_motores_narrativa.png")
    fig, ax = plt.subplots(figsize=(ancho, 0.42 * len(FRASES_COMUNICADO) + 1.4))
    g_lenguaje(ax, dec, fechas); guardar(fig, "07_lenguaje_comunicado.png")

    # Lámina resumen
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
def diagnostico(parrafos: list[dict], fechas: list[str]):
    conteo = (pd.DataFrame(parrafos, columns=["fecha", "seccion"])
                .groupby(["fecha", "seccion"]).size().unstack(fill_value=0)
                .reindex([str(f) for f in fechas], fill_value=0))  # una minuta sin párrafos también sale
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
    diagnostico(parrafos, sorted(rutas))

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
