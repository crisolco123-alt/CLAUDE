#!/usr/bin/env python3
"""
Organizador de libros de medicina.

Uso (desde Terminal en el Mac):

  1) Escanear y generar el plan (NO toca ningún archivo):
     python3 organizar_libros.py escanear

  2) Revisar/editar "plan_organizacion.csv" (Numbers o Excel) e "informe.txt".

  3) Aplicar el plan (renombra y mueve):
     python3 organizar_libros.py aplicar

  4) Si algo no te gusta, deshacer:
     python3 organizar_libros.py deshacer

Nunca borra nada: los duplicados exactos se mueven a la carpeta "_Duplicados (revisar)".
Solo usa la biblioteca estándar de Python 3.
"""

import argparse
import csv
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import unicodedata
from collections import defaultdict
from datetime import datetime
from pathlib import Path

BIBLIOTECA_POR_DEFECTO = (
    "/Users/cristinaolivercolin/Documents/2) RESIDENCIA MFyC/"
    "4) Cursos i Formació/9) Llibres"
)
ORIGEN_POR_DEFECTO = "LIBROSSSSS"  # relativo a la biblioteca
CARPETA_DUPLICADOS = "_Duplicados (revisar)"
CARPETA_SIN_CLASIFICAR = "_Sin Clasificar (revisar)"

EXTENSIONES = {
    ".pdf", ".epub", ".mobi", ".azw3", ".djvu", ".chm",
    ".doc", ".docx", ".ppt", ".pptx",
    ".png", ".jpg", ".jpeg", ".webp", ".heic",
}
EXT_IMAGEN = {".png", ".jpg", ".jpeg", ".webp", ".heic"}

# ---------------------------------------------------------------------------
# Especialidades: nombre de carpeta propuesto -> palabras clave (sin acentos,
# en minúscula). Se buscan en el nombre del archivo y en el título del PDF.
# Si ya tienes una carpeta que coincide (p. ej. "3) Cardiologia"), se usa esa.
# ---------------------------------------------------------------------------
ESPECIALIDADES = {
    "Medicina Interna": ["harrison", "farreras", "medicina interna", "cecil",
                         "manual washington", "washington", "12 de octubre",
                         "doce de octubre", "mcgraw internal"],
    "Medicina de Familia": ["atencion primaria", "medicina de familia",
                            "medicina familiar", "zurro", "semfyc", "fisterra",
                            "amf", "guia terapeutica", "mfyc", "papps"],
    "Cardiología": ["cardio", "corazon", "ecg", "electrocardio", "arritmi",
                    "braunwald", "coronari", "hipertension", "insuficiencia cardiaca",
                    "fibrilacion auricular", "valvulopat", "esc guidelines"],
    "Neumología": ["neumo", "pulmon", "respirator", "asma", "epoc", "espirometr",
                   "gesepoc", "gema", "gold", "ventilacion", "tuberculosis"],
    "Digestivo": ["digestiv", "gastro", "hepat", "higado", "sleisenger",
                  "endoscop", "cirrosis", "colon", "pancrea", "enfermedad inflamatoria intestinal"],
    "Endocrinología y Nutrición": ["endocrin", "diabet", "tiroid", "obesidad",
                                   "nutricion", "insulina", "dislipem", "lipid"],
    "Nefrología": ["nefro", "renal", "rinon", "dialisis", "hidroelectrol",
                   "electrolit", "equilibrio acido"],
    "Neurología": ["neurolog", "cefalea", "ictus", "epilep", "parkinson",
                   "demencia", "adams y victor", "neuroanatom"],
    "Psiquiatría": ["psiqui", "psicofarm", "dsm", "depresion", "ansiedad",
                    "kaplan", "stahl", "salud mental", "esquizofren"],
    "Enfermedades Infecciosas": ["infecci", "antibiot", "sanford", "mensa",
                                 "vih", "antimicrob", "mandell", "microbiolog",
                                 "sepsis", "vacun"],
    "Hematología": ["hematolog", "anemia", "anticoag", "leucemia", "linfoma",
                    "hemostasia", "trombos"],
    "Oncología": ["oncolog", "cancer", "tumor", "devita", "quimioterap"],
    "Reumatología": ["reumatolog", "artritis", "lupus", "gota", "espondil",
                     "vasculitis", "kelley"],
    "Dermatología": ["dermat", "piel", "fitzpatrick", "lesiones cutaneas"],
    "Pediatría": ["pediatr", "nelson", "neonat", "lactante", "infantil", "adolescen"],
    "Ginecología y Obstetricia": ["ginecolog", "obstetri", "embaraz", "anticoncep",
                                  "menopaus", "gestacion", "williams obstetric"],
    "Urología": ["urolog", "prostat", "incontinencia", "campbell"],
    "Oftalmología": ["oftalmolog", "kanski", "retina", "glaucoma", "ojo rojo"],
    "Otorrinolaringología": ["otorrino", "orl", "vertigo", "hipoacusia"],
    "Traumatología": ["traumatolog", "ortoped", "fractur", "musculoesquel",
                      "rodilla", "hombro", "columna", "infiltracion"],
    "Urgencias": ["urgenc", "emergenc", "critic", "uci", "reanimac", "rcp",
                  "toxicolog", "intoxicac", "tintinalli", "soporte vital"],
    "Geriatría y Paliativos": ["geriatr", "anciano", "paliativ", "fragilidad",
                               "final de vida"],
    "Farmacología": ["farmacolog", "goodman", "medicamento", "vademecum",
                     "florez", "prescripcion"],
    "Radiología y Ecografía": ["radiolog", "ecograf", "ecoscop", "imagen medica",
                               "tac ", "resonancia", "radiografia", "pocus"],
    "Semiología y Exploración": ["semiolog", "exploracion fisica", "anamnesis",
                                 "argente", "suros", "bates", "propedeut"],
    "Ciencias Básicas": ["anatomi", "fisiolog", "guyton", "netter", "histolog",
                         "bioquim", "embriolog", "patologia estructural", "robbins"],
    "Investigación y Estadística": ["estadistic", "investigacion", "epidemiolog",
                                    "metodolog", "medicina basada en la evidencia",
                                    "lectura critica"],
    "MIR": ["amir", "ctomir", "mir ", "manual cto", "desgloses"],
}

# Palabras que van en minúscula (salvo si son la primera palabra)
MINUSCULAS = {
    "el", "la", "los", "las", "lo", "un", "una", "unos", "unas",
    "de", "del", "al", "a", "y", "e", "o", "u", "en", "con", "por", "para",
    "sin", "sobre", "entre", "desde", "hasta", "según", "segun", "ante", "tras",
    "i", "amb", "per", "els", "les", "d", "l",          # catalán
    "the", "of", "and", "in", "for", "on", "to", "an", "at", "by", "with", "or",
}

# Siglas que se deben mantener en mayúsculas aunque el archivo venga en minúscula
SIGLAS = {
    "ecg", "epoc", "vih", "sida", "dsm", "mir", "amir", "cto", "uci", "rcp", "orl",
    "esc", "aha", "acc", "ada", "gold", "gema", "gesepoc", "semfyc", "semergen",
    "semg", "ics", "camfic", "pocus", "tac", "rm", "rx", "ets", "its", "hta", "dm",
    "irc", "erc", "iam", "sca", "tep", "tvp", "fa", "ic", "eii", "nice", "who",
    "oms", "cdc", "idsa", "aeped", "sego", "sen", "seen", "sep", "ser", "sec",
    "semes", "semicyuc", "atls", "acls", "bls", "svb", "sva", "ebm", "mbe",
    "ii", "iii", "iv", "vi", "vii", "viii", "ix", "xi", "xii", "xiii", "xiv", "xv",
    "papps", "amf", "ap", "mfyc",
}

# Nombres con mayúsculas/minúsculas propias
FORMA_EXACTA = {"semfyc": "semFYC", "camfic": "CAMFiC", "mfyc": "MFyC",
                "ecografia": "Ecografía", "uptodate": "UpToDate",
                "amir": "AMIR", "cto": "CTO", "vih": "VIH", "covid": "COVID"}

# Palabras frecuentes que suelen venir sin tilde en los nombres de archivo
ACENTOS = {
    "guia": "guía", "guias": "guías", "diagnostico": "diagnóstico",
    "diagnostica": "diagnóstica", "estadistico": "estadístico",
    "terapeutica": "terapéutica", "terapeutico": "terapéutico",
    "clinica": "clínica", "clinicas": "clínicas", "clinico": "clínico",
    "clinicos": "clínicos", "medico": "médico", "medica": "médica",
    "medicos": "médicos", "medicas": "médicas", "practica": "práctica",
    "practicas": "prácticas", "practico": "práctico", "basica": "básica",
    "basicas": "básicas", "basico": "básico", "tecnicas": "técnicas",
    "tecnica": "técnica", "cirugia": "cirugía", "critico": "crítico",
    "critica": "crítica", "criticos": "críticos", "rapida": "rápida",
    "rapido": "rápido", "analisis": "análisis", "sintomas": "síntomas",
    "sindrome": "síndrome", "sindromes": "síndromes", "farmaco": "fármaco",
    "farmacos": "fármacos", "pediatrica": "pediátrica", "pediatria": "pediatría",
    "psiquiatria": "psiquiatría", "geriatria": "geriatría", "cardiaca": "cardíaca",
    "cardiaco": "cardíaco", "fisica": "física", "fisico": "físico",
    "metodo": "método", "metodos": "métodos", "pagina": "página",
    "ginecologica": "ginecológica", "ortopedica": "ortopédica",
    "patologia": "patología", "fisiologia": "fisiología", "anatomia": "anatomía",
    "infografia": "infografía", "ecografia": "ecografía", "farmacologia": "farmacología", "semiologia": "semiología",
    "via": "vía", "vias": "vías", "oseo": "óseo",
    "organos": "órganos", "organo": "órgano", "musculo": "músculo", "indice": "índice", "torax": "tórax",
    "higado": "hígado", "rinon": "riñón", "nino": "niño", "ninos": "niños",
    "espanol": "español", "espanola": "española",
}

# Basura típica en nombres de archivo descargados
BASURA = [
    r"\(?z-?lib(?:rary)?(?:\.org)?\)?", r"\(?libgen(?:\.\w+)?\)?", r"\bb-ok\b",
    r"\bwww\.[\w.-]+", r"\b[\w-]+\.(?:com|net|org|es|info)\b",
    r"\bbooksmedicos\b", r"\bmedilibros\b", r"\bmedicomoderno\b", r"\bebook\b",
    r"\bpdf\b", r"\bepub\b", r"\bdescargar\b", r"\bgratis\b", r"\bfull\b",
    r"\bcopia\b", r"\bcopy\b", r"\bfinal\b", r"\bdefinitivo\b",
    r"\(\s*\d\s*\)", r"\[\s*\d\s*\]",
    r"\bisbn[\s:-]*[\dxX-]{10,17}\b", r"\b97[89][\d-]{10,14}\b",
    r"\bspanish\b", r"\bespa[nñ]ol\b", r"\bocr\b",
]

ORDINALES = {
    "primera": 1, "segunda": 2, "tercera": 3, "cuarta": 4, "quinta": 5,
    "sexta": 6, "septima": 7, "octava": 8, "novena": 9, "decima": 10,
    "undecima": 11, "duodecima": 12, "decimotercera": 13, "decimocuarta": 14,
    "decimoquinta": 15, "decimosexta": 16, "decimoseptima": 17,
    "decimoctava": 18, "decimonovena": 19, "vigesima": 20,
    "first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5, "sixth": 6,
    "seventh": 7, "eighth": 8, "ninth": 9, "tenth": 10, "eleventh": 11,
    "twelfth": 12, "thirteenth": 13, "fourteenth": 14, "fifteenth": 15,
    "sixteenth": 16, "seventeenth": 17, "eighteenth": 18, "nineteenth": 19,
    "twentieth": 20, "twenty-first": 21, "twenty-second": 22,
}


# ---------------------------------------------------------------------------
# Utilidades de texto
# ---------------------------------------------------------------------------
def sin_acentos(s):
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def normalizar(s):
    """Minúsculas, sin acentos, solo letras/números y espacios."""
    s = sin_acentos(s.lower())
    s = re.sub(r"[^a-z0-9ñ ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def extraer_edicion(texto):
    """Devuelve (numero_edicion | None, texto_sin_edicion)."""
    t = unicodedata.normalize("NFC", texto)
    patrones = [
        # 22ª Ed / 22a ed / 22 ed. / 22th edition / 22nd Ed / 22 edicion
        r"\b(\d{1,2})\s*(?:ª|º|a|o|th|st|nd|rd|va|ra|da|ta|na|°)?\.?\s*"
        r"(?:ed\b\.?|edic\w*\.?|edition\b|edn\b)",
        # Edición 22 / Edition 22 / ed. 22
        r"\b(?:edici[oó]n|edition|ed\.?)\s*(\d{1,2})\b",
    ]
    for p in patrones:
        m = re.search(p, t, flags=re.IGNORECASE)
        if m:
            return int(m.group(1)), (t[:m.start()] + " " + t[m.end():])
    # Ordinales en palabra: "vigésima edición", "fifth edition"
    m = re.search(r"\b([a-záéíóú-]+)\s+(?:edici[oó]n|edition|ed\b\.?)",
                  t, flags=re.IGNORECASE)
    if m:
        n = ORDINALES.get(sin_acentos(m.group(1).lower()))
        if n:
            return n, (t[:m.start()] + " " + t[m.end():])
    return None, t


def limpiar_basura(texto):
    t = texto
    for p in BASURA:
        t = re.sub(p, " ", t, flags=re.IGNORECASE)
    t = t.replace("_", " ")
    t = re.sub(r"(?<=\w)\.(?=\w)", " ", t)  # puntos entre palabras
    t = re.sub(r"\s-\s|--+|\+", " ", t)
    t = re.sub(r"(?<=[a-zA-Z])(?=\d{1,2}\s*(?:ª|º)?\s*ed)", " ", t,
               flags=re.IGNORECASE)  # "Harrison22ed"
    for p in BASURA:
        t = re.sub(p, " ", t, flags=re.IGNORECASE)
    t = re.sub(r"[\[\]{}]", " ", t)
    t = re.sub(r"\(\s*\)", " ", t)
    # separar camelCase tipo "PrincipiosDeMedicina"
    t = re.sub(r"(?<=[a-záéíóúñ])(?=[A-ZÁÉÍÓÚÑ][a-záéíóúñ])", " ", t)
    return re.sub(r"\s+", " ", t).strip(" -.,;")


def acentuar(palabra):
    """Añade la tilde a palabras españolas comunes que vienen sin ella."""
    low = palabra.lower()
    if low in ACENTOS:
        return ACENTOS[low]
    if low != sin_acentos(low):  # ya lleva tilde
        return low
    for suf, rep in (("ciones", "ciones"), ("cion", "ción"), ("logia", "logía"),
                     ("logias", "logías"), ("grafia", "grafía"), ("iatria", "iatría"),
                     ("scopia", "scopia")):
        if low.endswith(suf) and len(low) > len(suf) + 1:
            return low[: -len(suf)] + rep
    return low


def formato_titulo(texto):
    letras = [c for c in texto if c.isalpha()]
    if letras and sum(c.isupper() for c in letras) / len(letras) > 0.6:
        texto = texto.lower()  # venía TODO EN MAYÚSCULAS
    palabras = texto.split()
    salida = []
    for i, w in enumerate(palabras):
        base = normalizar(w)
        prefijo = re.match(r"^[\(\"'¿¡]*", w).group(0)
        nucleo = w[len(prefijo):]
        if base in FORMA_EXACTA:
            nuevo = FORMA_EXACTA[base]
        elif base in SIGLAS:
            nuevo = nucleo.upper()
        elif re.fullmatch(r"[A-Z0-9]{2,6}(?:-\d+)?", nucleo) and not nucleo.isdigit():
            nuevo = nucleo  # ya es una sigla (DSM-5, AMIR...)
        elif i > 0 and base in MINUSCULAS:
            nuevo = nucleo.lower()
        elif "-" in nucleo:
            nuevo = "-".join(acentuar(p)[:1].upper() + acentuar(p)[1:]
                             for p in nucleo.split("-"))
        else:
            a = acentuar(nucleo) if nucleo.isalpha() else nucleo.lower()
            nuevo = a[:1].upper() + a[1:]
        salida.append(prefijo + nuevo)
    return " ".join(salida)


def detectar_tipo(nombre, ext):
    n = normalizar(nombre)
    if ext in EXT_IMAGEN or "infografi" in n or "poster" in n or "esquema" in n:
        return "Infografía"
    if re.search(r"\b(guia|guideline|guidelines|consenso|protocolo|algoritmo|"
                 r"recomendaciones|documento de consenso|pcc)\b", n):
        return "Guía"
    if re.search(r"\b(manual|tratado|atlas|principios|libro|textbook)\b", n):
        return "Libro"
    return "Libro"


def titulo_metadatos(ruta):
    """Título interno del PDF/epub vía Spotlight (solo macOS). Vacío si no hay."""
    if sys.platform != "darwin":
        return ""
    try:
        r = subprocess.run(["mdls", "-raw", "-name", "kMDItemTitle", str(ruta)],
                           capture_output=True, text=True, timeout=10)
        t = r.stdout.strip()
        if t and t != "(null)" and len(t) > 4:
            return t
    except Exception:
        pass
    return ""


def nombre_parece_basura(nombre):
    n = normalizar(nombre)
    letras = re.sub(r"[^a-z]", "", n)
    return (len(letras) < 6 or re.fullmatch(r"[a-f0-9 ]{12,}", n) is not None
            or re.match(r"^(scan|doc|documento|img|image|file|download|untitled|"
                        r"sin titulo)\b", n) is not None)


def proponer_nombre(ruta):
    """Devuelve (nombre_nuevo_con_ext, clave_titulo, edicion, tipo, notas)."""
    ext = ruta.suffix.lower()
    original = ruta.stem
    notas = []
    fuente = original
    meta = ""
    if nombre_parece_basura(original):
        meta = titulo_metadatos(ruta)
        if meta and not nombre_parece_basura(meta):
            fuente = meta
            notas.append("título sacado de los metadatos del PDF")
        else:
            notas.append("nombre poco claro: revisar a mano")

    fuente = limpiar_basura(fuente)
    ed, resto = extraer_edicion(fuente)
    if ed is None and fuente is not original:
        ed, _ = extraer_edicion(limpiar_basura(original))
    limpio = limpiar_basura(resto)
    # Quitar años sueltos al final solo si hay edición (en guías el año es útil)
    tipo = detectar_tipo(fuente, ext)
    if ed and tipo == "Libro":
        limpio = re.sub(r"\s*\(?\b(19|20)\d{2}\b\)?\s*$", "", limpio)
    titulo = formato_titulo(limpio) or formato_titulo(original)
    if ed:
        titulo = f"{titulo} {ed}ª Ed"
    titulo = re.sub(r'[/:\\?*"<>|]', "-", titulo).strip()
    clave = normalizar(limpio)
    return titulo + ext, clave, ed, tipo, notas


# ---------------------------------------------------------------------------
# Clasificación por especialidad
# ---------------------------------------------------------------------------
def puntuar(texto):
    n = " " + normalizar(texto) + " "
    puntos = {}
    for esp, claves in ESPECIALIDADES.items():
        p = sum(1 for c in claves if c in n)
        if p:
            puntos[esp] = p
    return puntos


def mejor_especialidad(texto):
    p = puntuar(texto)
    if not p:
        return None, 0
    esp = max(p, key=lambda k: p[k])
    return esp, p[esp]


def carpetas_existentes(biblioteca, origen):
    """Mapa especialidad -> carpeta ya existente en la biblioteca."""
    mapa = {}
    existentes = [d for d in biblioteca.iterdir()
                  if d.is_dir() and d.resolve() != origen.resolve()
                  and not d.name.startswith((".", "_"))]
    for esp in ESPECIALIDADES:
        raiz = normalizar(esp).split()[0][:6]  # "cardio", "neumol", "medici"...
        candidatas = [d for d in existentes if raiz in normalizar(d.name)]
        if esp == "Medicina Interna":
            candidatas = [d for d in existentes if "interna" in normalizar(d.name)]
        elif esp == "Medicina de Familia":
            candidatas = [d for d in existentes
                          if re.search(r"familia|primaria|mfyc", normalizar(d.name))]
        if candidatas:
            mapa[esp] = candidatas[0]
    return mapa, existentes


def especialidad_de_carpeta(carpeta, mapa):
    for esp, d in mapa.items():
        if d.resolve() == carpeta.resolve():
            return esp
    return None


# ---------------------------------------------------------------------------
# Duplicados
# ---------------------------------------------------------------------------
def hash_archivo(ruta, bloque=1 << 20):
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        while True:
            b = f.read(bloque)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def buscar_archivos(carpeta):
    for dirpath, dirnames, filenames in os.walk(carpeta):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
        for f in filenames:
            p = Path(dirpath) / f
            if p.suffix.lower() in EXTENSIONES and not f.startswith("."):
                yield p


# ---------------------------------------------------------------------------
# Comandos
# ---------------------------------------------------------------------------
def escanear(args):
    biblioteca = Path(args.biblioteca).expanduser()
    origen = (biblioteca / args.origen) if not os.path.isabs(args.origen) else Path(args.origen)
    if not origen.is_dir():
        sys.exit(f"No encuentro la carpeta de origen: {origen}")

    mapa, existentes = carpetas_existentes(biblioteca, origen)
    print(f"Biblioteca: {biblioteca}")
    print(f"Carpetas de especialidad detectadas: {len(mapa)}")
    for esp, d in sorted(mapa.items()):
        print(f"   {esp:32s} -> {d.name}")

    todos = list(buscar_archivos(biblioteca))
    en_origen = [p for p in todos if origen.resolve() in p.resolve().parents]
    ya_clasificados = [p for p in todos if p not in set(en_origen)
                       and CARPETA_DUPLICADOS not in p.parts]
    print(f"Archivos en {origen.name}: {len(en_origen)}  |  "
          f"ya clasificados: {len(ya_clasificados)}")

    # --- Duplicados exactos (mismo contenido) ---
    print("Buscando duplicados exactos (puede tardar un rato)...")
    por_tamano = defaultdict(list)
    for p in todos:
        try:
            por_tamano[p.stat().st_size].append(p)
        except OSError:
            pass
    grupos_hash = defaultdict(list)
    for tam, lista in por_tamano.items():
        if len(lista) > 1:
            for p in lista:
                try:
                    grupos_hash[hash_archivo(p)].append(p)
                except OSError:
                    pass
    dup_exactos = [g for g in grupos_hash.values() if len(g) > 1]

    # Qué copia se queda: la que ya está clasificada; si no, la de nombre más completo
    set_origen = set(en_origen)
    descartar = {}
    for g in dup_exactos:
        def calidad(p):
            _, clave, ed, _, _ = proponer_nombre(p)
            return (p in set_origen, -(ed is not None), -len(clave))
        g_orden = sorted(g, key=calidad)
        for p in g_orden[1:]:
            descartar[p] = g_orden[0]

    # --- Plan para los archivos del origen ---
    filas = []
    claves = defaultdict(list)  # clave_titulo -> [(ruta, edicion, nombre)]
    usados = defaultdict(set)   # carpeta -> nombres

    def destino_para(esp):
        if esp is None:
            return biblioteca / CARPETA_SIN_CLASIFICAR
        return mapa.get(esp, biblioteca / esp)

    for p in sorted(en_origen):
        nuevo, clave, ed, tipo, notas = proponer_nombre(p)
        texto_clasif = f"{p.stem} {nuevo} {p.parent.name}"
        esp, conf = mejor_especialidad(texto_clasif)
        if esp is None:
            meta = titulo_metadatos(p)
            if meta:
                esp, conf = mejor_especialidad(meta)
        carpeta = destino_para(esp)
        if esp and esp not in mapa:
            notas.append(f"se creará la carpeta nueva '{esp}'")
        if p in descartar:
            accion = "DUPLICADO"
            carpeta = biblioteca / CARPETA_DUPLICADOS
            notas.insert(0, f"idéntico a: {descartar[p].relative_to(biblioteca)}")
        else:
            accion = "MOVER"
            base, ext = os.path.splitext(nuevo)
            n = 2
            while nuevo.lower() in usados[str(carpeta)] or (
                    (carpeta / nuevo).exists() and (carpeta / nuevo).resolve() != p.resolve()):
                nuevo = f"{base} ({n}){ext}"
                n += 1
            usados[str(carpeta)].add(nuevo.lower())
            claves[(clave, tipo)].append((p, ed, nuevo))
        filas.append({
            "accion": accion,
            "ruta_actual": str(p.relative_to(biblioteca)),
            "nombre_nuevo": nuevo,
            "carpeta_destino": str(carpeta.relative_to(biblioteca)),
            "tipo": tipo,
            "especialidad": esp or "?",
            "confianza": {0: "baja"}.get(conf, "media" if conf == 1 else "alta"),
            "notas": "; ".join(notas),
        })

    # Añadir también claves de libros ya clasificados para detectar ediciones repetidas
    for p in ya_clasificados:
        _, clave, ed, tipo, _ = proponer_nombre(p)
        claves[(clave, tipo)].append((p, ed, p.name))

    # --- Posibles duplicados (mismo título, distinto archivo / edición) ---
    # Se agrupan títulos iguales o en los que uno empieza por el otro
    # ("farreras" ~ "farreras rozman medicina interna"), ignorando números.
    genericas = {"manual", "guia", "tratado", "atlas", "principios", "protocolo",
                 "libro", "curso", "tema", "temas", "apuntes", "resumen", "esquema",
                 "infografia", "medicina", "introduccion", "fundamentos", "compendio"}
    items = []
    for (clave, tipo), lista in claves.items():
        k = re.sub(r"\b\d+\b", " ", clave).split()
        for elem in lista:
            items.append((k, tipo, elem))
    padre = list(range(len(items)))

    def raiz(i):
        while padre[i] != i:
            padre[i] = padre[padre[i]]
            i = padre[i]
        return i

    por_primera = defaultdict(list)
    for i, (k, tipo, _) in enumerate(items):
        if k:
            por_primera[(k[0], tipo)].append(i)
    for idxs in por_primera.values():
        for a_i, a in enumerate(idxs):
            for b in idxs[a_i + 1:]:
                ka, kb = items[a][0], items[b][0]
                corto, largo = sorted((ka, kb), key=len)
                if largo[:len(corto)] != corto:
                    continue
                if len(corto) == 1 and (corto[0] in genericas or len(corto[0]) < 5):
                    continue
                if len(" ".join(corto)) < 5:
                    continue
                padre[raiz(a)] = raiz(b)
    grupos = defaultdict(list)
    for i, (k, tipo, elem) in enumerate(items):
        grupos[raiz(i)].append((" ".join(k), elem))
    posibles = []
    for g in grupos.values():
        if len(g) > 1:
            clave = max((c for c, _ in g), key=len)
            posibles.append((clave, [e for _, e in g]))

    # --- Mal clasificados (libros ya en carpetas de especialidad) ---
    mal = []
    for p in ya_clasificados:
        carpeta_top = biblioteca / p.relative_to(biblioteca).parts[0]
        actual = especialidad_de_carpeta(carpeta_top, mapa)
        if not actual:
            continue
        pts = puntuar(p.stem)
        sugerida, conf = mejor_especialidad(p.stem)
        if sugerida and sugerida != actual and pts.get(actual, 0) == 0 and conf >= 1:
            mal.append((p, carpeta_top.name, sugerida, conf))

    # --- Guardar CSV ---
    salida = Path(args.plan)
    with open(salida, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(filas[0].keys()) if filas else
                           ["accion", "ruta_actual", "nombre_nuevo", "carpeta_destino",
                            "tipo", "especialidad", "confianza", "notas"],
                           delimiter=";")
        w.writeheader()
        w.writerows(filas)
    with open(salida.with_suffix(".json"), "w", encoding="utf-8") as f:
        json.dump({"biblioteca": str(biblioteca), "origen": str(origen)}, f)

    # --- Informe ---
    lineas = [f"INFORME DE ORGANIZACIÓN — {datetime.now():%d/%m/%Y %H:%M}", "=" * 70, ""]
    lineas.append(f"Archivos a organizar: {len(filas)}  "
                  f"(sin clasificar: {sum(1 for r in filas if r['especialidad'] == '?')})")
    lineas.append("")
    lineas.append(f"1) DUPLICADOS EXACTOS (mismo archivo): {len(dup_exactos)} grupos")
    lineas.append("-" * 70)
    for g in dup_exactos:
        for p in g:
            marca = "  [se queda]" if p not in descartar else "  [→ _Duplicados]"
            lineas.append(f"   {p.relative_to(biblioteca)}{marca}")
        lineas.append("")
    lineas.append(f"2) POSIBLES DUPLICADOS / VARIAS EDICIONES: {len(posibles)} grupos")
    lineas.append("-" * 70)
    for clave, lista in sorted(posibles):
        eds = [e for _, e, _ in lista if e]
        for p, e, nombre in sorted(lista, key=lambda x: x[1] or 0, reverse=True):
            extra = ""
            if eds and e and e < max(eds):
                extra = f"   ← edición antigua (tienes la {max(eds)}ª)"
            lineas.append(f"   {p.relative_to(biblioteca)}  →  {nombre}{extra}")
        lineas.append("")
    lineas.append(f"3) POSIBLEMENTE MAL CLASIFICADOS: {len(mal)}")
    lineas.append("-" * 70)
    for p, actual, sug, conf in mal:
        lineas.append(f"   {p.relative_to(biblioteca)}")
        lineas.append(f"        está en: {actual}   →   parece: {sug}")
    lineas.append("")
    lineas.append("4) SIN CLASIFICAR / NOMBRE POCO CLARO (revisar a mano)")
    lineas.append("-" * 70)
    for r in filas:
        if r["especialidad"] == "?" or "revisar" in r["notas"]:
            lineas.append(f"   {r['ruta_actual']}  →  {r['nombre_nuevo']}")
    Path(args.informe).write_text("\n".join(lineas), encoding="utf-8")

    print()
    print(f"✔ Plan guardado en:    {salida.resolve()}")
    print(f"✔ Informe guardado en: {Path(args.informe).resolve()}")
    print(f"   Duplicados exactos: {len(dup_exactos)} grupos | posibles duplicados: "
          f"{len(posibles)} | posibles mal clasificados: {len(mal)}")
    print("Revisa y edita el CSV (puedes cambiar 'nombre_nuevo', 'carpeta_destino' o poner"
          " 'IGNORAR' en 'accion'). Luego ejecuta:  python3 organizar_libros.py aplicar")


def aplicar(args):
    plan = Path(args.plan)
    if not plan.exists():
        sys.exit("No existe el plan. Ejecuta primero: python3 organizar_libros.py escanear")
    meta = json.loads(plan.with_suffix(".json").read_text(encoding="utf-8"))
    biblioteca = Path(meta["biblioteca"])
    with open(plan, encoding="utf-8-sig") as f:
        filas = list(csv.DictReader(f, delimiter=";"))

    registro = []
    hechos = errores = 0
    for r in filas:
        accion = r["accion"].strip().upper()
        if accion not in ("MOVER", "DUPLICADO"):
            continue
        src = biblioteca / r["ruta_actual"]
        dst_dir = biblioteca / r["carpeta_destino"]
        dst = dst_dir / r["nombre_nuevo"].strip()
        if not src.exists():
            print(f"  ✗ ya no existe: {src.name}")
            errores += 1
            continue
        if src.resolve() == dst.resolve():
            continue
        dst_dir.mkdir(parents=True, exist_ok=True)
        base, ext = os.path.splitext(dst.name)
        n = 2
        while dst.exists():
            dst = dst_dir / f"{base} ({n}){ext}"
            n += 1
        try:
            shutil.move(str(src), str(dst))
            registro.append({"de": str(src), "a": str(dst)})
            hechos += 1
        except OSError as e:
            print(f"  ✗ {src.name}: {e}")
            errores += 1

    log = Path(f"deshacer_{datetime.now():%Y%m%d_%H%M%S}.json")
    log.write_text(json.dumps(registro, ensure_ascii=False, indent=1), encoding="utf-8")

    # Borrar carpetas vacías que hayan quedado DENTRO del origen (nunca otras)
    origen = Path(meta["origen"])
    for dirpath, dirnames, filenames in os.walk(origen, topdown=False):
        d = Path(dirpath)
        if d == origen:
            continue
        restos = [x for x in d.iterdir() if x.name != ".DS_Store"]
        if not restos:
            try:
                ds = d / ".DS_Store"
                if ds.exists():
                    ds.unlink()
                d.rmdir()
            except OSError:
                pass

    print(f"✔ {hechos} archivos movidos/renombrados, {errores} errores.")
    print(f"  Para deshacer: python3 organizar_libros.py deshacer {log}")


def deshacer(args):
    logs = sorted(Path(".").glob("deshacer_*.json"))
    log = Path(args.log) if args.log else (logs[-1] if logs else None)
    if not log or not log.exists():
        sys.exit("No encuentro ningún registro para deshacer.")
    registro = json.loads(log.read_text(encoding="utf-8"))
    n = 0
    for mov in reversed(registro):
        a, de = Path(mov["a"]), Path(mov["de"])
        if a.exists() and not de.exists():
            de.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(a), str(de))
            n += 1
    log.rename(log.with_suffix(".deshecho"))
    print(f"✔ {n} archivos devueltos a su sitio original.")


def main():
    ap = argparse.ArgumentParser(description="Organiza y renombra libros de medicina.")
    ap.add_argument("--biblioteca", default=BIBLIOTECA_POR_DEFECTO,
                    help="Carpeta raíz con las carpetas por especialidad")
    ap.add_argument("--origen", default=ORIGEN_POR_DEFECTO,
                    help="Subcarpeta desordenada a organizar")
    ap.add_argument("--plan", default="plan_organizacion.csv")
    ap.add_argument("--informe", default="informe.txt")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("escanear", help="Analiza y genera plan + informe (no mueve nada)")
    sub.add_parser("aplicar", help="Aplica el plan del CSV")
    d = sub.add_parser("deshacer", help="Deshace el último 'aplicar'")
    d.add_argument("log", nargs="?")
    args = ap.parse_args()
    {"escanear": escanear, "aplicar": aplicar, "deshacer": deshacer}[args.cmd](args)


if __name__ == "__main__":
    main()
