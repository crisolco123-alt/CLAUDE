#!/usr/bin/env python3
"""
Organizador de libros de medicina.

Uso (desde Terminal en el Mac, dentro de la carpeta "9) Llibres"):

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
# Especialidades.
#   carpeta:  expresión que reconoce tu carpeta ya existente (p. ej. "10) CARDIOVASCULAR")
#   claves:   palabras clave (sin tildes). Coinciden con el INICIO de una palabra:
#             "cardio" encaja con "cardiología" y "cardiovascular", pero "asma" no
#             encaja con "eritrasma". Un espacio final obliga a palabra completa.
#   general:  las generales pierden los empates (un "Manual Washington de
#             Gastroenterología" va a Digestivo, no a Medicina Interna).
# ---------------------------------------------------------------------------
ESPECIALIDADES = {
    "Medicina Interna": dict(
        carpeta=r"interna", general=True,
        claves=["harrison", "farreras", "medicina interna", "cecil", "washington",
                "12 de octubre", "12 octubre", "doce de octubre", "terapeutica medica",
                "medicina de bolsillo", "sabatine", "clinical advisor", "ferri",
                "diagnostico clinico", "green book", "diagnostico diferencial",
                "biblia medico", "padecimientos", "huppert", "mnemonics",
                "diccionario medico", "rotaciones clinicas", "soap "]),
    "Medicina de Familia": dict(
        carpeta=r"familia|primaria|mfyc", general=True,
        claves=["atencion primaria", "medicina de familia", "medicina familiar",
                "zurro", "semfyc", "camfic", "fisterra", "amf ", "mfyc", "papps",
                "brujula", "50 principales consultas"]),
    "Urgencias": dict(
        carpeta=r"urgenc|emergenc", general=True,
        claves=["urgenc", "emergenc", "reanimac", "rcp ", "toxicolog", "intoxicac",
                "tintinalli", "rosen", "soporte vital", "phtls", "atls", "acls",
                "desfibril", "guardia", "compresion mecanica"]),
    "UCI": dict(
        carpeta=r"\buci\b|intensiv|critic",
        claves=["uci ", "cuidados critic", "cuidados intensiv", "ventilacion mecanica",
                "medicina intensiva", "paciente critico", "vexus"]),
    "Anestesiología": dict(
        carpeta=r"anestes",
        claves=["anestes", "via aerea", "walls", "dolor", "pain", "melzack"]),
    "Cardiología": dict(
        carpeta=r"cardio",
        claves=["cardio", "corazon", "braunwald", "coronari", "hipertension",
                "insuficiencia cardiaca", "fibrilacion auricular", "valvulopat",
                "auscultacion", "perfil cardiaco", "sketch med cardiologia"]),
    "ECG": dict(
        carpeta=r"\becg\b|\bekg\b|electrocardio",
        claves=["ecg ", "ekg ", "electrocardio", "arritmi", "holter", "qrs ",
                "taquicardia", "huszar"]),
    "Neumología": dict(
        carpeta=r"neumo|respirat|pulmon",
        claves=["neumo", "pulmon", "respirator", "asma ", "epoc ", "espirometr",
                "gesepoc", "gema ", "gold ", "tuberculosis", "torres duque"]),
    "Digestivo": dict(
        carpeta=r"digest|gastro",
        claves=["digestiv", "gastro", "hepat", "higado", "sleisenger", "endoscop",
                "cirrosis", "colon ", "pancrea", "diarrea"]),
    "Endocrinología y Nutrición": dict(
        carpeta=r"endocrin",
        claves=["endocrin", "diabet", "tiroid", "obesidad", "nutricion", "insulina",
                "dislipem", "dislipid", "lipid", "cetoacidosis", "hiperglucem",
                "hiperosmolar", "ada "]),
    "Nefrología": dict(
        carpeta=r"nefro",
        claves=["nefro", "renal", "rinon", "dialisis", "hidroelectrol", "electrolit",
                "acido base", "acido-base", "arias rodriguez", "ortiz arduan"]),
    "Neurología": dict(
        carpeta=r"neurol",
        claves=["neurolog", "cefalea", "ictus", "epilep", "parkinson", "demencia",
                "adams ", "neuroanatom", "neurocienc", "sistema nervioso",
                "cerebr", "ataque cerebro"]),
    "Psiquiatría": dict(
        carpeta=r"psiqui",
        claves=["psiqui", "psicofarm", "dsm", "depresion", "ansiedad", "kaplan",
                "stahl", "salud mental", "esquizofren", "drogodepend", "adiccion",
                "delirium"]),
    "Psicología": dict(
        carpeta=r"psicol",
        claves=["psicolog", "psicoterap", "autismo", "problemas de conduct",
                "trastornos de la conduct"]),
    "Enfermedades Infecciosas": dict(
        carpeta=r"infecc",
        claves=["infecci", "infecto", "antibiot", "sanford", "mensa ", "vih ",
                "antimicrob", "mandell", "microbiolog", "sepsis", "vacun",
                "inmunizacion"]),
    "Hematología": dict(
        carpeta=r"hemato",
        claves=["hematolog", "anemia", "anticoag", "leucemia", "linfoma",
                "hemostasia", "trombos", "hemoterap"]),
    "Oncología": dict(
        carpeta=r"oncol",
        claves=["oncolog", "cancer", "tumor", "devita", "quimioterap"]),
    "Reumatología": dict(
        carpeta=r"reuma",
        claves=["reumatolog", "artritis", "lupus", "gota ", "espondil", "vasculitis",
                "kelley", "firestein", "koretzky"]),
    "Dermatología": dict(
        carpeta=r"dermat",
        claves=["dermat", "piel ", "fitzpatrick", "bolognia", "lesiones cutaneas",
                "tegumentario", "penfigo", "piodermitis", "impetigo", "erisipela"]),
    "Pediatría": dict(
        carpeta=r"pediat",
        claves=["pediatr", "nelson", "neonat", "lactante", "infantil", "adolescen"]),
    "Ginecología y Obstetricia": dict(
        carpeta=r"gineco|obstet",
        claves=["ginecolog", "obstetri", "embaraz", "anticoncep", "menopaus",
                "gestacion"]),
    "Urología": dict(
        carpeta=r"urolog",
        claves=["urolog", "prostat", "incontinencia", "aparato gu", "genitourin"]),
    "Oftalmología": dict(
        carpeta=r"oftalm",
        claves=["oftalmolog", "kanski", "retina", "glaucoma", "ojo rojo", "dmae"]),
    "Otorrinolaringología": dict(
        carpeta=r"otorrino|\borl\b",
        claves=["otorrino", "orl ", "vertigo", "hipoacusia", "otolog", "lalwani"]),
    "Traumatología": dict(
        carpeta=r"trauma|\bcot\b",
        claves=["traumatolog", "ortoped", "fractur", "musculoesquel", "rodilla",
                "hombro", "columna", "infiltracion"]),
    "Medicina Física y Rehabilitación": dict(
        carpeta=r"rehab|medicina fisica",
        claves=["rehabilit", "fascia", "kendall", "lesiones deportivas", "sobotta",
                "musculos", "muscular", "fisioterap", "pruebas funcionales"]),
    "Cirugía General": dict(
        carpeta=r"cirug",
        claves=["cirug", "abdomen agudo", "schwartz", "turegano", "quirurg"]),
    "Geriatría y Paliativos": dict(
        carpeta=r"geriat|paliat",
        claves=["geriatr", "anciano", "paliativ", "fragilidad", "final de vida"]),
    "Farmacología": dict(
        carpeta=r"farmac",
        claves=["farmac", "farma ", "goodman", "medicamento", "vademecum", "florez",
                "prescripcion", "dosis", "lullmann"]),
    "Radiología y Ecografía": dict(
        carpeta=r"radiol|ecograf|imagen",
        claves=["radiolog", "radiology", "ecograf", "eco ", "ecoscop", "imagen",
                "imaginolog", "tac ", "resonancia", "radiograf", "pocus", "x-ray",
                "x ray", "chest x", "irm "]),
    "Análisis Clínicos": dict(
        carpeta=r"analisis|laborator",
        claves=["laboratorio", "hemograma", "gases arteriales", "gasometr",
                "quimica sanguinea", "valores normales", "liquido cefalorraq",
                "hepatograma", "balcells", "rangos de", "examen laboratorio",
                "gases "]),
    "Semiología y Exploración": dict(
        carpeta=r"semiol|exploraci", general=True,
        claves=["semiolog", "exploracion", "anamnesis", "argente", "suros", "bates",
                "propedeut", "fustinoni", "signos y sintomas", "signos vitales",
                "medicina clinica", "clinica practica"]),
    "Ciencias Básicas": dict(
        carpeta=r"basica|anatom|fisiol", general=True,
        claves=["anatomi", "fisiolog", "guyton", "netter", "histolog", "bioquim",
                "embriolog", "robbins", "fisiopatolog", "pathophysiology", "porth",
                "genetica", "biologia molecular", "inmunolog", "fisicoquim",
                "organic chemistry", "estructura y funcion", "berne", "harper",
                "mcphee"]),
    "Investigación y Estadística": dict(
        carpeta=r"investig|estadist",
        claves=["estadistic", "bioestad", "biostatistic", "investigacion",
                "epidemiolog", "metodolog", "medicina basada en la evidencia",
                "lectura critica"]),
    "MIR": dict(
        carpeta=r"\bmir\b",
        claves=["amir", "ctomir", "mir ", "promir", "manual cto", "desgloses"]),
}

# Carpetas que no se analizan como especialidad (copias de trabajo)
CARPETAS_ESPECIALES = re.compile(r"notebook\s*lm", re.IGNORECASE)

# Palabras que van en minúscula (salvo si son la primera palabra)
MINUSCULAS = {
    "el", "la", "los", "las", "lo", "un", "una", "unos", "unas",
    "de", "del", "al", "a", "y", "e", "o", "u", "en", "con", "por", "para",
    "sin", "sobre", "entre", "desde", "hasta", "según", "segun", "ante", "tras",
    "su", "sus", "et", "vs",
    "i", "amb", "per", "els", "les", "d", "l",          # catalán
    "the", "of", "and", "in", "for", "on", "to", "an", "at", "by", "with", "or",
}

# Siglas que se deben mantener en mayúsculas
SIGLAS = {
    "ecg", "ekg", "epoc", "vih", "sida", "dsm", "mir", "amir", "cto", "uci", "rcp",
    "orl", "esc", "aha", "acc", "ada", "gold", "gema", "gesepoc", "semergen",
    "semg", "ics", "pocus", "tac", "rm", "rx", "irm", "ets", "its", "hta", "dm",
    "irc", "erc", "iam", "sca", "tep", "tvp", "fa", "ic", "eii", "nice", "who",
    "oms", "cdc", "idsa", "aeped", "sego", "sen", "seen", "sep", "ser", "sec",
    "semes", "semicyuc", "atls", "acls", "bls", "svb", "sva", "ebm", "mbe", "phtls",
    "promir", "gpc", "clm", "soap", "dmae", "serv", "qrs", "toc", "sn", "gu", "abc",
    "rii", "tab", "lcr", "aetsa", "papps", "amf", "ap", "dx", "tx",
    "ii", "iii", "iv", "vi", "vii", "viii", "ix", "xi", "xii", "xiii", "xiv", "xv",
}

# Nombres con mayúsculas/minúsculas propias
FORMA_EXACTA = {
    "semfyc": "semFYC", "camfic": "CAMFiC", "mfyc": "MFyC", "uptodate": "UpToDate",
    "vexus": "VExUS", "dx": "Dx", "tx": "Tx", "dejong": "DeJong", "mcphee": "McPhee",
    "notebooklm": "NotebookLM", "rayospedia": "Rayospedia", "ekg": "EKG",
}

# Palabras frecuentes que suelen venir sin tilde (o mal escritas)
ACENTOS = {
    "guia": "guía", "guias": "guías", "diagnostico": "diagnóstico",
    "diagnostica": "diagnóstica", "diagnosticos": "diagnósticos",
    "estadistico": "estadístico", "terapeutica": "terapéutica",
    "terapeutico": "terapéutico", "clinica": "clínica", "clinicas": "clínicas",
    "clinico": "clínico", "clinicos": "clínicos", "medico": "médico",
    "medica": "médica", "medicos": "médicos", "medicas": "médicas",
    "practica": "práctica", "practicas": "prácticas", "practico": "práctico",
    "basica": "básica", "basicas": "básicas", "basico": "básico",
    "basicos": "básicos", "tecnicas": "técnicas", "tecnica": "técnica",
    "cirugia": "cirugía", "critico": "crítico", "critica": "crítica",
    "criticos": "críticos", "rapida": "rápida", "rapido": "rápido",
    "analisis": "análisis", "sintomas": "síntomas", "sindrome": "síndrome",
    "sindromes": "síndromes", "farmaco": "fármaco", "farmacos": "fármacos",
    "pediatrica": "pediátrica", "pediatricas": "pediátricas",
    "pediatria": "pediatría", "psiquiatria": "psiquiatría",
    "psiquiatrica": "psiquiátrica", "psiquiatricas": "psiquiátricas",
    "geriatria": "geriatría", "geriatrica": "geriátrica", "cardiaca": "cardíaca",
    "cardiaco": "cardíaco", "fisica": "física", "fisico": "físico",
    "metodo": "método", "metodos": "métodos", "pagina": "página",
    "ortopedica": "ortopédica", "via": "vía", "vias": "vías", "oseo": "óseo",
    "organos": "órganos", "organo": "órgano", "musculo": "músculo",
    "musculos": "músculos", "indice": "índice", "torax": "tórax",
    "higado": "hígado", "rinon": "riñón", "nino": "niño", "ninos": "niños",
    "espanol": "español", "espanola": "española", "neumonia": "neumonía",
    "imagenes": "imágenes", "mecanica": "mecánica", "mecanico": "mecánico",
    "quimica": "química", "bioquimica": "bioquímica",
    "fisicoquimica": "fisicoquímica", "diabetica": "diabética",
    "diabetico": "diabético", "lipidico": "lipídico", "macroscopica": "macroscópica",
    "microscopica": "microscópica", "patologico": "patológico",
    "sanguinea": "sanguínea", "liquido": "líquido", "liquidos": "líquidos",
    "capitulo": "capítulo", "anatomia": "anatomía", "neuroanatomia": "neuroanatomía",
    "propedeutica": "propedéutica", "reumatica": "reumática",
    "reumaticas": "reumáticas", "musculoesqueletica": "musculoesquelética",
    "musculoesqueleticas": "musculoesqueléticas", "matematicas": "matemáticas",
    "cronica": "crónica", "cronicas": "crónicas", "cronico": "crónico",
    "cronicos": "crónicos", "hemorragico": "hemorrágico", "genetica": "genética",
    "biologia": "biología", "electrolitos": "electrolitos", "acido": "ácido",
    "hematologicas": "hematológicas", "farmacoterapia": "farmacoterapia",
    "cateter": "catéter", "numero": "número", "ultimo": "último",
    "actualizacion": "actualización", "atencion": "atención",
    "cetoacidosis": "cetoacidosis", "mas": "más", "codigo": "código",
    "patologias": "patologías", "patologia": "patología", "sistematico": "sistemático",
    "oncologicas": "oncológicas", "psicoterapia": "psicoterapia",
    "rios": "Ríos", "alvarez": "Álvarez", "angel": "Ángel", "avendano": "Avendaño",
    "martinez": "Martínez", "gonzalez": "González", "rodriguez": "Rodríguez",
    "garcia": "García", "lopez": "López", "perez": "Pérez", "sanchez": "Sánchez",
    "gomez": "Gómez", "fernandez": "Fernández", "hernandez": "Hernández",
    "brujula": "brújula", "vertigo": "vértigo", "diaz": "Díaz", "jimenez": "Jiménez", "munoz": "Muñoz", "galan": "Galán",
    # erratas frecuentes
    "escencial": "esencial", "psiquiatrfcas": "psiquiátricas", "resúmen": "resumen",
    "espapl": "español",
}

# Basura típica en nombres de archivo descargados
BASURA = [
    r"\(?\bz-?lib(?:rary)?(?:\.org)?\b\)?", r"\(?\blibgen(?:\.\w+)?\)?", r"\bb-ok\b",
    r"\bwww\.[a-z0-9.-]+", r"\b[a-z0-9-]+(?:\.[a-z0-9-]+)*\.(?:com|net|org|es|info)\b",
    r"\bbooksmedicos\b", r"\bmedilibros\b", r"\bmedicomoderno\b", r"\bbmpdf\b",
    r"\bmedinternafacil\b", r"@\w+", r"\[medical force\]", r"\bmedical force\b",
    r"\bebook\b", r"\bpdf\b", r"\bepub\b", r"\bdescargar\b", r"\bgratis\b",
    r"\bcopia(?:r)?\b", r"\bcopy\b", r"\bdefinitivo\b", r"\bscan\b",
    r"\benglish\b", r"\bspanish\b", r"\bocr\b", r"\bmaquetado\b", r"\bfb\b\.?",
    r"\bmicrosoft (?:word|powerpoint|excel)\b\s*-?", r"\bcompl-\d\b", r"\bdialnet\b",
    r"\(\s*\d\s*\)", r"\[\s*\d\s*\]", r"۩+", r"(?<=\s)l(?=\s)",
    r"\bisbn[\s:-]*[\dxX-]{10,17}\b", r"\b97[89][\d-]{10,14}\b",
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
    "twentieth": 20,
}

# Palabras que no bastan para decir que dos títulos son el mismo libro
GENERICAS = {
    "manual", "guia", "guias", "tratado", "atlas", "principios", "protocolo",
    "libro", "curso", "tema", "temas", "apuntes", "resumen", "esquema", "esquemas",
    "infografia", "medicina", "introduccion", "fundamentos", "compendio", "clinica",
    "clinico", "medica", "medico", "practica", "practico", "urgencias",
    "emergencias", "interpretacion", "diagnostico", "tratamiento", "basico",
    "basica", "enfermedades", "enf", "errores", "comunes", "examenes", "laboratorio",
    "comprimido", "copia", "parte", "tomo", "texto", "bolsillo", "mapas", "mentales",
    "cardiologia", "semiologia", "farmacologia", "fisiologia", "radiologia",
    "humana", "general", "aplicada",
}

SUFIJO_ORD = r"(?:ª|º|°|a|o|er|ra|ro|da|do|ta|to|va|vo|na|no|ma|mo|th|st|nd|rd)"


# ---------------------------------------------------------------------------
# Utilidades de texto
# ---------------------------------------------------------------------------
def sin_acentos(s):
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def normalizar(s):
    """Minúsculas, sin acentos, solo letras/números y espacios."""
    s = sin_acentos(unicodedata.normalize("NFC", s).lower())
    s = re.sub(r"[^a-z0-9ñ ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def extraer_edicion(texto):
    """Devuelve (numero_edicion | None, texto_sin_edicion)."""
    t = texto
    patrones = [
        # 22ª Ed / 22a ed / 10ma Edición / 22th edition / 2daed / 14 edicion
        rf"\b(\d{{1,2}})\s*{SUFIJO_ORD}?\.?\s*(?:ed\b\.?|edic\w*\.?|edition\b|edn\b)",
        # Edición 22 / Edition 22 / ed. 22
        r"\b(?:edici[oó]n|edition|ed\.?)\s*(\d{1,2})\b",
        # "2e" (inglés)
        r"\b(\d{1,2})e\b",
        # "Manual Washington 5°" / "Gilberto Angel 7º"
        r"\b(\d{1,2})\s*(?:º|°|ª)(?!\s*\d)",
    ]
    for p in patrones:
        m = re.search(p, t, flags=re.IGNORECASE)
        if m and 0 < int(m.group(1)) < 60:
            return int(m.group(1)), (t[:m.start()] + " " + t[m.end():])
    m = re.search(r"\b([a-záéíóú-]+)\s+(?:edici[oó]n|edition|ed\b\.?)",
                  t, flags=re.IGNORECASE)
    if m:
        n = ORDINALES.get(sin_acentos(m.group(1).lower()))
        if n:
            return n, (t[:m.start()] + " " + t[m.end():])
    return None, t


def limpiar_basura(texto):
    """Quita restos de webs, guiones bajos, etc. Devuelve (texto, notas)."""
    notas = []
    t = unicodedata.normalize("NFC", texto)
    # nombres tipo "mapas-mentales-de-medicamentos": guiones = espacios
    if t.count("-") >= 2 and t.count("-") > t.count(" "):
        t = t.replace("-", " ")
    t = re.sub(r"(?<!\d)(\d)_(\d)(?!\d)", r"\1.\2", t)          # GEMA_5_0 -> 5.0
    t = t.replace("_", " ").replace("+", " ")
    if re.search(r"\d{5,7}[-\s]downloadable[-\s]\d+", t, re.IGNORECASE):
        t = re.sub(r"[-\s]*\d{5,7}[-\s]downloadable[-\s]\d+", " ", t, flags=re.IGNORECASE)
        notas.append("título cortado por la web de descarga: completar a mano")
    if "..." in t or "…" in t:
        notas.append("título cortado (...): completar a mano")
        t = t.replace("...", " ").replace("…", " ")
    for p in BASURA:
        t = re.sub(p, " ", t, flags=re.IGNORECASE)
    t = re.sub(r"(?<=[^\W\d])\.(?=[^\W\d])", " ", t)            # puntos entre palabras
    t = re.sub(r"(?<=[^\W\d])\.(?=\d)|(?<=\d)\.(?=[^\W\d])", " ", t)
    t = re.sub(r"\.{2,}$|\.\s*$", " ", t)
    t = re.sub(r"--+", " ", t)
    # separar "Harrison22ed", "GuiaEpilepsia", "topTenPROMIR", "Adquiri2022"
    t = re.sub(r"(?<=[a-zA-Z])(?=\d{1,2}\s*(?:ª|º)?\s*ed)", " ", t, flags=re.IGNORECASE)
    t = re.sub(r"(?<=[a-záéíóúñ])(?=(?:19|20)\d\d\b)", " ", t)
    t = re.sub(r"\S+", _separar_pegadas, t)
    t = re.sub(r"^resumen(?=[a-z])", "resumen ", t, flags=re.I)  # "resumenepoc"
    t = re.sub(r"\b\d{5,}\b", " ", t)                          # números de descarga
    t = re.sub(r"[\[\]{}]", " ", t)
    t = re.sub(r"\(\s*\)", " ", t)
    for p in BASURA:
        t = re.sub(p, " ", t, flags=re.IGNORECASE)
    t = re.sub(r"\s+", " ", t).strip(" -.,;")
    return t, notas


def _separar_pegadas(m):
    """'GuiaBasicaSobreLa' -> 'Guia Basica Sobre La'; respeta DeJong, McPhee, VExUS."""
    w = m.group(0)
    if len(w) < 8 or normalizar(w) in FORMA_EXACTA or re.match(r"Ma?c[A-Z]", w):
        return w
    w = re.sub(r"(?<=[a-záéíóúñ])(?=[A-ZÁÉÍÓÚÑ])", " ", w)
    return re.sub(r"(?<=[A-Z])(?=[A-Z][a-z]{2})", " ", w)


def acentuar(palabra):
    """Añade la tilde a palabras españolas comunes que vienen sin ella."""
    low = palabra.lower()
    if low in ACENTOS:
        return ACENTOS[low]
    if low != sin_acentos(low):  # ya lleva tilde
        return low
    reglas = (("cion", "ción"), ("logias", "logías"), ("logia", "logía"),
              ("grafias", "grafías"), ("grafia", "grafía"), ("iatria", "iatría"),
              ("logicas", "lógicas"), ("logicos", "lógicos"), ("logica", "lógica"),
              ("logico", "lógico"), ("patia", "patía"), ("patias", "patías"),
              ("metria", "metría"))
    for suf, rep in reglas:
        if low.endswith(suf) and len(low) > len(suf) + 1:
            return low[: -len(suf)] + rep
    return low


def _formatear_palabra(nucleo, primera, todo_mayus):
    base = normalizar(nucleo)
    if not nucleo:
        return nucleo
    if "-" in nucleo and len(nucleo) > 1:
        partes = nucleo.split("-")
        return "-".join(_formatear_palabra(p, primera and i == 0, todo_mayus)
                        for i, p in enumerate(partes))
    if base in FORMA_EXACTA:
        return FORMA_EXACTA[base]
    if base in SIGLAS:
        return nucleo.upper()
    if not todo_mayus:
        if re.fullmatch(r"[A-Z]{2,4}\d*", nucleo):
            return nucleo                                   # sigla corta: TOC, RII
        if re.search(r"[a-záéíóúñ][A-ZÁÉÍÓÚÑ]", nucleo) and not nucleo.isupper():
            return nucleo                                   # VExUS, DeJong, iPhone
    if any(c.isdigit() for c in nucleo):
        return nucleo.lower()
    if not primera and base in MINUSCULAS:
        return nucleo.lower()
    a = acentuar(nucleo)
    if a.startswith("mc") and len(a) > 3:
        return "Mc" + a[2:].capitalize()
    return a[:1].upper() + a[1:]


def formato_titulo(texto):
    letras = [c for c in texto if c.isalpha()]
    todo_mayus = bool(letras) and sum(c.isupper() for c in letras) / len(letras) > 0.6
    if todo_mayus:
        texto = texto.lower()  # venía TODO EN MAYÚSCULAS
    salida = []
    palabras = texto.split()
    for i, w in enumerate(palabras):
        m = re.match(r"^([\W_]*)(.*?)([\W_]*)$", w)
        pre, nucleo, post = m.groups()
        if post and nucleo and post[0] == "-":
            nucleo, post = nucleo + post, ""
        primera = i == 0 or (salida and salida[-1].endswith(("-", ":", ".")))
        salida.append(pre + _formatear_palabra(nucleo, primera, todo_mayus) + post)
    # quitar palabras repetidas seguidas ("2025 2025")
    limpio = []
    for w in salida:
        if not limpio or normalizar(limpio[-1]) != normalizar(w) or not normalizar(w):
            limpio.append(w)
    return " ".join(limpio)


def detectar_tipo(nombre, ext, carpeta=""):
    n = normalizar(nombre)
    c = normalizar(carpeta)
    if ext in EXT_IMAGEN or "infografi" in n or "infografi" in c or "poster" in n:
        return "Infografía"
    if "articul" in c:
        return "Artículo"
    if re.search(r"\b(guia|guias|guideline|guidelines|consenso|protocolo|algoritmo|"
                 r"recomendaciones|gpc|codigo)\b", n):
        return "Guía"
    return "Libro"


def titulo_metadatos(ruta):
    """Título interno del PDF/epub vía Spotlight (solo macOS). Vacío si no hay."""
    if sys.platform != "darwin":
        return ""
    try:
        r = subprocess.run(["mdls", "-raw", "-name", "kMDItemTitle", str(ruta)],
                           capture_output=True, text=True, timeout=10)
        t = r.stdout.strip()
        t = re.sub(r"^microsoft (word|powerpoint|excel)\s*-\s*", "", t, flags=re.I)
        if t and t != "(null)" and len(t) > 4:
            return t
    except Exception:
        pass
    return ""


def nombre_parece_basura(nombre):
    n = normalizar(nombre)
    letras = re.sub(r"[^a-z]", "", n)
    return (len(letras) < 6 or re.fullmatch(r"[a-f0-9 ]{12,}", n) is not None
            or re.fullmatch(r"[\d ]+", n) is not None
            or re.match(r"^\d+ \d{8,}$", n) is not None
            or re.match(r"^(scan|doc|documento|img|image|file|download|untitled|"
                        r"sin titulo|whatsapp|captura|screenshot)\b", n) is not None
            or re.fullmatch(r"(19|20)\d\d \d\d \d\d.*", n) is not None)


ANIO = r"\(?\b(?:19|20)\d{2}\b\)?"


def proponer_nombre(ruta, en_serie=False):
    """Devuelve dict con nombre, clave (set de palabras), edicion, tipo, notas."""
    ext = ruta.suffix.lower()
    original = ruta.stem
    notas = []
    fuente = original
    if nombre_parece_basura(original):
        meta = titulo_metadatos(ruta)
        if meta and not nombre_parece_basura(meta):
            fuente = meta
            notas.append("título sacado de los metadatos del PDF")
        else:
            if ext in EXT_IMAGEN:
                notas.append("imagen sin nombre: mira qué es y renómbrala")
            else:
                notas.append("nombre poco claro: revisar a mano")
            nombre = original + ext
            return dict(nombre=nombre, clave=frozenset(), edicion=None,
                        tipo=detectar_tipo(original, ext, ruta.parent.name),
                        notas=notas, titulo_base="")

    zlib = re.search(r"z-?lib", fuente, re.IGNORECASE) is not None
    comprimido = re.search(r"compress(ed)?(?![a-z])|comprimido", fuente, re.I) is not None
    fuente = re.sub(r"[_\s.-]*\(?(compress(ed)?|comprimido)\)?(?![a-z])", " ", fuente,
                    flags=re.I)
    limpio, n2 = limpiar_basura(fuente)
    notas += n2
    ed, resto = extraer_edicion(limpio)
    if ed is None and fuente != original:
        ed, _ = extraer_edicion(limpiar_basura(original)[0])
    resto = re.sub(r"\s+", " ", resto).strip(" -.,;")
    parte = ""
    mp = re.search(r"[,.\s-]*\b(parte|part|tomo|vol(?:umen)?)\.?\s*(\d+|[ivx]+)\b"
                   r"(?:\s*de\s*(\d+))?", resto, re.IGNORECASE)
    if mp:
        num = mp.group(2)
        num = str(int(num)) if num.isdigit() else num.upper()
        palabra = {"part": "Parte", "volumen": "Vol", "vol": "Vol"}.get(
            mp.group(1).lower(), mp.group(1).capitalize())
        parte = f"{palabra} {num}" + (f" de {mp.group(3)}" if mp.group(3) else "")
        resto = (resto[:mp.start()] + " " + resto[mp.end():]).strip(" -.,;")

    # Separar título y autor (tu formato: "Título 7ª Ed - Autor")
    titulo, autor = resto, ""
    m = re.match(r"^(.*?\S)\s*-\s+(.*)$", resto) if zlib else re.match(
        r"^(.*?\S)\s+-\s+(.*)$", resto)
    if m:
        a, b = m.group(1).strip(), m.group(2).strip()
        if re.fullmatch(ANIO, a):                    # "2015 - Manual de ..."
            titulo, autor = f"{b} {a}", ""
        elif zlib:                                   # Z-Library: "Autor - Título"
            titulo, autor = b, a
        elif len(a.split()) <= 2 and not puntuar(a) and puntuar(b):
            titulo, autor = b, a                     # "West - Fisiología Respiratoria"
        else:
            titulo, autor = a, b
    tipo = detectar_tipo(fuente, ext, ruta.parent.name)

    if ed and tipo == "Libro":
        m_p = re.search(ANIO + r"\s+(\d)$", titulo)             # "2026 1" -> Parte 1
        if m_p and not parte:
            parte = f"Parte {m_p.group(1)}"
            titulo = titulo[:m_p.start()]
        titulo = re.sub(r"\s*" + ANIO, " ", titulo)
        autor = re.sub(r"\s*" + ANIO, " ", autor)
    if not en_serie:
        titulo = re.sub(r"^\d{1,2}(?:[)\s.-]+(?=[^\W\d])|(?=[A-Z][a-z]))", "", titulo)
    titulo = re.sub(r"(?<=\w)- (?=\w)", " - ", titulo)
    titulo = re.sub(r"\s+,", ",", re.sub(r"\s+", " ", titulo)).strip(" -.,;")
    autor = re.sub(r"\s+,", ",", re.sub(r"\s+", " ", autor)).strip(" -.,;")
    if re.fullmatch(r"[\d\s]*", autor):
        autor = ""

    nombre = formato_titulo(titulo) or formato_titulo(original)
    if ed:
        nombre += f" {ed}ª Ed"
    if parte:
        nombre += f" - {parte}"
    if autor:
        nombre += f" - {formato_titulo(autor)}"
    if comprimido:
        nombre += " (comprimido)"
    nombre = re.sub(r'[/:\\?*"<>|]', "-", nombre).strip()

    if re.match(r"^[A-ZÁÉÍÓÚ][a-záéíóú]+ [A-Z]{1,2} ?[A-ZÁÉÍÓÚ][a-záéíóú]+,", nombre):
        notas.append("parece que empieza por el nombre de los autores")
    return dict(nombre=nombre + ext, clave=clave_titulo(titulo), edicion=ed,
                tipo=tipo, notas=notas, titulo_base=titulo)


def clave_titulo(titulo):
    palabras = normalizar(titulo).split()
    return frozenset(p for p in palabras
                     if p not in MINUSCULAS and not p.isdigit() and len(p) > 1
                     and p not in {"ed", "edicion", "comprimido", "copia"})


# ---------------------------------------------------------------------------
# Clasificación por especialidad
# ---------------------------------------------------------------------------
_PATRONES = {
    esp: [re.compile(r"\b" + re.escape(normalizar(c)) + (r"\b" if c.endswith(" ") else ""))
          for c in d["claves"]]
    for esp, d in ESPECIALIDADES.items()
}


def puntuar(texto):
    n = normalizar(texto)
    puntos = {}
    for esp, patrones in _PATRONES.items():
        p = sum(1 for pat in patrones if pat.search(n))
        if p:
            puntos[esp] = p
    return puntos


def mejor_especialidad(texto):
    p = puntuar(texto)
    if not p:
        return None, 0
    esp = max(p, key=lambda k: (p[k], not ESPECIALIDADES[k].get("general", False)))
    return esp, p[esp]


def numero_carpeta(nombre):
    m = re.match(r"^\s*(\d+)\)", nombre)
    return int(m.group(1)) if m else None


def carpetas_existentes(biblioteca, origen):
    """Mapa especialidad -> carpeta ya existente en la biblioteca."""
    mapa = {}
    existentes = [d for d in biblioteca.iterdir()
                  if d.is_dir() and d.resolve() != origen.resolve()
                  and not d.name.startswith((".", "_"))
                  and not CARPETAS_ESPECIALES.search(d.name)]
    for esp, datos in ESPECIALIDADES.items():
        patron = re.compile(r"\b(?:" + datos["carpeta"] + ")")
        candidatas = [d for d in existentes
                      if patron.search(normalizar(re.sub(r"^\s*\d+\)\s*", "", d.name)))]
        if candidatas:
            mapa[esp] = sorted(candidatas, key=lambda d: numero_carpeta(d.name) or 999)[0]
    return mapa, existentes


def especialidad_de_carpeta(carpeta, mapa):
    for esp, d in mapa.items():
        if d.resolve() == carpeta.resolve():
            return esp
    return None


def nombre_carpeta_nueva(esp, numero):
    nombre = esp.upper().replace(" Y ", " y ")
    return f"{numero}) {nombre}"


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


ES_PARTE = re.compile(r"\bparte?\s*\d|\btomo\s*\w+|\(\d+ de \d+|\bvol(umen)?\.?\s*\d|"
                      r"\bharrison-\d{4}", re.IGNORECASE)


def detectar_series(archivos):
    """Capítulos numerados 00, 01, 02... en la misma carpeta -> se mantienen juntos."""
    series = {}
    por_carpeta = defaultdict(list)
    for p in archivos:
        m = re.match(r"^(\d{2})[\s_.-]", p.name)
        if m:
            por_carpeta[p.parent].append((int(m.group(1)), p))
    for carpeta, lista in por_carpeta.items():
        nums = sorted({n for n, _ in lista})
        if not nums or nums[0] > 1:
            continue
        seguidos = [nums[0]]
        saltos = 0
        for n in nums[1:]:
            if n - seguidos[-1] == 1 or (n - seguidos[-1] == 2 and saltos == 0):
                saltos += n - seguidos[-1] - 1
                seguidos.append(n)
            else:
                break
        if len(seguidos) < 4:
            continue
        miembros = sorted((n, p) for n, p in lista if n in seguidos)
        cabeza = miembros[0][1]
        repetidos = {n for n in seguidos if sum(1 for m, _ in miembros if m == n) > 1}
        for n, p in miembros:
            series[p] = (cabeza, n in repetidos)
    return series


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
    print(f"Carpetas de especialidad reconocidas: {len(mapa)}")
    for esp, d in sorted(mapa.items(), key=lambda x: numero_carpeta(x[1].name) or 999):
        print(f"   {unicodedata.normalize('NFC', d.name):35s} <- {esp}")
    sin_uso = [d.name for d in existentes if d not in mapa.values()]
    if sin_uso:
        print("Carpetas que no he asociado a ninguna especialidad: " + ", ".join(sin_uso))

    todos = list(buscar_archivos(biblioteca))
    set_origen = {p for p in todos if origen.resolve() in p.resolve().parents}
    en_origen = sorted(set_origen)
    ya_clasificados = [p for p in todos if p not in set_origen
                       and CARPETA_DUPLICADOS not in p.parts
                       and Path(args.plan).name != p.name]
    print(f"Archivos en {origen.name}: {len(en_origen)}  |  "
          f"ya clasificados: {len(ya_clasificados)}")

    cache = {}

    def info(p, en_serie=False):
        if (p, en_serie) not in cache:
            cache[(p, en_serie)] = proponer_nombre(p, en_serie)
        return cache[(p, en_serie)]

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
        if len(lista) > 1 and tam > 0:
            for p in lista:
                try:
                    grupos_hash[hash_archivo(p)].append(p)
                except OSError:
                    pass
    dup_exactos = [g for g in grupos_hash.values() if len(g) > 1]

    # Qué copia se queda: la ya clasificada; si no, la de nombre más completo
    descartar = {}
    for g in dup_exactos:
        def calidad(p):
            i = info(p)
            top = biblioteca / p.relative_to(biblioteca).parts[0]
            actual = especialidad_de_carpeta(top, mapa)
            sugerida, _ = mejor_especialidad(p.stem)
            mal_sitio = bool(actual and sugerida and actual != sugerida
                             and not ESPECIALIDADES[actual].get("general"))
            return (p in set_origen, bool(CARPETAS_ESPECIALES.search(str(p))),
                    re.search(r"\bcopia\b", p.stem, re.I) is not None, mal_sitio,
                    nombre_parece_basura(p.stem), i["edicion"] is None, -len(i["clave"]))
        g_orden = sorted(g, key=calidad)
        for p in g_orden[1:]:
            descartar[p] = g_orden[0]

    # --- Carpetas nuevas: siguen tu numeración "N) NOMBRE" ---
    siguiente = max([numero_carpeta(d.name) or 0 for d in existentes] + [0]) + 1
    nuevas = {}

    def destino_para(esp):
        nonlocal siguiente
        if esp is None:
            return biblioteca / CARPETA_SIN_CLASIFICAR
        if esp in mapa:
            return mapa[esp]
        if esp not in nuevas:
            nuevas[esp] = biblioteca / nombre_carpeta_nueva(esp, siguiente)
            siguiente += 1
        return nuevas[esp]

    series = detectar_series(en_origen)
    esp_serie = {}
    for p, (cabeza, _) in series.items():
        if cabeza not in esp_serie:
            miembros = [q for q, (c, _) in series.items() if c == cabeza]
            esp, conf = mejor_especialidad(cabeza.stem)
            if esp is None:
                esp, conf = mejor_especialidad(" ".join(q.stem for q in miembros))
            esp_serie[cabeza] = esp

    # --- Plan para los archivos del origen ---
    filas = []
    usados = defaultdict(set)
    for p in en_origen:
        en_serie = p in series
        i = info(p, en_serie)
        nuevo, notas = i["nombre"], list(i["notas"])
        if en_serie:
            cabeza, dudoso = series[p]
            esp = esp_serie[cabeza]
            conf = 1 if esp else 0
            nombre_serie = info(cabeza, True)["nombre"]
            nombre_serie = re.sub(r"^\d{2}\s+", "", os.path.splitext(nombre_serie)[0])
            carpeta = destino_para(esp) / nombre_serie
            notas.append(f"capítulo de la serie '{nombre_serie}'"
                         + (" (¿seguro que es de la serie? hay dos con el mismo número)"
                            if dudoso else ""))
        else:
            texto_clasif = f"{p.stem} {nuevo}"
            esp, conf = mejor_especialidad(texto_clasif)
            if esp is None:
                meta = titulo_metadatos(p)
                if meta:
                    esp, conf = mejor_especialidad(meta)
            carpeta = destino_para(esp)
            if esp is None and i["tipo"] == "Infografía":
                carpeta = carpeta / "Infografías"
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
        filas.append({
            "accion": accion,
            "ruta_actual": str(p.relative_to(biblioteca)),
            "nombre_nuevo": nuevo,
            "carpeta_destino": str(carpeta.relative_to(biblioteca)),
            "tipo": i["tipo"],
            "especialidad": esp or "?",
            "confianza": "baja" if conf == 0 else ("media" if conf == 1 else "alta"),
            "notas": "; ".join(notas),
        })

    # --- Posibles duplicados / varias ediciones ---
    # Se comparan las palabras del título: si las de uno están todas en el otro
    # (y al menos 2 son significativas) se avisa. Se ignoran partes/tomos y NotebookLM.
    items = []
    for p in en_origen + ya_clasificados:
        if p in descartar or CARPETAS_ESPECIALES.search(str(p.relative_to(biblioteca))):
            continue
        if ES_PARTE.search(p.stem) or p.suffix.lower() in EXT_IMAGEN or p in series:
            continue
        i = info(p)
        if i["clave"]:
            items.append((p, i))
    padre = list(range(len(items)))

    def raiz(k):
        while padre[k] != k:
            padre[k] = padre[padre[k]]
            k = padre[k]
        return k

    def mismo_libro(a, b, na, nb):
        cifras = lambda n: re.findall(r"\d+", re.sub(r"^\d+\s", "", n))
        if a == b and cifras(na) != cifras(nb):
            return False                      # "Farma Infecciosas 2" / "... 3": serie
        corto, largo = sorted((a, b), key=len)
        if not corto <= largo:
            return False
        return len(corto - GENERICAS) >= 2 or (corto == largo and len(corto) >= 2)

    for x in range(len(items)):
        for y in range(x + 1, len(items)):
            if mismo_libro(items[x][1]["clave"], items[y][1]["clave"],
                           normalizar(items[x][0].stem), normalizar(items[y][0].stem)):
                padre[raiz(x)] = raiz(y)
    grupos = defaultdict(list)
    for k, it in enumerate(items):
        grupos[raiz(k)].append(it)
    posibles = [g for g in grupos.values() if len(g) > 1]

    # --- Mal clasificados (libros ya en carpetas de especialidad) ---
    mal = []
    for p in ya_clasificados:
        rel = p.relative_to(biblioteca)
        carpeta_top = biblioteca / rel.parts[0]
        actual = especialidad_de_carpeta(carpeta_top, mapa)
        if not actual or ESPECIALIDADES[actual].get("general"):
            continue
        texto_ruta = " ".join(rel.parts[1:])  # incluye subcarpetas
        sugerida, conf = mejor_especialidad(p.stem)
        if sugerida and sugerida != actual and puntuar(texto_ruta).get(actual, 0) == 0 \
                and not ESPECIALIDADES[sugerida].get("general"):
            mal.append((p, carpeta_top.name, sugerida))

    # --- Guardar CSV ---
    salida = Path(args.plan)
    campos = ["accion", "ruta_actual", "nombre_nuevo", "carpeta_destino",
              "tipo", "especialidad", "confianza", "notas"]
    with open(salida, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=campos, delimiter=";")
        w.writeheader()
        w.writerows(filas)
    with open(salida.with_suffix(".json"), "w", encoding="utf-8") as f:
        json.dump({"biblioteca": str(biblioteca), "origen": str(origen)}, f)

    # --- Informe ---
    rel = lambda p: str(p.relative_to(biblioteca))
    L = [f"INFORME DE ORGANIZACIÓN — {datetime.now():%d/%m/%Y %H:%M}", "=" * 70, ""]
    L.append(f"Archivos a organizar: {len(filas)}  "
             f"(sin clasificar: {sum(1 for r in filas if r['especialidad'] == '?')})")
    L.append("")
    if nuevas:
        L.append("CARPETAS NUEVAS QUE SE CREARÁN (cámbialas en el CSV si prefieres otras):")
        for esp, d in nuevas.items():
            n = sum(1 for r in filas if r["carpeta_destino"].startswith(d.name))
            L.append(f"   {d.name}   ({n} archivos)")
        L.append("")
    L.append(f"1) DUPLICADOS EXACTOS (mismo archivo): {len(dup_exactos)} grupos")
    L.append("-" * 70)
    L.append("   Solo se mueven las copias que están en LIBROSSSSS. Si las dos copias ya")
    L.append("   están en tus carpetas no se toca nada: borra tú la que sobre (la marcada ✗).")
    for g in dup_exactos:
        for p in g:
            if p not in descartar:
                marca = "  [se queda]"
            elif p in set_origen:
                marca = "  [→ _Duplicados]"
            else:
                marca = "  [✗ sobra: no se toca, decide tú]"
            L.append(f"   {rel(p)}{marca}")
        L.append("")
    L.append(f"2) POSIBLES DUPLICADOS / VARIAS EDICIONES: {len(posibles)} grupos")
    L.append("   (mismo título; no son archivos idénticos. Se omiten partes, tomos"
             " y copias de NotebookLM)")
    L.append("-" * 70)
    for g in sorted(posibles, key=lambda g: info(g[0][0])["nombre"]):
        eds = [i["edicion"] for _, i in g if i["edicion"]]
        for p, i in sorted(g, key=lambda x: -(x[1]["edicion"] or 0)):
            extra = ""
            if eds and i["edicion"] and i["edicion"] < max(eds):
                extra = f"   ← edición anterior (tienes la {max(eds)}ª)"
            L.append(f"   {rel(p)}{extra}")
        L.append("")
    L.append(f"3) POSIBLEMENTE MAL CLASIFICADOS: {len(mal)}")
    L.append("-" * 70)
    for p, actual, sug in mal:
        L.append(f"   {rel(p)}")
        L.append(f"        está en: {actual}   →   parece: {sug}")
    L.append("")
    L.append("4) SIN CLASIFICAR O CON NOMBRE A REVISAR")
    L.append("-" * 70)
    for r in filas:
        if r["especialidad"] == "?" or "revisar" in r["notas"] or "mano" in r["notas"] \
                or "autores" in r["notas"]:
            nota = f"   [{r['notas']}]" if r["notas"] else ""
            L.append(f"   {r['ruta_actual']}\n        →  {r['nombre_nuevo']}{nota}")
    Path(args.informe).write_text("\n".join(L), encoding="utf-8")

    print()
    print(f"✔ Plan guardado en:    {salida.resolve()}")
    print(f"✔ Informe guardado en: {Path(args.informe).resolve()}")
    print(f"   Duplicados exactos: {len(dup_exactos)} grupos | posibles duplicados: "
          f"{len(posibles)} | posibles mal clasificados: {len(mal)} | "
          f"sin clasificar: {sum(1 for r in filas if r['especialidad'] == '?')}")
    print("Revisa y edita el CSV (puedes cambiar 'nombre_nuevo', 'carpeta_destino' o poner"
          " 'IGNORAR' en 'accion'). Luego ejecuta:  python3 organizar_libros.py aplicar")


def aplicar(args):
    plan = Path(args.plan)
    if not plan.exists():
        sys.exit("No existe el plan. Ejecuta primero: python3 organizar_libros.py escanear")
    meta = json.loads(plan.with_suffix(".json").read_text(encoding="utf-8"))
    biblioteca = Path(meta["biblioteca"])
    with open(plan, encoding="utf-8-sig") as f:
        muestra = f.read(4096)
        f.seek(0)
        delim = ";" if muestra.count(";") >= muestra.count(",") else ","
        filas = list(csv.DictReader(f, delimiter=delim))

    registro = []
    creadas = []
    hechos = errores = 0
    for r in filas:
        accion = (r.get("accion") or "").strip().upper()
        if accion not in ("MOVER", "DUPLICADO"):
            continue
        src = biblioteca / r["ruta_actual"]
        dst_dir = biblioteca / r["carpeta_destino"].strip()
        dst = dst_dir / r["nombre_nuevo"].strip()
        if not src.exists():
            print(f"  ✗ ya no existe: {src.name}")
            errores += 1
            continue
        if src.resolve() == dst.resolve():
            continue
        faltan = [d for d in [dst_dir, *dst_dir.parents] if not d.exists()]
        dst_dir.mkdir(parents=True, exist_ok=True)
        creadas += [str(d) for d in faltan]
        base, ext = os.path.splitext(dst.name)
        if not ext:                       # si al editar se borró la extensión
            ext = src.suffix
            dst = dst_dir / (base + ext)
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
    log.write_text(json.dumps({"movimientos": registro, "carpetas_creadas": creadas},
                              ensure_ascii=False, indent=1), encoding="utf-8")

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
    datos = json.loads(log.read_text(encoding="utf-8"))
    registro = datos["movimientos"] if isinstance(datos, dict) else datos
    creadas = datos.get("carpetas_creadas", []) if isinstance(datos, dict) else []
    n = 0
    for mov in reversed(registro):
        a, de = Path(mov["a"]), Path(mov["de"])
        if a.exists() and not de.exists():
            de.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(a), str(de))
            n += 1
    for d in sorted(creadas, key=len, reverse=True):   # quitar carpetas que creó y quedan vacías
        d = Path(d)
        try:
            if d.exists() and not [x for x in d.iterdir() if x.name != ".DS_Store"]:
                if (d / ".DS_Store").exists():
                    (d / ".DS_Store").unlink()
                d.rmdir()
        except OSError:
            pass
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
