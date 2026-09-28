#!/usr/bin/env python3
"""
Renumera las carpetas de "9) Llibres".

  python3 renumerar_carpetas.py            -> solo muestra los cambios (no toca nada)
  python3 renumerar_carpetas.py aplicar    -> renombra las carpetas
  python3 renumerar_carpetas.py deshacer   -> vuelve a los nombres anteriores

Las carpetas de FIJAS se quedan tal cual y en ese orden. El resto de carpetas
de especialidad se ordenan alfabéticamente y se numeran a continuación (5, 6, 7...).
Las carpetas LIBROSSSSS, NotebookLM y las que empiezan por "_" no se tocan.
"""
import json
import re
import sys
import unicodedata
from pathlib import Path

BIBLIOTECA = Path("/Users/cristinaolivercolin/Documents/2) RESIDENCIA MFyC/"
                  "4) Cursos i Formació/9) Llibres")

FIJAS = [
    "1) URGENCIAS",
    "1b) MEDICINA INTERNA",
    "2) ANÁLISIS CLÍNICOS",
    "2a) ECG",
    "3) RADIOLOGÍA",
    "3a) ANAMNESIS",
    "3b) SEMIOLOGÍA y EF",
    "4) FISIOPATOLOGÍA",
]
PRIMER_NUMERO = 5
NO_TOCAR = re.compile(r"^(_|\.)|librossss|notebook\s*lm", re.IGNORECASE)
REGISTRO = BIBLIOTECA / "deshacer_renumerar.json"


def nfc(s):
    return unicodedata.normalize("NFC", s)


def sin_numero(nombre):
    return re.sub(r"^\s*\d+[a-z]?\)\s*", "", nfc(nombre))


def clave_orden(nombre):
    s = unicodedata.normalize("NFD", sin_numero(nombre).lower())
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def plan():
    carpetas = [d for d in BIBLIOTECA.iterdir() if d.is_dir()]
    fijas = {nfc(n) for n in FIJAS}
    for n in FIJAS:
        if not any(nfc(d.name) == nfc(n) for d in carpetas):
            print(f"  (aviso) no encuentro la carpeta fija: {n}")
    resto = [d for d in carpetas
             if nfc(d.name) not in fijas and not NO_TOCAR.search(nfc(d.name))]
    resto.sort(key=lambda d: clave_orden(d.name))
    cambios = []
    for i, d in enumerate(resto, start=PRIMER_NUMERO):
        nuevo = f"{i}) {sin_numero(d.name)}"
        if nfc(d.name) != nuevo:
            cambios.append((d, BIBLIOTECA / nuevo))
    return cambios, resto


def main():
    modo = sys.argv[1] if len(sys.argv) > 1 else "ver"
    if not BIBLIOTECA.is_dir():
        sys.exit(f"No encuentro la carpeta: {BIBLIOTECA}")

    if modo == "deshacer":
        if not REGISTRO.exists():
            sys.exit("No hay nada que deshacer.")
        movs = json.loads(REGISTRO.read_text(encoding="utf-8"))
        tmp = [(Path(a), Path(a).with_name(Path(a).name + " __tmp")) for de, a in movs]
        for a, t in tmp:
            a.rename(t)
        for (de, a), (_, t) in zip(movs, tmp):
            t.rename(de)
        REGISTRO.unlink()
        print(f"✔ {len(movs)} carpetas devueltas a su nombre anterior.")
        return

    cambios, resto = plan()
    print("Orden final:")
    for n in FIJAS:
        print(f"   {n}")
    destino = {d: n for d, n in cambios}
    for d in resto:
        print(f"   {nfc(destino.get(d, d).name)}"
              + (f"      (antes: {nfc(d.name)})" if d in destino else ""))

    if modo != "aplicar":
        print(f"\n{len(cambios)} carpetas cambiarían de número. Para hacerlo de verdad:")
        print("   python3 renumerar_carpetas.py aplicar")
        return

    # En dos pasos para que ningún nombre choque con otro
    temporales = []
    for d, n in cambios:
        t = d.with_name(d.name + " __tmp")
        d.rename(t)
        temporales.append((t, n, d))
    for t, n, d in temporales:
        t.rename(n)
    REGISTRO.write_text(json.dumps([[str(d), str(n)] for d, n in cambios],
                                   ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\n✔ {len(cambios)} carpetas renumeradas. "
          "Para deshacer: python3 renumerar_carpetas.py deshacer")


if __name__ == "__main__":
    main()
