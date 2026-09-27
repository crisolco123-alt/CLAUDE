# Organizador de Libros de Medicina

Programa para ordenar la carpeta `9) Llibres/LIBROSSSSS`:

- **Renombra** con el formato `Harrison Principios de Medicina Interna 22ª Ed`
  (mayúscula inicial excepto artículos y preposiciones, siglas como EPOC o DSM-5 en
  mayúsculas, tildes corregidas, basura tipo `(z-lib.org)`, `_`, `copia`, ISBN… eliminada).
- **Clasifica** por especialidad y reutiliza tus carpetas (`1) Cardiologia`…).
- **Detecta duplicados**: copias idénticas (mismo contenido aunque tengan otro nombre)
  y el mismo libro en varias ediciones.
- **Avisa** de los libros que parecen estar en la carpeta equivocada.
- **Nunca borra nada**: los duplicados van a `_Duplicados (revisar)` y todo se puede deshacer.

## Cómo usarlo (en el Mac)

1. Descarga `organizar_libros.py` en tu carpeta `9) Llibres`.
2. Abre **Terminal**, escribe `cd ` (con un espacio), arrastra la carpeta `9) Llibres`
   a la ventana y pulsa Intro.
3. Escanea (no mueve nada):
   ```
   python3 organizar_libros.py escanear
   ```
   (Si el Mac pide instalar las “herramientas de desarrollo”, acepta: incluyen Python.)
4. Revisa los dos archivos que aparecen:
   - `informe.txt`: duplicados, ediciones repetidas, libros mal clasificados y los que no se
     han podido clasificar.
   - `plan_organizacion.csv`: ábrelo con Numbers o Excel. Puedes cambiar `nombre_nuevo`,
     `carpeta_destino`, o poner `IGNORAR` en `accion` para dejar un archivo como está.
5. Aplica el plan:
   ```
   python3 organizar_libros.py aplicar
   ```
6. Si algo no te gusta:
   ```
   python3 organizar_libros.py deshacer
   ```

## Personalizar

Al principio del script están el diccionario `ESPECIALIDADES` (palabras clave de cada
especialidad), `SIGLAS`, `ACENTOS` y `FORMA_EXACTA` (p. ej. `semFYC`). Añade lo que
necesites.
