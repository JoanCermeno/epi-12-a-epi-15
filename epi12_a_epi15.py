#!/usr/bin/env python3
"""
epi12_a_epi15.py
================
Llena el EPI-15 mensual de los CPT que NO entregaron su reporte, sumando las
semanas del EPI-12 (4 o 5 archivos) y escribiendo el resultado en la columna P
de la hoja de cada CPT.

Uso (ejemplo):
    python epi12_a_epi15.py "EPI-15 NUEVO OCTUBRE.xlsx" carpeta_epi12/ --mes OCTUBRE

  - Primer argumento: el libro EPI-15 del mes (con una hoja por CPT).
  - Siguientes: archivos EPI-12 o una carpeta que los contenga (todas las semanas del mes).
  - Nunca modifica los archivos originales: escribe uno nuevo (<epi15>_AUTO.xlsx).

Qué hace con cada CPT que tenga la hoja del EPI-15 VACÍA (sin ningún número en la columna P):
  1. Suma las semanas del EPI-12 de ese CPT.
  2. Llena las filas del EPI-15 que tienen equivalente en el EPI-12 (tabla EQUIVALENCIAS abajo).
  3. Calcula OTRAS CAUSAS DE CONSULTA (fila 269) para que TOTAL CAUSAS = total de pacientes atendidos.
  4. Escribe el mes y una nota "llenado automáticamente" en la hoja.
Las hojas que ya tienen datos NO se tocan.
"""
import argparse
import re
import sys
import unicodedata
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

# --------------------------------------------------------------------------------------
# Grupos de edad del EPI-12 (cada uno tiene columnas H y M). El índice 12 = "edad ignorada".
# --------------------------------------------------------------------------------------
TODOS = tuple(range(13))                 # todas las edades (incluye ignorada)
MENOR1, DE1A4 = (0,), (1,)
MENOR5, MAS5 = (0, 1), tuple(range(2, 12))
MENOR10, MAS10 = (0, 1, 2, 3), tuple(range(4, 12))
MENOR15, MAS15 = tuple(range(0, 6)), tuple(range(6, 12))
DE15A44, MAS45 = (6, 7, 8), (9, 10, 11)

TOTAL_ATENDIDOS = 96                     # fila "Total pacientes atendidos" del EPI-12
FILA_OTRAS = 269                         # n° de "OTRAS CAUSAS DE CONSULTA" en el EPI-15

# --------------------------------------------------------------------------------------
# EQUIVALENCIAS:  n° fila EPI-15  ->  (lista de n° de evento EPI-12 que se suman, edades)
# Si un diagnóstico del EPI-15 no aparece aquí, queda en blanco (se va a OTRAS CAUSAS).
# Se puede editar esta tabla sin tocar el resto del script.
# --------------------------------------------------------------------------------------
EQUIVALENCIAS = {
    1: ([1], TODOS),      2: ([3], TODOS),
    3: ([2], MENOR1),     4: ([2], DE1A4),      5: ([2], MAS5),          # diarreas por edad
    11: ([4], TODOS),     12: ([5], TODOS),     13: ([6], TODOS),        14: ([7], TODOS),
    15: ([8], TODOS),     19: ([9], TODOS),
    20: ([88], TODOS),    21: ([87], TODOS),    22: ([10], TODOS),
    23: ([11], TODOS),    24: ([12], TODOS),
    25: ([90], TODOS),    26: ([91], TODOS),    27: ([92], TODOS),       33: ([89], TODOS),
    35: ([13], TODOS),    36: ([14], TODOS),    37: ([15], TODOS),       38: ([16], TODOS),
    39: ([17], TODOS),    40: ([18], TODOS),    41: ([19], TODOS),       42: ([20], TODOS),
    43: ([21], TODOS),    44: ([22], TODOS),    45: ([23], TODOS),       46: ([24], TODOS),
    47: ([25], TODOS),    48: ([26], TODOS),    49: ([27], TODOS),
    50: ([71], TODOS),    51: ([70], TODOS),    52: ([72], TODOS),       53: ([73], TODOS),   # paludismo
    54: ([28], TODOS),    55: ([29], TODOS),    56: ([30], TODOS),       57: ([31], TODOS),   # leishmaniasis
    58: ([32], TODOS),    59: ([33], TODOS),    61: ([74], TODOS),       62: ([34], TODOS),
    63: ([35], TODOS),    64: ([36], TODOS),    67: ([64], TODOS),       68: ([67], TODOS),
    69: ([37], TODOS),    70: ([38], TODOS),    71: ([39], TODOS),       72: ([40], TODOS),
    73: ([41], TODOS),    74: ([42], TODOS),    75: ([43], TODOS),       76: ([44], TODOS),
    77: ([45], TODOS),    78: ([94], TODOS),                                # síndrome viral
    83: ([75], TODOS),    84: ([68], TODOS),    86: ([66], TODOS),       87: ([83], TODOS),
    98: ([84], TODOS),
    110: ([46], MENOR15), 111: ([46], MAS15),   112: ([76], TODOS),
    131: ([77], MENOR15), 132: ([77], DE15A44), 133: ([77], MAS45),      # hipertensión
    137: ([79], TODOS),   141: ([78], TODOS),
    142: ([53], MENOR5),  143: ([53], MAS5),    144: ([55], MENOR5),     145: ([55], MAS5),
    146: ([56], TODOS),
    147: ([47], MENOR1),  148: ([47], DE1A4),   149: ([47], MAS5),       # neumonía
    151: ([61], TODOS),   152: ([60], TODOS),
    154: ([81], MENOR10), 155: ([81], MAS10),                             # asma
    157: ([80], TODOS),   158: ([54], TODOS),   159: ([57], TODOS),      160: ([58], TODOS),
    161: ([59], TODOS),   162: ([62], TODOS),   163: ([63], TODOS),      164: ([65], TODOS),
    187: ([93], TODOS),   193: ([85], TODOS),   219: ([69], TODOS),
    221: ([50], TODOS),                                                    # fiebre
    231: ([48], TODOS),   235: ([86], TODOS),   238: ([49], TODOS),
    242: ([51], TODOS),   243: ([52], TODOS),
}

NOTA_AUTO = ("NOTA: Hoja llenada automáticamente a partir de los EPI-12 de las semanas {semanas}. "
             "El CPT no entregó el EPI-15 de este mes.")


# ------------------------------------------------------------------ utilidades de nombres
def clave(nombre):
    """Normaliza el nombre de un CPT para poder cruzar EPI-12 con EPI-15."""
    s = unicodedata.normalize("NFKD", str(nombre)).encode("ascii", "ignore").decode().upper().strip()
    s = re.sub(r"\(.*?\)", " ", s)
    s = re.sub(r"^\d+(?=[A-Z])", "", s)          # '1CPTII ...' -> 'CPTII ...'
    s = re.sub(r"(?<=[A-Z])\d+$", "", s)         # '...APONTE8' -> '...APONTE'
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    ignorar = {"CPT", "CPTI", "CPTII", "CPTIII", "BASE", "MISION", "DE", "DEL", "EL", "LA"}
    return " ".join(w for w in s.split() if w not in ignorar)


def mismo_cpt(a, b):
    return a == b or a.startswith(b + " ") or b.startswith(a + " ")


def a_numero(v):
    if isinstance(v, bool):
        return 0
    if isinstance(v, (int, float)):
        return int(round(v))
    if isinstance(v, str) and v.strip().replace(".", "", 1).isdigit():
        return int(round(float(v)))
    return 0


# ------------------------------------------------------------------ lectura del EPI-12
def leer_epi12(ruta):
    """Devuelve (semana, {clave: {'nombre', 'eventos': {n: [13 sumas por edad]}}})."""
    wb = load_workbook(ruta, read_only=True, data_only=True)
    ws = wb["GUANARITO"] if "GUANARITO" in wb.sheetnames else wb.worksheets[0]
    filas = list(ws.iter_rows(min_row=1, max_col=31, values_only=True))
    wb.close()

    inicios = [i for i, f in enumerate(filas)
               if f[20] and str(f[20]).strip().upper().startswith("ESTABLECIMIENTOS")]
    if not inicios:
        raise ValueError(f"{ruta.name}: no encontré bloques de establecimientos (¿es un EPI-12?)")

    semana = None
    for i in inicios:
        s = a_numero(filas[i + 2][30]) if i + 2 < len(filas) else 0
        if s:
            semana = s
            break
    if semana is None:
        m = re.search(r"(\d+)", ruta.stem[::-1])
        semana = int(m.group(1)[::-1]) if m else 0

    bloques = {}
    for n, i0 in enumerate(inicios):
        i1 = inicios[n + 1] if n + 1 < len(inicios) else len(filas)
        nombre = filas[i0][23]
        eventos = {}
        for j in range(i0, i1):
            f = filas[j]
            try:
                orden = int(str(f[0]).strip())
            except (TypeError, ValueError):
                continue
            if not (1 <= orden <= 97) or orden in eventos or f[1] is None:
                continue
            celdas = [a_numero(x) for x in f[2:28]]               # 13 grupos x (H, M)
            eventos[orden] = [celdas[2 * k] + celdas[2 * k + 1] for k in range(13)]
        bloques[n] = (nombre, eventos)
    return semana, bloques


# ------------------------------------------------------------------ utilidades del EPI-15
def filas_por_numero(ws):
    return {ws.cell(r, 1).value: r for r in range(21, 317) if isinstance(ws.cell(r, 1).value, int)}


def filas_que_suman_en_total(ws):
    """Lee la fórmula de TOTAL CAUSAS (F317) para saber qué filas cuentan en el total."""
    formula = str(ws["F317"].value or "")
    filas = set()
    for a, b in re.findall(r"F(\d+)(?::F(\d+))?", formula):
        filas.update(range(int(a), int(b or a) + 1))
    return filas


def hoja_vacia(ws, filas):
    """Una hoja está vacía solo si NINGUNA fila de diagnóstico tiene algo en la columna P.
    Ojo: cuenta también las fórmulas escritas a mano (ej. =1+1+3), no solo números."""
    return not any(ws.cell(r, 6).value not in (None, "") for r in filas.values())


# ------------------------------------------------------------------ programa principal
def main():
    ap = argparse.ArgumentParser(description="Llena el EPI-15 desde las semanas del EPI-12.")
    ap.add_argument("epi15", help="libro EPI-15 del mes")
    ap.add_argument("epi12", nargs="+", help="archivos EPI-12 (o carpetas) de las semanas del mes")
    ap.add_argument("--mes", help="nombre del mes para escribir en la hoja (ej. OCTUBRE)")
    ap.add_argument("--salida", help="archivo de salida (por defecto <epi15>_AUTO.xlsx)")
    ap.add_argument("--solo", help="lista de CPT (separados por coma) a llenar; por defecto, todos los vacíos")
    ap.add_argument("--otras-total", action="store_true",
                    help="poner en OTRAS CAUSAS el total atendido completo, sin restar los diagnósticos llenados")
    args = ap.parse_args()

    if not Path(args.epi15).is_file():
        parecidos = sorted(p.name for p in (Path(args.epi15).parent).glob("*.xlsx"))
        sys.exit(f"No encuentro el archivo EPI-15: '{args.epi15}'\n"
                 "Ojo: en Linux las mayúsculas y los espacios cuentan. Archivos .xlsx en esa carpeta:\n  "
                 + "\n  ".join(parecidos or ["(ninguno)"])
                 + "\nTip: escribe las primeras letras del nombre y pulsa Tab para autocompletar.")

    # ---- reunir archivos EPI-12
    archivos = []
    for p in map(Path, args.epi12):
        archivos += sorted(x for x in p.glob("*.xlsx") if not x.name.startswith("~$")) if p.is_dir() else [p]
    if not archivos:
        sys.exit("No encontré archivos EPI-12.")

    # ---- leer y acumular semanas
    acum = {}            # clave -> {'nombre', 'eventos': {n: [13]}, 'semanas': [..]}
    semanas = []
    for ruta in archivos:
        semana, bloques = leer_epi12(ruta)
        if semana in semanas:
            print(f"AVISO: la semana {semana} aparece repetida ({ruta.name}). Revisa que no sumes una semana dos veces.")
        semanas.append(semana)
        for nombre, eventos in bloques.values():
            if not nombre or not str(nombre).strip():
                if sum(eventos.get(TOTAL_ATENDIDOS, [0] * 13)):
                    print(f"AVISO: un bloque SIN NOMBRE en {ruta.name} tiene pacientes atendidos; no se pudo asignar.")
                continue
            k = clave(nombre)
            reg = acum.setdefault(k, {"nombre": str(nombre).strip(), "eventos": {}, "semanas": []})
            for n, v in eventos.items():
                dest = reg["eventos"].setdefault(n, [0] * 13)
                for g in range(13):
                    dest[g] += v[g]
            if sum(eventos.get(TOTAL_ATENDIDOS, [0] * 13)) > 0:
                reg["semanas"].append(semana)
    semanas = sorted(set(semanas))
    texto_semanas = ", ".join(map(str, semanas))
    print(f"EPI-12 leídos: {len(archivos)} archivo(s), semanas {texto_semanas}")
    if len(semanas) not in (4, 5):
        print(f"AVISO: normalmente un mes tiene 4 o 5 semanas; recibí {len(semanas)}.")

    # ---- abrir EPI-15
    wb = load_workbook(args.epi15)
    solo = {clave(x) for x in args.solo.split(",")} if args.solo else None
    reporte, usados = [], set()

    for ws in wb.worksheets:
        if ws.title.strip().upper() == "HOJA EN BLANCO":
            continue
        k15 = clave(ws.title)
        filas = filas_por_numero(ws)
        if not hoja_vacia(ws, filas):
            reporte.append((ws.title, "ya tiene datos (no se toca)"))
            continue
        if solo is not None and not any(mismo_cpt(k15, s) for s in solo):
            continue
        k12 = next((k for k in acum if mismo_cpt(k15, k)), None)
        if k12 is None:
            reporte.append((ws.title, "SIN BLOQUE EN EPI-12 (revisar nombre)"))
            continue
        usados.add(k12)
        reg = acum[k12]
        total = sum(reg["eventos"].get(TOTAL_ATENDIDOS, [0] * 13))
        if total == 0:
            reporte.append((ws.title, "no tiene pacientes en ninguna semana del EPI-12: se deja VACÍA"))
            continue

        # ---- diagnósticos con equivalencia
        cuentan = filas_que_suman_en_total(ws)
        llenados, suma_en_total, perdidos_edad_ign = 0, 0, 0
        for n15, (ev12, edades) in EQUIVALENCIAS.items():
            if n15 not in filas:
                continue
            valor = 0
            for e in ev12:
                v = reg["eventos"].get(e, [0] * 13)
                valor += sum(v[g] for g in edades)
                if 12 not in edades:
                    perdidos_edad_ign += v[12]
            if valor > 0:
                ws.cell(filas[n15], 6).value = valor
                llenados += 1
                if filas[n15] in cuentan:
                    suma_en_total += valor

        # ---- otras causas de consulta
        if args.otras_total:
            otras = total
        else:
            otras = total - suma_en_total
        avisos = []
        if otras < 0:
            avisos.append(f"los diagnósticos sumaron {suma_en_total} > atendidos {total}; OTRAS quedó en 0")
            otras = 0
        ws.cell(filas[FILA_OTRAS], 6).value = otras
        if perdidos_edad_ign:
            avisos.append(f"{perdidos_edad_ign} caso(s) con 'edad ignorada' no entraron en filas separadas por edad")

        # ---- mes y nota
        if args.mes:
            ws["I4"].value = args.mes.upper()
        ws["A3"].value = NOTA_AUTO.format(semanas=texto_semanas)
        ws["A3"].font = Font(name="Arial", size=9, italic=True, bold=True, color="C00000")
        ws["A3"].alignment = Alignment(horizontal="left", vertical="center")

        faltan = [s for s in semanas if s not in reg["semanas"]]
        if faltan:
            avisos.append(f"sin datos en EPI-12 de la(s) semana(s) {', '.join(map(str, faltan))}")
        detalle = f"llenado: {llenados} diagnósticos + OTRAS={otras} (atendidos {total})"
        if avisos:
            detalle += "  |  AVISO: " + "; ".join(avisos)
        reporte.append((ws.title, detalle))

    # ---- guardar
    wb.calculation.fullCalcOnLoad = True
    salida = Path(args.salida) if args.salida else Path(args.epi15).with_name(Path(args.epi15).stem + "_AUTO.xlsx")
    wb.save(salida)

    sin_hoja = [acum[k]["nombre"] for k in acum if k not in usados
                and not any(mismo_cpt(clave(w.title), k) for w in wb.worksheets)]
    lineas = [f"Archivo generado: {salida}", f"Semanas usadas: {texto_semanas}", ""]
    lineas += [f"- {t}: {d}" for t, d in reporte]
    if sin_hoja:
        lineas += ["", "Bloques del EPI-12 sin hoja en el EPI-15: " + ", ".join(sin_hoja)]
    texto = "\n".join(lineas)
    print("\n" + texto)
    salida.with_name(salida.stem + "_reporte.txt").write_text(texto, encoding="utf-8")


if __name__ == "__main__":
    main()
