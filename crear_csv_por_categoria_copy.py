import csv
import codecs
import os
import re
import tempfile
import time
import traceback
import unicodedata
from pathlib import Path

import pandas as pd

BASES_DIR = "/workspaces/Cinvestav/01 Informacion de las bases/Bases de datos/Egresos secretaria"
LISTA_SOBREDOSIS = "/workspaces/Cinvestav/02 Tareas/12 Causas de egresos/Diccionario/Sobredosis CIE10  por sustancias.xlsx"
TABLA_DERECHOHABIENCIA = "/workspaces/Cinvestav/02 Tareas/12 Causas de egresos/Diccionario/Derechohabiencia.xlsx"
SALIDA = "/workspaces/Cinvestav/02 Tareas/12 Causas de egresos/Resumen de CIE10"

ANIOS = list(range(2010, 2025))   
DECIMALES_EDAD = 2                
CHUNK = 500_000                   

CATEGORIAS_EXTRA = {
    "Drogas inyectables": ["A41", "B171", "B182", "B24", "B376", "I330", "I339", "I38", "I39", "L022", "L024", "L03x"],
    "Prueba en sangre": ["R781", "R782", "R783", "R784", "R785"],
}

COLS_NECESARIAS = ["ID", "ENTIDAD", "MUNIC", "SEXO", "EDAD", "CVEEDAD", "DERHAB", "AFECPRIN", "CAUSAEXT"]
COLS_SALIDA = ["ENTIDAD", "MUNIC", "ANIO", "SEXO", "EDAD_ANIOS", "DERHAB", "DERECHOHABIENCIA2",
               "CIE10_BASE", "CIE10_CATEGORIA"]

UNIDADES_ANTES_2020 = {"0": "h", "1": "d", "2": "m", "3": "a"}
UNIDADES_DESDE_2020 = {"2": "h", "3": "d", "4": "m", "5": "a"}
FACTOR_A_ANIOS = {"h": 1 / (24 * 365), "d": 1 / 365, "m": 1 / 12, "a": 1.0}


def ruta(p):
    return Path(str(p).strip().strip('"'))


def sin_acentos(t):
    return "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn")


def limpia(s):
    """Texto sin espacios; '', 'NULL' y 'nan' pasan a vacío (NaN)."""
    s = s.astype(object).str.strip()
    return s.where(~s.isin(["", "NULL", "null", "NaN", "nan"]))


def norm_num(s):
    """Códigos numéricos sin '.0' ni ceros a la izquierda ('01' -> '1'); letras en mayúscula."""
    s = limpia(s).astype(object).str.upper().str.replace(r"\.0$", "", regex=True)
    es_num = s.str.fullmatch(r"\d+").fillna(False).astype(bool)
    s = s.copy()
    s[es_num] = s[es_num].str.lstrip("0").replace("", "0")
    return s


def norm_geo(s, ancho):
    """Clave geográfica como texto con ceros a la izquierda (ENTIDAD 2 dígitos, MUNIC 3)."""
    s = limpia(s).astype(object).str.replace(r"\.0$", "", regex=True)
    es_num = s.str.fullmatch(r"\d+").fillna(False).astype(bool)
    s = s.copy()
    s[es_num] = s[es_num].str.zfill(ancho)
    return s


def norm_codigo(x):
    x = re.sub(r"\.0$", "", str(x).strip().upper())
    return str(int(x)) if x.isdigit() else x


# ----------------------------------- catálogos -----------------------------------
def cargar_codigos():
    lista = pd.read_excel(ruta(LISTA_SOBREDOSIS), dtype=str).iloc[:, :3]
    lista.columns = ["codigo", "descripcion", "categoria"]
    lista = lista.dropna(subset=["codigo", "categoria"])
    pares = [(c.strip().upper(), cat.strip()) for c, cat in zip(lista.codigo, lista.categoria)]
    for cat, codigos in CATEGORIAS_EXTRA.items():
        pares += [(c.strip().upper(), cat) for c in codigos]

    exactos, prefijos, categorias = {}, {}, []
    for cod, cat in pares:
        if cat not in categorias:
            categorias.append(cat)
        if cod.endswith("X") and len(cod) >= 3:        # L03x -> prefijo L03
            destino, clave = prefijos, cod[:-1]
        elif len(cod) == 4:
            destino, clave = exactos, cod
        elif len(cod) == 3:
            destino, clave = prefijos, cod
        else:
            raise ValueError(f"Código con formato no previsto: {cod}")
        if clave in destino and destino[clave] != cat:
            raise ValueError(f"El código {cod} está en dos categorías: {destino[clave]} y {cat}")
        destino[clave] = cat
    for e in exactos:
        if e[:3] in prefijos:
            raise ValueError(f"El código {e} se traslapa con el prefijo {e[:3]}")
    return exactos, prefijos, categorias


def cargar_tabla_derechohabiencia():
    t = pd.read_excel(ruta(TABLA_DERECHOHABIENCIA), header=None, dtype=str)
    patron = re.compile(r"^\d{4}(-\d{4})?$")
    fila_anios = next(i for i in range(len(t))
                      if sum(bool(patron.match(str(x).strip())) for x in t.iloc[i]) >= 5)
    tablas = {}
    for c in range(t.shape[1] - 1):
        h = str(t.iloc[fila_anios, c]).strip()
        if not patron.match(h):
            continue
        d = {}
        for r in range(fila_anios + 2, len(t)):            # +2: salta la fila "ID | Descripción"
            if t.iloc[r].isna().all():
                break
            k, v = t.iloc[r, c], t.iloc[r, c + 1]
            if pd.notna(k) and pd.notna(v):
                d[norm_codigo(k)] = str(v).strip()
        a, b = h.split("-") if "-" in h else (h, h)
        for y in range(int(a), int(b) + 1):
            tablas[y] = d
    return tablas


# ----------------------------------- lectura de archivos -----------------------------------
def encontrar_archivo(anio):
    cand = sorted(p for p in ruta(BASES_DIR).iterdir()
                  if p.suffix.lower() == ".csv" and str(anio) in p.stem)
    if not cand:
        return None
    if len(cand) > 1:
        print(f"   Aviso: varios archivos para {anio}, se usa {cand[0].name}")
    return cand[0]


def preparar_archivo(path):
    """Detecta codificación, separador y filas entre comillas. Devuelve (ruta_a_leer, sep, enc, es_temporal)."""
    with open(path, "rb") as f:
        muestra = f.read(2_000_000)
    enc, texto = "latin-1", None
    try:
        texto = codecs.getincrementaldecoder("utf-8-sig")().decode(muestra, final=False)
        enc = "utf-8-sig"
    except UnicodeDecodeError:
        texto = muestra.decode("latin-1")
    primera = texto.splitlines()[0]
    entrecomillada = primera.startswith('"') and primera.endswith('"') and primera.count('"') == 2
    linea = primera.strip('"')
    sep = "|" if linea.count("|") > linea.count(",") else ","
    if not entrecomillada:
        return path, sep, enc, False
    # Cada fila viene completa entre comillas (pasa en 2019): se quitan para poder separar columnas
    tmp = tempfile.NamedTemporaryFile("w", delete=False, suffix=".csv", encoding="utf-8", newline="")
    with open(path, encoding=enc, errors="replace", newline="") as fin, tmp:
        for ln in fin:
            ln = ln.rstrip("\r\n")
            if ln.startswith('"') and ln.endswith('"'):
                ln = ln[1:-1].replace('""', '"')
            tmp.write(ln + "\n")
    return Path(tmp.name), sep, "utf-8", True


# ----------------------------------- transformaciones -----------------------------------
def clasifica(valores, exactos, prefijos):
    """Categoría de cada código (NaN si no está en la lista)."""
    v = valores.fillna("")
    cat = v.str[:4].map(exactos)
    return cat.fillna(v.str[:3].map(prefijos))


def edad_en_anios(anio, edad, cve):
    mapa = UNIDADES_DESDE_2020 if anio >= 2020 else UNIDADES_ANTES_2020
    e = pd.to_numeric(limpia(edad), errors="coerce")
    e = e.mask(e == 999)                                   # 999 = se ignora
    factor = norm_num(cve).map(mapa).map(FACTOR_A_ANIOS)   # unidad 9 o desconocida -> vacío
    return (e * factor).round(DECIMALES_EDAD)


def finalizar(anio, categoria, piezas, tabla_anio, contadores):
    if piezas:
        d = pd.concat(piezas, ignore_index=True)
    else:
        d = pd.DataFrame(columns=["ENTIDAD", "MUNIC", "SEXO", "EDAD", "CVEEDAD", "DERHAB", "CIE10_BASE"])
    out = pd.DataFrame(index=d.index)
    out["ENTIDAD"] = norm_geo(d["ENTIDAD"], 2)
    out["MUNIC"] = norm_geo(d["MUNIC"], 3)
    out["ANIO"] = anio
    out["SEXO"] = norm_num(d["SEXO"])
    out["EDAD_ANIOS"] = edad_en_anios(anio, d["EDAD"], d["CVEEDAD"])
    out["DERHAB"] = norm_num(d["DERHAB"])
    nombres = out["DERHAB"].map(tabla_anio)
    out["DERECHOHABIENCIA2"] = nombres.where(
        nombres.notna(), out["DERHAB"].isna().map({True: "SIN DATO", False: "SIN EQUIVALENCIA"}))
    out["CIE10_BASE"] = d["CIE10_BASE"]
    out["CIE10_CATEGORIA"] = categoria
    out = out[COLS_SALIDA]
    # para el resumen
    contadores["edad_sin_valor"] += int(out["EDAD_ANIOS"].isna().sum())
    contadores["derhab_sin_dato"] += int((out["DERECHOHABIENCIA2"] == "SIN DATO").sum())
    sin_eq = out.loc[out["DERECHOHABIENCIA2"] == "SIN EQUIVALENCIA", "DERHAB"]
    for cod, n in sin_eq.value_counts().items():
        contadores["sin_equivalencia"][cod] = contadores["sin_equivalencia"].get(cod, 0) + int(n)
    return out


# ----------------------------------- proceso por año -----------------------------------
def procesar_anio(anio, exactos, prefijos, categorias, tablas_derhab):
    archivo = encontrar_archivo(anio)
    if archivo is None:
        print(f"   No se encontró archivo para {anio}")
        return None, []
    if anio not in tablas_derhab:
        raise ValueError(f"Derechohabiencia.xlsx no tiene columna para el año {anio}")

    ruta_lectura, sep, enc, es_tmp = preparar_archivo(archivo)
    contadores = {"filas_leidas": 0, "cabeceras_descartadas": 0, "egresos_con_coincidencia": 0,
                  "egresos_en_mas_de_una_categoria": 0, "edad_sin_valor": 0, "derhab_sin_dato": 0,
                  "sin_equivalencia": {}}
    piezas = {c: [] for c in categorias}
    try:
        lector = pd.read_csv(ruta_lectura, sep=sep, dtype=str, keep_default_na=False, chunksize=CHUNK,
                             usecols=lambda c: c.strip().upper() in COLS_NECESARIAS,
                             encoding=enc, encoding_errors="replace")
        for df in lector:
            df.columns = [c.strip().upper() for c in df.columns]
            falta = [c for c in COLS_NECESARIAS if c not in df.columns]
            if falta:
                raise ValueError(f"Faltan columnas en {archivo.name}: {falta}")
            contadores["filas_leidas"] += len(df)

            cab = (df["ID"].str.strip().str.upper() == "ID") | (df["AFECPRIN"].str.strip().str.upper() == "AFECPRIN")
            contadores["cabeceras_descartadas"] += int(cab.sum())
            df = df[~cab]

            af = limpia(df["AFECPRIN"]).str.upper()
            ce = limpia(df["CAUSAEXT"]).str.upper()
            cat_af, cat_ce = clasifica(af, exactos, prefijos), clasifica(ce, exactos, prefijos)
            hay = cat_af.notna() | cat_ce.notna()
            if not hay.any():
                continue
            d, af, ce, cat_af, cat_ce = df[hay], af[hay], ce[hay], cat_af[hay], cat_ce[hay]
            contadores["egresos_con_coincidencia"] += len(d)
            contadores["egresos_en_mas_de_una_categoria"] += int(
                (cat_af.notna() & cat_ce.notna() & (cat_af != cat_ce)).sum())

            base = d[["ENTIDAD", "MUNIC", "SEXO", "EDAD", "CVEEDAD", "DERHAB"]]
            for cat in categorias:
                m = (cat_af == cat) | (cat_ce == cat)
                if m.any():
                    sub = base[m].copy()
                    sub["CIE10_BASE"] = af[m].where(cat_af[m] == cat, ce[m])
                    piezas[cat].append(sub)
    finally:
        if es_tmp:
            os.remove(ruta_lectura)

    carpeta = ruta(SALIDA) / f"{anio} Causas de egresos"
    carpeta.mkdir(parents=True, exist_ok=True)
    fila = {"anio": anio, "archivo": archivo.name}
    fila.update({k: v for k, v in contadores.items() if k != "sin_equivalencia"})
    for cat in categorias:
        out = finalizar(anio, cat, piezas[cat], tablas_derhab[anio], contadores)
        out.to_csv(carpeta / f"{anio} {sin_acentos(cat)}.csv", index=False, encoding="utf-8-sig")
        fila[f"filas_{sin_acentos(cat)}"] = len(out)
    fila["edad_sin_valor"] = contadores["edad_sin_valor"]
    fila["derhab_sin_dato"] = contadores["derhab_sin_dato"]
    sin_eq = [{"anio": anio, "codigo_derhab": k, "filas": v} for k, v in sorted(contadores["sin_equivalencia"].items())]
    fila["derhab_sin_equivalencia"] = sum(contadores["sin_equivalencia"].values())
    return fila, sin_eq


def main():
    exactos, prefijos, categorias = cargar_codigos()
    tablas = cargar_tabla_derechohabiencia()
    print(f"Categorías ({len(categorias)}): {categorias}")
    print(f"Códigos exactos: {len(exactos)} | prefijos: {len(prefijos)}\n")

    salida = ruta(SALIDA)
    salida.mkdir(parents=True, exist_ok=True)
    resumen, sin_equivalencia, fallidos = [], [], []
    for anio in ANIOS:
        t0 = time.time()
        print(f"== {anio}")
        try:
            fila, sin_eq = procesar_anio(anio, exactos, prefijos, categorias, tablas)
        except Exception:
            traceback.print_exc()
            fallidos.append(anio)
            continue
        if fila is None:
            fallidos.append(anio)
            continue
        resumen.append(fila)
        sin_equivalencia += sin_eq
        pd.DataFrame(resumen).to_csv(salida / "Resumen de conteos.csv", index=False, encoding="utf-8-sig")
        pd.DataFrame(sin_equivalencia, columns=["anio", "codigo_derhab", "filas"]).to_csv(
            salida / "Derechohabiencia sin equivalencia.csv", index=False, encoding="utf-8-sig")
        print(f"   filas leídas: {fila['filas_leidas']:,} | egresos con coincidencia: {fila['egresos_con_coincidencia']:,} "
              f"| en más de una categoría: {fila['egresos_en_mas_de_una_categoria']:,} | {time.time() - t0:.0f} s")

    print("\nListo. Resultados en:", salida)
    if fallidos:
        print("Años que NO se pudieron procesar:", fallidos)


if __name__ == "__main__":
    main()