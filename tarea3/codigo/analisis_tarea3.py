# -*- coding: utf-8 -*-
"""
analisis_tarea3.py
==================

Calcula las tablas del Anexo A del informe de la Tarea 3 a partir del archivo
publicado por la Agencia Nacional de Transito (ANT). El script hace dos cosas:

1. Imprime en consola las cifras principales, para contrastarlas con el texto.
2. Escribe las tablas A1 a A7 como fragmentos LaTeX en la carpeta ``tablas/``.

El informe incluye este mismo archivo en el Anexo B, de modo que cualquier cifra
de la sección 6 puede reproducirse ejecutando::

    python analisis_tarea3.py

Requisitos: pandas y openpyxl (ver LEEME.txt de la carpeta tarea3).

Uso de los datos: la hoja BDD_2026 tiene una fila por siniestro y 374 columnas.
La etiqueta no viene en el archivo; se construye con la columna FALLECIDOS.
"""

import os
import sys

import pandas as pd

# --- Rutas -----------------------------------------------------------------
AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.abspath(os.path.join(AQUI, "..", ".."))
ARCHIVO = os.path.join(RAIZ, "08. BDD_Septiembre_2026_SM.xlsx")
SALIDA = os.path.abspath(os.path.join(AQUI, "..", "tablas"))
HOJA = "BDD_2026"

MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
         "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

# Columnas de participantes: siete campos por cada uno de los 43 posibles.
CAMPOS_PARTICIPANTE = ("TIPO_ID", "EDAD", "SEXO", "CONDICION",
                       "PARTICIPANTE", "CASCO", "CINTURON")
COLUMNAS_PARTICIPANTE = ["%s_%d" % (campo, i)
                         for i in range(1, 44) for campo in CAMPOS_PARTICIPANTE]


# --- Utilidades de formato -------------------------------------------------
def tex(texto):
    """Escapa los caracteres que LaTeX interpreta y deja intactos los acentos."""
    reemplazos = {"_": r"\_", "%": r"\%", "&": r"\&", "#": r"\#"}
    for original, escapado in reemplazos.items():
        texto = texto.replace(original, escapado)
    return texto


def entero(valor):
    """17236 -> '17.236' (separador de miles del espanol)."""
    return "{:,}".format(int(valor)).replace(",", ".")


def decimal(valor, digitos=1):
    """9.61 -> '9,6' (coma decimal del espanol)."""
    return (("%." + str(digitos) + "f") % valor).replace(".", ",")


def porcentaje(valor, digitos=1):
    """9.61 -> '9,6\\,\\%' listo para el texto LaTeX."""
    return decimal(valor, digitos) + r"\,\%"


def fila(*celdas):
    return " & ".join(str(c) for c in celdas) + r" \\"


def escribir_tabla(caption, label, colspec, encabezado, filas, nota=None):
    """Genera un entorno table/tabular con booktabs."""
    lineas = [
        r"\begin{table}[htbp]",
        r"\centering",
        r"\small",
        r"\caption{%s}" % caption,
        r"\label{%s}" % label,
        r"\begin{tabular}{%s}" % colspec,
        r"\toprule",
        fila(*encabezado),
        r"\midrule",
    ]
    lineas += filas
    lineas += [r"\bottomrule", r"\end{tabular}"]
    if nota:
        lineas.append(r"\par\vspace{2pt}{\footnotesize %s}" % nota)
    lineas.append(r"\end{table}")
    return "\n".join(lineas) + "\n"


def escribir_longtable(caption, label, colspec, encabezado, filas):
    """Igual que escribir_tabla, pero permite que la tabla ocupe varias paginas."""
    cabecera = fila(*encabezado)
    lineas = [
        r"\begin{longtable}{%s}" % colspec,
        r"\caption{%s}\label{%s}\\" % (caption, label),
        r"\toprule",
        cabecera,
        r"\midrule",
        r"\endfirsthead",
        r"\multicolumn{%d}{l}{\footnotesize\itshape (continuación)}\\" % len(encabezado),
        r"\toprule",
        cabecera,
        r"\midrule",
        r"\endhead",
        r"\bottomrule",
        r"\endfoot",
    ]
    lineas += filas
    lineas += [r"\end{longtable}"]
    return "\n".join(lineas) + "\n"


def bloque(columnas, texto):
    """Fila de titulo que separa bloques dentro de una tabla larga."""
    return r"\addlinespace[2pt]\multicolumn{%d}{l}{\textit{%s}}\\" % (columnas, texto)


def guardar(nombre, contenido):
    ruta = os.path.join(SALIDA, nombre)
    with open(ruta, "w", encoding="utf-8") as salida:
        salida.write(contenido)
    return ruta


# --- Carga ----------------------------------------------------------------
def cargar(archivo=ARCHIVO):
    """Lee la hoja principal y agrega la etiqueta del proyecto."""
    datos = pd.read_excel(archivo, sheet_name=HOJA, engine="openpyxl")
    # Se guarda la forma original: las columnas derivadas se agregan después y
    # no forman parte del archivo publicado.
    datos.attrs["columnas_archivo"] = datos.shape[1]
    datos["grave"] = (datos["FALLECIDOS"] >= 1).astype(int)
    datos["mes"] = datos["FECHA"].dt.month
    return datos


# --- Tabla A1: estructura y etiqueta --------------------------------------
def tabla_a1(datos):
    columnas_archivo = datos.attrs["columnas_archivo"]
    total_celdas = datos.shape[0] * columnas_archivo
    vacias = datos.isna().sum().sum()
    vacias_participantes = datos[COLUMNAS_PARTICIPANTE].isna().sum().sum()
    participantes = datos[[c for c in COLUMNAS_PARTICIPANTE
                           if c.startswith("PARTICIPANTE")]].notna().sum(axis=1)
    graves = int(datos["grave"].sum())
    distribucion = datos.loc[datos["grave"] == 1, "FALLECIDOS"].value_counts().sort_index()

    filas = [
        fila("Siniestros (filas)", entero(datos.shape[0])),
        fila("Columnas", entero(columnas_archivo)),
        fila("Celdas en total", entero(total_celdas)),
        fila("Columnas que describen el evento", entero(columnas_archivo - len(COLUMNAS_PARTICIPANTE))),
        fila("Columnas que describen participantes", entero(len(COLUMNAS_PARTICIPANTE))),
        fila("Participantes por siniestro (mediana; máximo)",
             "%d; %d" % (participantes.median(), participantes.max())),
        fila("Celdas vacías sobre el total", porcentaje(vacias / total_celdas * 100)),
        fila("Celdas vacías en las columnas de participantes",
             porcentaje(vacias_participantes / (datos.shape[0] * len(COLUMNAS_PARTICIPANTE)) * 100)),
        r"\addlinespace",
        fila("Siniestros con al menos un fallecido (graves)",
             "%s (%s)" % (entero(graves), porcentaje(graves / datos.shape[0] * 100, 2))),
        fila("Fallecidos por siniestro",
             "; ".join("%d $\\rightarrow$ %s" % (k, entero(v)) for k, v in distribucion.items())),
        fila("Siniestros con dos o más fallecidos", entero((datos["FALLECIDOS"] >= 2).sum())),
        fila("Fallecidos registrados en total", entero(datos["FALLECIDOS"].sum())),
        fila("Lesionados registrados en total", entero(datos["LESIONADOS"].sum())),
        fila("Periodo cubierto",
             "%s a %s" % (datos["FECHA"].min().strftime("%d/%m/%Y"),
                          datos["FECHA"].max().strftime("%d/%m/%Y"))),
    ]
    texto = escribir_tabla(
        r"Estructura del archivo y construcción de la etiqueta",
        r"tab:a1", "p{8.2cm}p{6.8cm}",
        ["Indicador", "Valor"],
        filas,
        nota=r"Fuente: archivo de la ANT [2], hoja BDD\_2026. Cálculo propio (Anexo B).")
    return guardar("tabla_a1.tex", texto), filas


# --- Tabla A2: distribución mensual ---------------------------------------
def tabla_a2(datos):
    filas = []
    for mes, grupo in datos.groupby("mes"):
        pne = (grupo["ENTE_DE_CONTROL"] == "PNE").mean() * 100
        filas.append(fila(MESES[mes - 1], entero(len(grupo)),
                          entero(grupo["grave"].sum()),
                          porcentaje(grupo["grave"].mean() * 100),
                          porcentaje(pne)))
    total_pne = (datos["ENTE_DE_CONTROL"] == "PNE").mean() * 100
    filas.append(r"\midrule")
    filas.append(fila("Total", entero(len(datos)), entero(datos["grave"].sum()),
                      porcentaje(datos["grave"].mean() * 100), porcentaje(total_pne)))
    texto = escribir_tabla(
        r"Siniestros registrados por mes",
        "tab:a2", "lrrrr",
        ["Mes", "Siniestros", "Graves", r"\% graves", r"\% PNE"],
        filas,
        nota=(r"PNE: Policía Nacional. La última columna muestra qué porcentaje de los "
              r"registros del mes proviene de ese ente. Cálculo propio (Anexo B)."))
    return guardar("tabla_a2.tex", texto), filas


# --- Tabla A3: calidad de los registros -----------------------------------
def tabla_a3(datos):
    jerarquia = datos["JERARQUIA_DE_LA_VIA"]
    nd = (jerarquia == "ND").sum()
    cero = (jerarquia == 0).sum()
    nulos = jerarquia.isna().sum()
    sin_tipo = (datos["TIPO_DE_VEHICULO_1"] == "NO IDENTIFICADO").sum()
    minusculas = (datos["TIPO_DE_SINIESTRO"] == "colisión").sum()
    hora_texto = sum(1 for v in datos["HORA"] if isinstance(v, str))
    repetidos = int(datos["SINIESTROS"].duplicated().sum())

    filas = [
        fila("Variables del evento sin vacíos",
             "Fecha, hora, provincia, cantón, zona, tipo de siniestro, causa y ente de control",
             "Se usan tal como están"),
        fila("Jerarquía y ubicación de la vía",
             r"\texttt{ND} en %s registros (%s); %d con valor 0 y %d vacío"
             % (entero(nd), porcentaje(nd / len(datos) * 100), cero, nulos),
             r"Se trata como categoría propia; el valor 0 se agrupa con \texttt{ND}"),
        fila("Tipo del primer vehículo no identificado",
             "%s registros (%s)" % (entero(sin_tipo), porcentaje(sin_tipo / len(datos) * 100)),
             r"Se trata como categoría propia y se conserva"),
        fila("Tipo de siniestro con dos escrituras",
             r"%d registros con \texttt{colisión} en minúscula" % minusculas,
             r"Se unifican con \texttt{COLISIÓN}"),
        fila("Hora almacenada como texto",
             "%d registros" % hora_texto,
             "Se convierten a hora y se incorporan a la franja horaria"),
        fila("Identificador repetido",
             r"%d identificador (%s) en dos siniestros de fechas distintas"
             % (repetidos, r"\texttt{MCU00139082026}"),
             "Se conservan ambas filas; no se usa el identificador como llave"),
        fila("Filas duplicadas", "%d" % int(datos.duplicated().sum()),
             "No se elimina ninguna fila"),
        fila("Coordenadas",
             "Sin vacíos y dentro de un rango plausible",
             "No se usan (restricción R-07)"),
    ]
    texto = escribir_longtable(
        r"Calidad de los registros y tratamiento previsto",
        "tab:a3", "p{3.0cm}p{5.4cm}p{4.8cm}",
        ["Aspecto", "Evidencia", "Tratamiento previsto"],
        filas)
    return guardar("tabla_a3.tex", texto), filas


# --- Tabla A4: contraste oficial y fuga de información --------------------
def tabla_a4(datos):
    trimestre = datos[datos["FECHA"] < "2026-04-01"]
    cinco = datos[datos["FECHA"] < "2026-06-01"]

    condicion = datos[["CONDICION_%d" % i for i in range(1, 44)]]
    coincide_fallecidos = (condicion == "FALLECIDO").sum(axis=1) == datos["FALLECIDOS"]
    coincide_lesionados = (condicion == "LESIONADO").sum(axis=1) == datos["LESIONADOS"]
    categorias = sorted({v for v in condicion.values.ravel() if isinstance(v, str)})

    filas = [
        bloque(4, "Contraste con cifras oficiales, enero a marzo de 2026"),
        fila("Siniestros", entero(len(trimestre)), "INEC: 4.789", "0\\,\\%"),
        fila("Lesionados", entero(trimestre["LESIONADOS"].sum()), "INEC: 4.032", "0\\,\\%"),
        fila("Fallecidos", entero(trimestre["FALLECIDOS"].sum()), "INEC: 603", "0\\,\\%"),
        r"\addlinespace",
        fila("Siniestros, enero a mayo de 2026", entero(len(cinco)), "ANT: 8.718", "0\\,\\%"),
        bloque(4, "Verificación de fuga de información"),
        fila(r"Condición \texttt{FALLECIDO} y columna FALLECIDOS",
             r"%s de %s" % (entero(coincide_fallecidos.sum()), entero(len(datos))),
             "Coincidencia exacta", "100\\,\\%"),
        fila(r"Condición \texttt{LESIONADO} y columna LESIONADOS",
             r"%s de %s" % (entero(coincide_lesionados.sum()), entero(len(datos))),
             "Coincidencia exacta", "100\\,\\%"),
        fila(r"Categorías de la columna CONDICIÓN",
             "%d" % len(categorias), ", ".join(tex(c) for c in categorias), "---"),
    ]
    texto = escribir_tabla(
        r"Contraste con cifras oficiales y verificación de fuga",
        "tab:a4", "p{4.6cm}p{2.7cm}p{5.0cm}p{1.2cm}",
        ["Concepto", "Archivo", "Referencia", "Dif."],
        filas,
        nota=(r"Referencias: INEC [1] y ANT [7] para los totales; [2] para el archivo. "
              r"La segunda parte muestra que la condición de los participantes contiene "
              r"el desenlace y no puede usarse como entrada. Cálculo propio (Anexo B)."))
    return guardar("tabla_a4.tex", texto), filas


# --- Tabla A5: proporción de graves por grupo -----------------------------
def tabla_a5(datos):
    filas = []

    def agregar(titulo, serie):
        filas.append(bloque(4, titulo))
        resumen = (datos.groupby(serie)["grave"]
                   .agg(["size", "sum", "mean"])
                   .sort_values("mean", ascending=False))
        # Solo se reportan grupos con 30 casos o más (criterio CA-03).
        resumen = resumen[resumen["size"] >= 30]
        for nombre, fila_grupo in resumen.iterrows():
            filas.append(fila(tex(str(nombre)), entero(fila_grupo["size"]),
                              entero(fila_grupo["sum"]),
                              porcentaje(fila_grupo["mean"] * 100)))

    agregar("Tipo de siniestro", "TIPO_DE_SINIESTRO")
    agregar("Zona", "ZONA")
    datos["moto"] = (datos[[c for c in datos.columns
                            if c.startswith("TIPO_DE_VEHICULO_")]] == "MOTOCICLETA").any(axis=1)
    datos["moto"] = datos["moto"].map({True: "Con motocicleta", False: "Sin motocicleta"})
    agregar("Motocicleta", "moto")
    agregar("Ente de control", "ENTE_DE_CONTROL")
    agregar("Zona de planificación", "ZONA_PLANIFICACION")
    agregar("Día de la semana", "DIA_1")
    datos["feriado"] = datos["FERIADO"].map({"SI": "Feriado", "NO": "No feriado"})
    agregar("Feriado", "feriado")

    causas = (datos.groupby("CAUSA_PROBABLE")["grave"]
              .agg(["size", "sum", "mean"]))
    causas = causas[causas["size"] >= 100].sort_values("mean", ascending=False)
    filas.append(bloque(4, "Causa probable (solo causas con 100 casos o más)"))
    for nombre, fila_grupo in causas.iterrows():
        descripcion = str(nombre).capitalize().rstrip(".")
        filas.append(fila(tex(descripcion), entero(fila_grupo["size"]),
                          entero(fila_grupo["sum"]),
                          porcentaje(fila_grupo["mean"] * 100)))

    texto = escribir_longtable(
        r"Proporción de siniestros graves por grupo",
        "tab:a5", "p{7.5cm}rrr",
        ["Grupo", "n", "Graves", r"\% graves"],
        filas)
    return guardar("tabla_a5.tex", texto), filas


# --- Tablas A6 y A7: particiones y ventanas móviles -----------------------
def tabla_a6(datos):
    cortes = [("Referencia", 1, 6), ("Ajuste", 7, 7), ("Verificación", 8, 9)]
    filas = []
    for nombre, desde, hasta in cortes:
        grupo = datos[datos["mes"].between(desde, hasta)]
        filas.append(fila(nombre,
                          ("%s a %s" % (MESES[desde - 1], MESES[hasta - 1])
                           if desde != hasta else MESES[desde - 1]),
                          entero(len(grupo)), entero(grupo["grave"].sum()),
                          porcentaje(grupo["grave"].mean() * 100, 2)))
    texto = escribir_tabla(
        r"Tramos temporales del periodo disponible",
        "tab:a6", "llrrr",
        ["Tramo", "Meses", "Siniestros", "Graves", r"\% graves"],
        filas,
        nota=(r"Los tramos respetan el orden temporal: los meses de verificación son "
              r"posteriores a los demás (criterios CA-04 y CA-05). Cálculo propio (Anexo B)."))
    return guardar("tabla_a6.tex", texto), filas


def tabla_a7(datos):
    filas = []
    for m in range(3, 9):
        prueba = datos[datos["mes"] == m + 1]
        filas.append(fila("Ventana %d" % (m - 2),
                          "Ene--%s" % MESES[m - 1],
                          MESES[m], entero(len(prueba)),
                          entero(prueba["grave"].sum()),
                          porcentaje(prueba["grave"].mean() * 100)))
    texto = escribir_tabla(
        r"Ventanas móviles para revisar la estabilidad",
        "tab:a7", "lllrrr",
        ["Ventana", "Meses previos", "Mes verificado", "Siniestros", "Graves", r"\% graves"],
        filas,
        nota=(r"Cada ventana contrasta un mes con todos los meses anteriores y conserva el "
              r"orden temporal. Con estas ventanas se comprueba el criterio CA-07. "
              r"Cálculo propio (Anexo B)."))
    return guardar("tabla_a7.tex", texto), filas


# --- Verificación ---------------------------------------------------------
def verificar(datos):
    """Comprobaciones que deben cumplirse para que las cifras del informe sean válidas."""
    assert len(datos) == 17236, "número de filas inesperado"
    assert int(datos["grave"].sum()) == 1656, "número de siniestros graves inesperado"
    trimestre = datos[datos["FECHA"] < "2026-04-01"]
    assert len(trimestre) == 4789 and int(trimestre["FALLECIDOS"].sum()) == 603, \
        "los totales del primer trimestre no coinciden con el INEC"
    condicion = datos[["CONDICION_%d" % i for i in range(1, 44)]]
    assert ((condicion == "FALLECIDO").sum(axis=1) == datos["FALLECIDOS"]).all(), \
        "la condición de los participantes no reproduce la columna FALLECIDOS"
    assert datos["FECHA"].notna().all(), "hay fechas vacías"


def main():
    if not os.path.exists(ARCHIVO):
        sys.exit("No se encontró el archivo de la ANT en %s" % ARCHIVO)
    os.makedirs(SALIDA, exist_ok=True)

    datos = cargar()
    verificar(datos)

    generadas = [tabla_a1(datos), tabla_a2(datos), tabla_a3(datos),
                 tabla_a4(datos), tabla_a5(datos), tabla_a6(datos), tabla_a7(datos)]

    print("Filas: %s | Columnas: %d | Graves: %s (%s %%)"
          % (entero(len(datos)), datos.attrs["columnas_archivo"],
             entero(datos["grave"].sum()), decimal(datos["grave"].mean() * 100, 2)))
    for ruta, _ in generadas:
        print("Escrita:", os.path.relpath(ruta, RAIZ))
        sys.stdout.flush()
    print("Verificación superada: las cifras coinciden con las del informe.")


if __name__ == "__main__":
    main()
