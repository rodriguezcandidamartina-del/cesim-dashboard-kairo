import io
import re
import pandas as pd

REGIONES = {
    "EE.UU.": {"titulo": "Informe de mercado, EE.UU.", "moneda": "USD"},
    "Asia": {"titulo": "Informe de mercado, Asia", "moneda": "RMB"},
    "Europa": {"titulo": "Informe de mercado, Europa", "moneda": "EUR"},
}
TECNOLOGIAS = ["Tec 1", "Tec 2", "Tec 3", "Tec 4"]


def load_results_sheet(source, sheet_name="Results"):
    """Lee la hoja Results de un .xls o .xlsx."""
    return pd.read_excel(source, sheet_name=sheet_name, header=None)


def extraer_ronda(nombre, df=None):
    m = re.search(r"[rR](?:onda)?[_\-\s]?0*(\d+)", str(nombre))
    if m:
        return int(m.group(1))
    if df is not None and len(df):
        c0 = str(df.iloc[0, 0])
        m = re.search(r"Ronda\s+(\d+)", c0, re.I)
        if m:
            return int(m.group(1))
    m = re.search(r"(\d+)", str(nombre))
    return int(m.group(1)) if m else -1


def _fila_por_texto(df, texto, inicio=0, fin=None, exacto=True):
    fin = len(df) if fin is None else min(fin, len(df))
    objetivo = texto.strip().lower()
    for i in range(inicio, fin):
        v = df.iloc[i, 0]
        if not isinstance(v, str):
            continue
        actual = v.strip().lower()
        if (exacto and actual == objetivo) or ((not exacto) and objetivo in actual):
            return i
    return None


def _header_equipos_cercano(df, fila_titulo, max_busqueda=12):
    for i in range(fila_titulo + 1, min(fila_titulo + 1 + max_busqueda, len(df))):
        vals = []
        for j in range(1, df.shape[1]):
            v = df.iloc[i, j]
            if isinstance(v, str) and v.strip():
                vals.append((j, v.strip()))
        if len(vals) >= 3:
            return i, dict(vals)
    return None, {}


def get_equipos(df):
    # El encabezado global aparece al principio del archivo.
    for i in range(min(40, len(df))):
        vals = [str(v).strip() for v in df.iloc[i, 1:].tolist()
                if isinstance(v, str) and str(v).strip()]
        if len(vals) >= 3:
            return vals
    return []


def extraer_mercado(df, ronda):
    """Precio, características, marketing, ventas y demanda por región/tec/equipo."""
    registros = []

    titulos = {}
    for region, meta in REGIONES.items():
        r = _fila_por_texto(df, meta["titulo"])
        if r is not None:
            titulos[region] = r

    orden = sorted(titulos.items(), key=lambda x: x[1])
    for pos, (region, inicio) in enumerate(orden):
        fin = orden[pos + 1][1] if pos + 1 < len(orden) else min(inicio + 120, len(df))
        _, equipos_cols = _header_equipos_cercano(df, inicio)
        if not equipos_cols:
            continue

        for tech in TECNOLOGIAS:
            fila_tech = _fila_por_texto(df, tech, inicio=inicio, fin=fin)
            if fila_tech is None:
                continue

            # Cada bloque tecnológico tiene 5 filas de datos inmediatas.
            datos = {}
            for i in range(fila_tech + 1, min(fila_tech + 8, fin)):
                label = df.iloc[i, 0]
                if not isinstance(label, str):
                    continue
                lbl = label.strip()
                if lbl.startswith("Precio de venta"):
                    datos["precio"] = i
                elif lbl == "Cantidad de características ofrecidas":
                    datos["caracteristicas"] = i
                elif lbl == "Enfoque de la estrategia de marketing":
                    datos["marketing"] = i
                elif lbl == "Ventas, miles unidades":
                    datos["ventas"] = i
                elif lbl == "Demanda, miles unidades":
                    datos["demanda"] = i

            for col, equipo in equipos_cols.items():
                rec = {
                    "ronda": ronda, "region": region, "tecnologia": tech,
                    "equipo": equipo, "moneda": REGIONES[region]["moneda"],
                }
                for campo, fila in datos.items():
                    rec[campo] = df.iloc[fila, col] if col < df.shape[1] else None
                if any(pd.notna(rec.get(k)) for k in ["precio", "ventas", "demanda", "caracteristicas", "marketing"]):
                    registros.append(rec)

    result = pd.DataFrame(registros)
    if not result.empty:
        for c in ["precio", "caracteristicas", "ventas", "demanda"]:
            if c in result:
                result[c] = pd.to_numeric(result[c], errors="coerce")
        result["demanda_insatisfecha_calculada"] = (result["demanda"] - result["ventas"]).clip(lower=0)
        result["cobertura_demanda_pct"] = (
            result["ventas"] / result["demanda"].replace(0, pd.NA) * 100
        )
    return result



def extraer_market_share(df, ronda):
    """Cuotas por región, tecnología y total."""
    secciones = {
        "Global": "Cuotas de mercado globales, %",
        "EE.UU.": "EE.UU. cuotas de mercado, %",
        "Asia": "Asia cuotas de mercado, %",
        "Europa": "Europa cuotas de mercado, %",
    }

    equipos = get_equipos(df)
    if not equipos:
        return pd.DataFrame()

    equipos_cols = {j + 1: equipo for j, equipo in enumerate(equipos)}
    registros = []

    for region, titulo in secciones.items():
        fila = _fila_por_texto(df, titulo)
        if fila is None:
            continue

        for i in range(fila + 1, min(fila + 8, len(df))):
            etiqueta = df.iloc[i, 0]
            if etiqueta not in TECNOLOGIAS + ["Total"]:
                continue

            for col, equipo in equipos_cols.items():
                if col >= df.shape[1]:
                    continue
                valor = pd.to_numeric(df.iloc[i, col], errors="coerce")
                if pd.notna(valor):
                    registros.append({
                        "ronda": ronda,
                        "region": region,
                        "tecnologia": etiqueta,
                        "equipo": equipo,
                        "market_share_pct": float(valor),
                    })

    return pd.DataFrame(registros)

def _extraer_fila_global(df, label, occurrence=0):
    inicio = _fila_por_texto(df, "Cuenta de resultados, miles USD, Global")
    fin = _fila_por_texto(df, "Cuenta de resultados, miles USD, EE.UU.")
    if inicio is None:
        return {}
    fin = fin if fin is not None else min(inicio + 80, len(df))
    _, cols = _header_equipos_cercano(df, inicio)

    candidatos = []
    for i in range(inicio, fin):
        v = df.iloc[i, 0]
        if isinstance(v, str) and v.strip().lower() == label.strip().lower():
            if not df.iloc[i, 1:].isna().all():
                candidatos.append(i)
    if not candidatos or occurrence >= len(candidatos):
        return {}
    fila = candidatos[occurrence]
    return {eq: pd.to_numeric(df.iloc[fila, col], errors="coerce") for col, eq in cols.items()}


def _extraer_fila_unica(df, label):
    equipos = get_equipos(df)
    for i in range(len(df)):
        v = df.iloc[i, 0]
        if isinstance(v, str) and v.strip().lower() == label.strip().lower():
            vals = df.iloc[i, 1:1+len(equipos)].tolist()
            if any(pd.notna(x) for x in vals):
                return {eq: pd.to_numeric(val, errors="coerce")
                        for eq, val in zip(equipos, vals)}
    return {}


def extraer_finanzas(df, ronda):
    """KPIs financieros globales principales."""
    equipos = get_equipos(df)
    if not equipos:
        return pd.DataFrame()

    filas = {
        "Ingresos por ventas": _extraer_fila_global(df, "Ingresos por ventas"),
        "EBITDA": _extraer_fila_global(df, "Beneficio operativo antes de depreciación (EBITDA)"),
        "EBIT": _extraer_fila_global(df, "Beneficio operativo (EBIT)"),
        "Beneficio de la ronda": _extraer_fila_global(df, "Beneficio de la ronda"),
        "Activos totales": _extraer_fila_unica(df, "Activos Totales"),
        "Patrimonio neto": _extraer_fila_unica(df, "Total patrimonio neto"),
        "Deuda largo plazo": _extraer_fila_unica(df, "Deudas a largo plazo"),
        "ROCE, %": _extraer_fila_unica(df, "Rentabilidad del capital empleado (ROCE)"),
        "ROE, %": _extraer_fila_unica(df, "Rendimiento de los Fondos Propios (ROE)"),
        "Capitalización de mercado, miles USD": _extraer_fila_unica(df, "Capitalización de mercado de la empresa, miles USD"),
    }

    registros = []
    for kpi, valores in filas.items():
        for equipo, valor in valores.items():
            if pd.notna(valor):
                registros.append({"ronda": ronda, "equipo": equipo, "kpi": kpi, "valor": float(valor)})
    return pd.DataFrame(registros)


def extraer_demanda_insatisfecha_logistica(df, ronda):
    """
    Extrae la demanda insatisfecha de la sección logística por tecnología/región.
    Es preferible a demanda-ventas porque CESIM la reporta explícitamente.
    """
    equipos = get_equipos(df)
    if not equipos:
        return pd.DataFrame()

    registros = []
    for tech_num, tech in enumerate(TECNOLOGIAS, start=1):
        titulo = f"{tech}, miles unidades"
        inicio = _fila_por_texto(df, titulo)
        if inicio is None:
            continue
        # El siguiente bloque tecnológico está ~36 filas después.
        siguiente = _fila_por_texto(df, f"Tec {tech_num+1}, miles unidades", inicio=inicio+1) if tech_num < 4 else None
        fin = siguiente if siguiente is not None else min(inicio + 40, len(df))

        region_actual = None
        for i in range(inicio + 1, fin):
            c0 = df.iloc[i, 0]
            if c0 in ["EE.UU.", "Asia", "Europa"] and df.iloc[i, 1:].isna().all():
                region_actual = c0
                continue
            if isinstance(c0, str) and c0.strip() == "Demanda insatisfecha" and region_actual:
                for j, eq in enumerate(equipos, start=1):
                    if j >= df.shape[1]:
                        continue
                    val = pd.to_numeric(df.iloc[i, j], errors="coerce")
                    if pd.notna(val):
                        registros.append({
                            "ronda": ronda, "region": region_actual,
                            "tecnologia": tech, "equipo": eq,
                            "demanda_insatisfecha": float(val),
                        })
    return pd.DataFrame(registros)


def procesar_archivo(nombre, contenido):
    df = load_results_sheet(io.BytesIO(contenido))
    ronda = extraer_ronda(nombre, df)
    return {
        "ronda": ronda,
        "mercado": extraer_mercado(df, ronda),
        "market_share": extraer_market_share(df, ronda),
        "finanzas": extraer_finanzas(df, ronda),
        "demanda_insatisfecha": extraer_demanda_insatisfecha_logistica(df, ronda),
    }
