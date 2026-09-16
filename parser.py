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



def _extraer_bloque_region_tecnologia(df, titulo, ronda, campo, fin_titulo=None, max_filas=40):
    """Extrae una matriz región × tecnología × equipo desde una sección de CESIM."""
    equipos = get_equipos(df)
    if not equipos:
        return pd.DataFrame()

    inicio = _fila_por_texto(df, titulo)
    if inicio is None:
        return pd.DataFrame()

    if fin_titulo:
        fin = _fila_por_texto(df, fin_titulo, inicio=inicio + 1)
    else:
        fin = None
    fin = fin if fin is not None else min(inicio + max_filas, len(df))

    registros = []
    region_actual = None

    for i in range(inicio + 1, fin):
        c0 = df.iloc[i, 0]

        if c0 in ["EE.UU.", "Asia"]:
            region_actual = c0
            continue

        if c0 in TECNOLOGIAS and region_actual:
            for j, equipo in enumerate(equipos, start=1):
                if j >= df.shape[1]:
                    continue
                valor = pd.to_numeric(df.iloc[i, j], errors="coerce")
                if pd.notna(valor):
                    registros.append({
                        "ronda": ronda,
                        "region_produccion": region_actual,
                        "tecnologia": c0,
                        "equipo": equipo,
                        campo: float(valor),
                    })

    return pd.DataFrame(registros)


def extraer_produccion_costos(df, ronda):
    """
    Producción interna/contratada y costo unitario correspondiente.
    Producción: miles de unidades.
    Costo unitario: USD por unidad.
    Costo total calculado: miles USD.
    """
    prod_int = _extraer_bloque_region_tecnologia(
        df,
        "Producción interna, miles unidades",
        ronda,
        "produccion_interna",
        fin_titulo="Producción contratada, miles unidades",
        max_filas=20,
    )

    prod_con = _extraer_bloque_region_tecnologia(
        df,
        "Producción contratada, miles unidades",
        ronda,
        "produccion_contratada",
        fin_titulo="Capacidad empleada, %",
        max_filas=20,
    )

    costo_int = _extraer_bloque_region_tecnologia(
        df,
        "Costo de producción interna por unidad, USD",
        ronda,
        "costo_unit_interno",
        fin_titulo="Costos de fabricación contratada por unidad, USD",
        max_filas=20,
    )

    costo_con = _extraer_bloque_region_tecnologia(
        df,
        "Costos de fabricación contratada por unidad, USD",
        ronda,
        "costo_unit_contratado",
        fin_titulo="Costo unitario promedio por producto vendido, USD",
        max_filas=20,
    )

    claves = ["ronda", "region_produccion", "tecnologia", "equipo"]
    frames = [x for x in [prod_int, prod_con, costo_int, costo_con] if not x.empty]

    if not frames:
        return pd.DataFrame()

    resultado = frames[0]
    for frame in frames[1:]:
        resultado = resultado.merge(frame, on=claves, how="outer")

    for col in [
        "produccion_interna",
        "produccion_contratada",
        "costo_unit_interno",
        "costo_unit_contratado",
    ]:
        if col not in resultado:
            resultado[col] = pd.NA
        resultado[col] = pd.to_numeric(resultado[col], errors="coerce")

    resultado["produccion_interna"] = resultado["produccion_interna"].fillna(0)
    resultado["produccion_contratada"] = resultado["produccion_contratada"].fillna(0)
    resultado["produccion_total"] = (
        resultado["produccion_interna"] + resultado["produccion_contratada"]
    )

    # miles unidades × USD/unidad = miles USD
    resultado["costo_interno_miles_usd"] = (
        resultado["produccion_interna"] * resultado["costo_unit_interno"]
    )
    resultado["costo_contratado_miles_usd"] = (
        resultado["produccion_contratada"] * resultado["costo_unit_contratado"]
    )
    resultado["costo_fabricacion_miles_usd"] = (
        resultado[["costo_interno_miles_usd", "costo_contratado_miles_usd"]]
        .sum(axis=1, min_count=1)
    )

    resultado["costo_unit_promedio_produccion"] = (
        resultado["costo_fabricacion_miles_usd"]
        / resultado["produccion_total"].replace(0, pd.NA)
    )

    return resultado


def extraer_stock_final(df, ronda):
    """
    Inventario final reportado en Detalles de logística.
    Stock en miles de unidades, por tecnología, ubicación y equipo.
    """
    equipos = get_equipos(df)
    if not equipos:
        return pd.DataFrame()

    inicio_log = _fila_por_texto(df, "Detalles de logística")
    if inicio_log is None:
        return pd.DataFrame()

    registros = []

    for tech_num, tech in enumerate(TECNOLOGIAS, start=1):
        inicio = _fila_por_texto(
            df, f"{tech}, miles unidades",
            inicio=inicio_log
        )
        if inicio is None:
            continue

        if tech_num < 4:
            fin = _fila_por_texto(
                df, f"Tec {tech_num + 1}, miles unidades",
                inicio=inicio + 1
            )
        else:
            fin = _fila_por_texto(
                df, "Origen de productos vendidos en, miles unidades",
                inicio=inicio + 1
            )
        fin = fin if fin is not None else min(inicio + 40, len(df))

        ubicacion_actual = None
        for i in range(inicio + 1, fin):
            c0 = df.iloc[i, 0]

            # Europa no almacena inventario en este reporte.
            if c0 in ["EE.UU.", "Asia"] and df.iloc[i, 1:].isna().all():
                ubicacion_actual = c0
                continue
            if c0 == "Europa":
                ubicacion_actual = None
                continue

            if isinstance(c0, str) and c0.strip() == "Inventario final" and ubicacion_actual:
                for j, equipo in enumerate(equipos, start=1):
                    if j >= df.shape[1]:
                        continue
                    valor = pd.to_numeric(df.iloc[i, j], errors="coerce")
                    if pd.notna(valor):
                        registros.append({
                            "ronda": ronda,
                            "ubicacion": ubicacion_actual,
                            "tecnologia": tech,
                            "equipo": equipo,
                            "stock_final": float(valor),
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
        "produccion": extraer_produccion_costos(df, ronda),
        "stock": extraer_stock_final(df, ronda),
    }

# ===== V8: control de gestión por áreas =====
def _extraer_seccion_filas(df, titulo, labels, ronda, region=None, fin_titulo=None, max_filas=80):
    inicio = _fila_por_texto(df, titulo)
    if inicio is None:
        return pd.DataFrame()
    fin = _fila_por_texto(df, fin_titulo, inicio=inicio+1) if fin_titulo else None
    fin = fin if fin is not None else min(inicio + max_filas, len(df))
    _, cols = _header_equipos_cercano(df, inicio)
    if not cols:
        return pd.DataFrame()
    out=[]
    wanted={x.lower():x for x in labels}
    for i in range(inicio+1, fin):
        v=df.iloc[i,0]
        if not isinstance(v,str): continue
        key=v.strip().lower()
        if key in wanted:
            for col,eq in cols.items():
                val=df.iloc[i,col] if col<df.shape[1] else None
                out.append({'ronda':ronda,'region':region,'equipo':eq,'indicador':wanted[key],'valor':val})
    return pd.DataFrame(out)


def extraer_rrhh(df, ronda):
    inicio=_fila_por_texto(df,'Informe de RRHH')
    if inicio is None: return pd.DataFrame()
    _, cols=_header_equipos_cercano(df,inicio)
    labels={
        'Salario mensual, USD':'salario_mensual_usd',
        'Presupuesto mensual para capacitación, USD':'capacitacion_mensual_usd',
        'Multiplicador de la eficiencia de RRHH':'multiplicador_eficiencia',
        'Número de personal de I+D, esta ronda':'dotacion',
        'Rotación de personal, %':'rotacion_pct',
    }
    out=[]
    for label,campo in labels.items():
        fila=_fila_por_texto(df,label,inicio=inicio,fin=min(inicio+35,len(df)))
        if fila is None: continue
        for col,eq in cols.items():
            val=pd.to_numeric(df.iloc[fila,col],errors='coerce')
            if pd.notna(val): out.append({'ronda':ronda,'equipo':eq,'indicador':campo,'valor':float(val)})
    return pd.DataFrame(out)


def extraer_promocion(df, ronda):
    secciones=[('Global','Cuenta de resultados, miles USD, Global','Cuenta de resultados, miles USD, EE.UU.'),
               ('EE.UU.','Cuenta de resultados, miles USD, EE.UU.','Cuenta de resultados, miles USD, Asia'),
               ('Asia','Cuenta de resultados, miles USD, Asia','Cuenta de resultados, miles USD, Europa'),
               ('Europa','Cuenta de resultados, miles USD, Europa','Informe de mercado, EE.UU.')]
    out=[]
    for region,titulo,fin_t in secciones:
        inicio=_fila_por_texto(df,titulo)
        fin=_fila_por_texto(df,fin_t,inicio=inicio+1) if inicio is not None else None
        if inicio is None: continue
        fin=fin or min(inicio+100,len(df))
        _,cols=_header_equipos_cercano(df,inicio)
        fila=_fila_por_texto(df,'Promoción',inicio=inicio,fin=fin)
        if fila is None: continue
        for col,eq in cols.items():
            val=pd.to_numeric(df.iloc[fila,col],errors='coerce')
            if pd.notna(val): out.append({'ronda':ronda,'region':region,'equipo':eq,'promocion_miles_usd':float(val)})
    return pd.DataFrame(out)


def extraer_logistica_detalle(df, ronda):
    equipos=get_equipos(df); inicio_log=_fila_por_texto(df,'Detalles de logística')
    if not equipos or inicio_log is None: return pd.DataFrame()
    out=[]
    for n,tech in enumerate(TECNOLOGIAS,1):
        ini=_fila_por_texto(df,f'{tech}, miles unidades',inicio=inicio_log)
        if ini is None: continue
        fin=_fila_por_texto(df,f'Tec {n+1}, miles unidades',inicio=ini+1) if n<4 else _fila_por_texto(df,'Origen de productos vendidos en, miles unidades',inicio=ini+1)
        fin=fin or min(ini+40,len(df))
        region=None
        for i in range(ini+1,fin):
            c0=df.iloc[i,0]
            if c0 in ['EE.UU.','Asia','Europa'] and df.iloc[i,1:].isna().all(): region=c0; continue
            if not region or not isinstance(c0,str): continue
            label=c0.strip()
            mapa={'Inventario inicial':'stock_inicial','Producción interna':'produccion_interna','Producción contratada':'produccion_contratada','Total disponible':'disponible','Inventario final':'stock_final','Demanda insatisfecha':'demanda_insatisfecha'}
            if label.startswith('Ventas en '): campo='ventas'
            elif label in mapa: campo=mapa[label]
            else: continue
            for j,eq in enumerate(equipos,1):
                if j>=df.shape[1]: continue
                val=pd.to_numeric(df.iloc[i,j],errors='coerce')
                if pd.notna(val):
                    if campo=='ventas': val=abs(float(val))
                    out.append({'ronda':ronda,'region':region,'tecnologia':tech,'equipo':eq,'indicador':campo,'valor':float(val)})
    if not out: return pd.DataFrame()
    long=pd.DataFrame(out)
    wide=long.pivot_table(index=['ronda','region','tecnologia','equipo'],columns='indicador',values='valor',aggfunc='sum').reset_index()
    wide.columns.name=None
    for c in ['stock_inicial','produccion_interna','produccion_contratada','disponible','ventas','stock_final','demanda_insatisfecha']:
        if c not in wide: wide[c]=0.0
    return wide


def extraer_finanzas_control(df,ronda):
    equipos=get_equipos(df)
    if not equipos: return pd.DataFrame()
    # Global/balance/cash-flow values relevant to the defense.
    labels=[
        ('Deudas a largo plazo','deuda_largo_plazo'),
        ('Deudas a corto plazo (no planificadas)','deuda_corto_plazo'),
        ('Préstamos internos','prestamos_internos'),
        ('Inversiones en fábricas (-) / desinversiones (+)','inversion_desinversion_fabricas'),
        ('Cambio de deuda a largo plazo (incr+ / dism-)','cambio_deuda_lp'),
        ('Cambio en deuda a corto plazo (incr+ /dism - )','cambio_deuda_cp'),
    ]
    out=[]
    # Use first occurrence for consolidated/global items except factory investments: collect USA + Asia separately below.
    for label,campo in labels[:3]+labels[4:]:
        vals=_extraer_fila_unica(df,label)
        for eq,val in vals.items():
            if pd.notna(val): out.append({'ronda':ronda,'region':'Global','equipo':eq,'indicador':campo,'valor':float(val)})
    # Factory investments/desinvestments by cash-flow section.
    for region,titulo,fin_t in [('EE.UU.','Flujo de efectivo de casa matriz, miles USD','Cuenta de resultados, miles USD, Asia'),('Asia','Estado de flujo de efectivo, miles USD, Asia','Cuenta de resultados, miles USD, Europa')]:
        ini=_fila_por_texto(df,titulo)
        if ini is None: continue
        fin=_fila_por_texto(df,fin_t,inicio=ini+1) or min(ini+120,len(df))
        _,cols=_header_equipos_cercano(df,ini)
        fila=_fila_por_texto(df,'Inversiones en fábricas (-) / desinversiones (+)',inicio=ini,fin=fin)
        if fila is not None:
            for col,eq in cols.items():
                val=pd.to_numeric(df.iloc[fila,col],errors='coerce')
                if pd.notna(val): out.append({'ronda':ronda,'region':region,'equipo':eq,'indicador':'inversion_desinversion_fabricas','valor':float(val)})
    return pd.DataFrame(out)


def extraer_logistica_financiera(df,ronda):
    out=[]
    secciones=[('Global','Cuenta de resultados, miles USD, Global','Cuenta de resultados, miles USD, EE.UU.'),('EE.UU.','Cuenta de resultados, miles USD, EE.UU.','Cuenta de resultados, miles USD, Asia'),('Asia','Cuenta de resultados, miles USD, Asia','Cuenta de resultados, miles USD, Europa'),('Europa','Cuenta de resultados, miles USD, Europa','Informe de mercado, EE.UU.')]
    for region,titulo,fin_t in secciones:
        ini=_fila_por_texto(df,titulo)
        if ini is None: continue
        fin=_fila_por_texto(df,fin_t,inicio=ini+1) or min(ini+100,len(df))
        _,cols=_header_equipos_cercano(df,ini)
        for label,campo in [('Costos de transporte y aranceles','costos_transporte_aranceles'),('de transferencias internas','transferencias_internas')]:
            fila=_fila_por_texto(df,label,inicio=ini,fin=fin)
            if fila is None: continue
            for col,eq in cols.items():
                val=pd.to_numeric(df.iloc[fila,col],errors='coerce')
                if pd.notna(val): out.append({'ronda':ronda,'region':region,'equipo':eq,'indicador':campo,'valor':float(val)})
    return pd.DataFrame(out)


# Replace/extend the public processor with V8 datasets.
def procesar_archivo(nombre, contenido):
    df = load_results_sheet(io.BytesIO(contenido))
    ronda = extraer_ronda(nombre, df)
    return {
        'ronda': ronda,
        'mercado': extraer_mercado(df, ronda),
        'market_share': extraer_market_share(df, ronda),
        'finanzas': extraer_finanzas(df, ronda),
        'demanda_insatisfecha': extraer_demanda_insatisfecha_logistica(df, ronda),
        'produccion': extraer_produccion_costos(df, ronda),
        'stock': extraer_stock_final(df, ronda),
        'logistica_detalle': extraer_logistica_detalle(df,ronda),
        'rrhh': extraer_rrhh(df,ronda),
        'promocion': extraer_promocion(df,ronda),
        'finanzas_control': extraer_finanzas_control(df,ronda),
        'logistica_financiera': extraer_logistica_financiera(df,ronda),
    }
