import os
import gc
import duckdb
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)
app.secret_key = "clave_secreta_analisis_app"

USUARIO_CORRECTO = "DICKSON"
PASSWORD_CORRECTO = "1234"

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

DATA_STORE = {}

@app.route("/")
def inicio():
    if session.get("usuario"):
        return redirect(url_for("cargar_excel"))
    return redirect(url_for("login"))

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        usuario_ingresado = request.form.get("username", "")
        password_ingresado = request.form.get("password", "")

        if usuario_ingresado == USUARIO_CORRECTO and password_ingresado == PASSWORD_CORRECTO:
            session["usuario"] = usuario_ingresado
            return redirect(url_for("cargar_excel"))
        else:
            error = "Usuario o contraseña incorrectos."
    
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.pop("usuario", None)
    return redirect(url_for("login"))

@app.route("/cargar-excel", methods=["GET", "POST"])
def cargar_excel():
    if not session.get("usuario"): 
        return redirect(url_for("login"))
    
    error = None
    user_key = session.get("usuario", "DICKSON")

    if request.method == "POST":
        if "archivo_excel" in request.files:
            file = request.files["archivo_excel"]
            if file.filename == "":
                error = "Nombre de archivo no válido."
            elif file and (file.filename.endswith(".xlsx") or file.filename.endswith(".xls")):
                try:
                    file_path = os.path.join(UPLOAD_FOLDER, f"{user_key}_{file.filename}")
                    file.save(file_path)

                    engine_type = "calamine" if file.filename.endswith(".xlsx") else "xlrd"
                    try:
                        df_preview = pd.read_excel(file_path, engine=engine_type, nrows=5)
                    except Exception:
                        df_preview = pd.read_excel(file_path, nrows=5)

                    columnas = [str(col).strip() for col in df_preview.columns]

                    DATA_STORE[user_key] = {
                        'file_path': file_path,
                        'filename': file.filename,
                        'columnas': columnas
                    }
                    
                    del df_preview
                    gc.collect()
                    
                    return redirect(url_for("dashboard"))

                except Exception as e:
                    error = f"Error al leer la estructura del Excel: {str(e)}"
            else:
                error = "Por favor, sube un archivo con extensión .xlsx o .xls."

    return render_template("cargar_excel.html", error=error)

@app.route("/dashboard")
def dashboard():
    if not session.get("usuario"):
        return redirect(url_for("login"))

    user_key = session.get("usuario", "DICKSON")
    if user_key not in DATA_STORE:
        return redirect(url_for("cargar_excel"))

    file_path = DATA_STORE[user_key]['file_path']
    filename = DATA_STORE[user_key]['filename']

    try:
        engine_type = "calamine" if filename.endswith(".xlsx") else "xlrd"
        try:
            df = pd.read_excel(file_path, engine=engine_type)
        except Exception:
            df = pd.read_excel(file_path)
        
        df.columns = [str(col).strip() for col in df.columns]

    except Exception as e:
        return f"Error al procesar el archivo Excel: {str(e)}"

    # Mapeo flexible de columnas
    cols_existentes = {col.upper().strip(): col for col in df.columns}
    
    def buscar_col(lista_opciones):
        for opcion in lista_opciones:
            for k, v in cols_existentes.items():
                if opcion.upper() in k:
                    return v
        return None

    col_gln_real = buscar_col(["CANT GLN", "CANTIDAD", "GALON", "GLN"]) or "CANT GLN"
    col_desc_real = buscar_col(["DESC X GALON", "DESCUENTO", "DESC"]) or "DESC X GALON"
    col_obsv_real = buscar_col(["OBSERVACION", "OBSERVACIONES", "OBSV", "ADICIONAL"]) or "OBSV ADICIONAL"
    col_cat1_real = buscar_col(["ANALISTA", "USUARIO", "RESPONSABLE"]) or "ANALISTA"
    col_fecha_real = buscar_col(["MES ENVIO", "FECHA", "MES"]) or "MES ENVIO"
    col_cc_real = buscar_col(["CENTRO DE COSTO", "CENTRO COSTO", "CENTRO_COSTO", "CC", "COSTO"]) or "CENTRO DE COSTO"

    # Asegurar existencia de columnas y limpieza numérica
    for col in [col_gln_real, col_desc_real, col_obsv_real, col_cat1_real, col_fecha_real, col_cc_real]:
        if col not in df.columns:
            df[col] = None

    for col in [col_gln_real, col_desc_real, col_obsv_real]:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.replace('$', '', regex=False)
            df[col] = df[col].str.replace(',', '', regex=False)
            df[col] = df[col].str.strip()

    mapeo = {
        'col_gln': col_gln_real,
        'col_desc_galon': col_desc_real,
        'col_obsv': col_obsv_real,
        'categoria_1': col_cat1_real,
        'fecha': col_fecha_real,
        'centro_costo': col_cc_real
    }
    DATA_STORE[user_key]['mapeo_columnas'] = mapeo

    con = duckdb.connect()
    con.register("tabla_excel", df)

    col_gln = f'"{col_gln_real}"'
    col_desc_galon = f'"{col_desc_real}"'
    col_obsv = f'"{col_obsv_real}"'
    col_cat1 = f'"{col_cat1_real}"'
    col_fecha = f'"{col_fecha_real}"'
    col_cc = f'"{col_cc_real}"'

    kpis = {
        'total_gln': "0.00",
        'total_desc_galon': "0.00",
        'total_obsv': "0.00",
        'gran_total_desc': "0.00",
        'total_registros': "0"
    }

    # 1. KPIs Generales
    try:
        query_kpis = f"""
            SELECT 
                COALESCE(SUM(TRY_CAST({col_gln} AS DOUBLE)), 0) as total_gln,
                COALESCE(SUM(TRY_CAST({col_desc_galon} AS DOUBLE)), 0) as total_desc_galon,
                COALESCE(SUM(TRY_CAST({col_obsv} AS DOUBLE)), 0) as total_obsv,
                (COALESCE(SUM(TRY_CAST({col_desc_galon} AS DOUBLE)), 0) + COALESCE(SUM(TRY_CAST({col_obsv} AS DOUBLE)), 0)) as gran_total_desc,
                COUNT(*) as total_registros
            FROM tabla_excel
        """
        res_kpis = con.execute(query_kpis).fetchone()
        if res_kpis:
            kpis = {
                'total_gln': f"{res_kpis[0]:,.2f}",
                'total_desc_galon': f"{res_kpis[1]:,.2f}",
                'total_obsv': f"{res_kpis[2]:,.2f}",
                'gran_total_desc': f"{res_kpis[3]:,.2f}",
                'total_registros': f"{res_kpis[4]:,}"
            }
    except Exception:
        pass

    # 2. Resumen por ANALISTA y MES ENVIO
    resumen_tabla = []
    q_tabla = f"""
        SELECT 
            CAST({col_cat1} AS VARCHAR) as analista,
            SUBSTRING(CAST({col_fecha} AS VARCHAR), 1, 7) as mes_envio,
            COALESCE(SUM(TRY_CAST({col_gln} AS DOUBLE)), 0) as total_gln,
            COALESCE(SUM(TRY_CAST({col_desc_galon} AS DOUBLE)), 0) as sum_desc_galon,
            COALESCE(SUM(TRY_CAST({col_obsv} AS DOUBLE)), 0) as sum_obsv,
            (COALESCE(SUM(TRY_CAST({col_desc_galon} AS DOUBLE)), 0) + COALESCE(SUM(TRY_CAST({col_obsv} AS DOUBLE)), 0)) as total_descuento
        FROM tabla_excel
        GROUP BY analista, mes_envio
        ORDER BY mes_envio ASC, analista ASC
    """
    try:
        resumen_tabla = con.execute(q_tabla).fetchall()
    except Exception:
        resumen_tabla = []

    # 3. Gráfico 1: Analistas vs Galones
    q_analistas = f"""
        SELECT 
            CAST({col_cat1} AS VARCHAR) as analista,
            SUBSTRING(CAST({col_fecha} AS VARCHAR), 1, 7) as mes_envio,
            COALESCE(SUM(TRY_CAST({col_gln} AS DOUBLE)), 0) as total_gln,
            COALESCE(SUM(TRY_CAST({col_desc_galon} AS DOUBLE)), 0) as sum_desc_galon,
            COALESCE(SUM(TRY_CAST({col_obsv} AS DOUBLE)), 0) as sum_obsv,
            (COALESCE(SUM(TRY_CAST({col_desc_galon} AS DOUBLE)), 0) + COALESCE(SUM(TRY_CAST({col_obsv} AS DOUBLE)), 0)) as total_descuento
        FROM tabla_excel
        WHERE {col_cat1} IS NOT NULL
        GROUP BY analista, mes_envio
        ORDER BY mes_envio ASC, analista ASC
    """
    
    chart_labels = []
    chart_gln = []
    chart_desc = []
    chart_obsv = []
    chart_total = []
    chart_colors = []
    datos_analistas = {}
    total_gln_general = 0.0

    try:
        res_analistas = con.execute(q_analistas).fetchall()
        for r in res_analistas:
            nombre = str(r[0]).strip().upper() if r[0] is not None else "N/A"
            mes = str(r[1]).strip() if r[1] is not None else "N/A"
            gln_val = float(r[2]) if r[2] else 0.0

            chart_labels.append(f"{nombre} - {mes}")
            chart_gln.append(gln_val)
            chart_desc.append(float(r[3]) if r[3] else 0.0)
            chart_obsv.append(float(r[4]) if r[4] else 0.0)
            chart_total.append(float(r[5]) if r[5] else 0.0)

            if nombre not in datos_analistas:
                datos_analistas[nombre] = {}
            datos_analistas[nombre][mes] = gln_val
            total_gln_general += gln_val

            if "DICKSON" in nombre:
                chart_colors.append("#0d6efd")  # Azul
            elif "FABIAN" in nombre or "FABIÁN" in nombre:
                chart_colors.append("#dc3545")  # Rojo
            else:
                chart_colors.append("#6c757d")  # Gris
    except Exception:
        pass

    # Análisis explicativo Gráfico 1
    analisis_grafico1 = "No hay información suficiente para procesar el análisis de analistas."
    if datos_analistas and total_gln_general > 0:
        totales_por_analista = {analista: sum(meses.values()) for analista, meses in datos_analistas.items()}
        lider = max(totales_por_analista, key=totales_por_analista.get)
        lider_gln = totales_por_analista[lider]
        lider_pct = (lider_gln / total_gln_general) * 100

        segundo = min(totales_por_analista, key=totales_por_analista.get)
        segundo_gln = totales_por_analista[segundo]
        segundo_pct = (segundo_gln / total_gln_general) * 100

        diferencia = abs(lider_gln - segundo_gln)

        meses_totales = {}
        for analista, meses in datos_analistas.items():
            for m, v in meses.items():
                meses_totales[m] = meses_totales.get(m, 0.0) + v
        mes_pico = max(meses_totales, key=meses_totales.get) if meses_totales else "N/A"

        analisis_grafico1 = (
            f"<strong>Conclusión y Análisis del Volumen Operativo:</strong><br>"
            f"• <strong>Liderazgo de Volumen:</strong> {lider} concentra la mayor carga operativa con <strong>{lider_gln:,.2f} GLN</strong> (<strong>{lider_pct:.1f}%</strong> del total).<br>"
            f"• <strong>Comparativo Operativo:</strong> {segundo} registró <strong>{segundo_gln:,.2f} GLN</strong> (<strong>{segundo_pct:.1f}%</strong>), existiendo una brecha de <strong>{diferencia:,.2f} GLN</strong>.<br>"
            f"• <strong>Mes Pico:</strong> El periodo de mayor concentración de volumen fue <strong>{mes_pico}</strong> con un total de <strong>{meses_totales.get(mes_pico, 0):,.2f} GLN</strong>.<br>"
            f"• <strong>Recomendación:</strong> Monitorear la carga en picos operativos para balancear revisiones de galonaje."
        )

    # 4. Gráfico 2: Frecuencia de Transacciones por Centro de Costo y Mes (Barras Apiladas)
    chart_cc_labels = []
    chart_cc_datasets = []
    analisis_grafico2 = "No hay información de Centros de Costo para mostrar."
    total_tx_general = 0

    try:
        q_cc = f"""
            SELECT 
                SUBSTRING(CAST({col_fecha} AS VARCHAR), 1, 7) as mes_envio,
                CAST({col_cc} AS VARCHAR) as centro_costo,
                COUNT(*) as total_transacciones
            FROM tabla_excel
            WHERE {col_cc} IS NOT NULL AND {col_cc} != '' AND {col_cc} != 'None'
            GROUP BY mes_envio, centro_costo
            ORDER BY mes_envio ASC, centro_costo ASC
        """
        res_cc = con.execute(q_cc).fetchall()
        
        if res_cc:
            # Obtener meses únicos (labels en eje X)
            meses_unicos = sorted(list(set(str(r[0]).strip() for r in res_cc if r[0])))
            chart_cc_labels = meses_unicos

            # Agrupar conteos por Centro de Costo
            dict_cc = {}
            for r in res_cc:
                mes = str(r[0]).strip()
                cc = str(r[1]).strip()
                count = int(r[2])
                
                if cc not in dict_cc:
                    dict_cc[cc] = {m: 0 for m in meses_unicos}
                dict_cc[cc][mes] = count
                total_tx_general += count

            # Paleta de colores para los centros de costo
            colores_cc = ["#0d6efd", "#198754", "#ffc107", "#0dcaf0", "#6f42c1", "#fd7e14", "#d63384", "#20c997"]
            
            idx = 0
            for cc, valores in dict_cc.items():
                chart_cc_datasets.append({
                    "label": f"CC: {cc}",
                    "data": [valores[m] for m in meses_unicos],
                    "backgroundColor": colores_cc[idx % len(colores_cc)]
                })
                idx += 1

            # Generar análisis estadístico dinámico para Gráfico 2
            totales_por_cc = {cc: sum(m.values()) for cc, m in dict_cc.items()}
            top_cc = max(totales_por_cc, key=totales_por_cc.get)
            top_cc_tx = totales_por_cc[top_cc]
            top_cc_pct = (top_cc_tx / total_tx_general) * 100 if total_tx_general > 0 else 0

            analisis_grafico2 = (
                f"<strong>Conclusión y Análisis de Frecuencia Operativa por Centro de Costos:</strong><br>"
                f"• <strong>Centro de Costo Mayoritario:</strong> El Centro de Costos <strong>{top_cc}</strong> lidera la frecuencia transaccional con <strong>{top_cc_tx:,} envíos/registros</strong> (<strong>{top_cc_pct:.1f}%</strong> del total transaccional).<br>"
                f"• <strong>Volumen Transaccional Total:</strong> Se procesaron un total de <strong>{total_tx_general:,} operaciones</strong> distribuidas en los periodos analizados.<br>"
                f"• <strong>Distribución Temporal:</strong> Las solicitudes por Centro de Costos muestran un comportamiento de concentración mensual reflejado en las barras apiladas.<br>"
                f"• <strong>Recomendación:</strong> Priorizar la auditoría de procesos y autorizaciones en el centro de mayor frecuencia (CC {top_cc})."
            )
    except Exception:
        pass

    # 5. Tarjeta 3: Análisis Estadístico Comparativo Cruzado
    analisis_cruzado = None
    if datos_analistas and total_tx_general > 0:
        promedio_gln_por_tx = total_gln_general / total_tx_general if total_tx_general > 0 else 0.0
        
        analisis_cruzado = (
            f"<strong>Diagnóstico Comparativo Cruzado (Volumen vs Frecuencia):</strong><br>"
            f"• <strong>Eficiencia Transaccional (Galonaje Promedio por Operación):</strong> En promedio, cada transacción registrada en los centros de costo moviliza <strong>{promedio_gln_por_tx:,.2f} GLN</strong>.<br>"
            f"• <strong>Correlación Operativa:</strong> Mientras el total acumulado alcanza <strong>{total_gln_general:,.2f} GLN</strong> en volumen, la operatividad implicó <strong>{total_tx_general:,} transacciones</strong> de gestión por parte del equipo analista.<br>"
            f"• <strong>Conclusión Estratégica:</strong> Un alto volumen de transacciones concentrado en un solo centro de costo puede requerir mayor dedicación operativa que grandes volúmenes de galonaje agrupados en pocos registros."
        )

    del df
    gc.collect()

    return render_template(
        "dashboard.html",
        kpis=kpis,
        registros=resumen_tabla,
        filename=filename,
        mapeo=mapeo,
        # Tarjeta 1
        chart_labels=chart_labels,
        chart_gln=chart_gln,
        chart_desc=chart_desc,
        chart_obsv=chart_obsv,
        chart_total=chart_total,
        chart_colors=chart_colors,
        analisis_grafico1=analisis_grafico1,
        # Tarjeta 2
        chart_cc_labels=chart_cc_labels,
        chart_cc_datasets=chart_cc_datasets,
        analisis_grafico2=analisis_grafico2,
        # Tarjeta 3 Cruzada
        analisis_cruzado=analisis_cruzado
    )

if __name__ == "__main__":
    app.run(debug=True)