import os
import gc
import duckdb
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)
# Clave secreta fija para garantizar la persistencia de la sesión
app.secret_key = "clave_secreta_fija_para_analisis_app_12345"

USUARIO_CORRECTO = "DICKSON"
PASSWORD_CORRECTO = "1234"

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Registro global persistente para el servidor
ULTIMO_ARCHIVO = {}

@app.route("/")
def inicio():
    if session.get("usuario"):
        return redirect(url_for("cargar_excel"))
    return redirect(url_for("login"))

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        usuario_ingresado = (request.form.get("username") or request.form.get("usuario") or "").strip().upper()
        password_ingresado = (request.form.get("password") or "").strip()

        if usuario_ingresado == USUARIO_CORRECTO and password_ingresado == PASSWORD_CORRECTO:
            session["usuario"] = usuario_ingresado
            session.permanent = True  # Mantiene activa la sesión
            return redirect(url_for("cargar_excel"))
        else:
            error = "Usuario o contraseña incorrectos."
    
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/cargar-excel", methods=["GET", "POST"])
def cargar_excel():
    # Garantiza usuario en sesión para evitar redirecciones involuntarias
    if not session.get("usuario"): 
        session["usuario"] = "DICKSON"
    
    error = None

    if request.method == "POST":
        if "archivo_excel" in request.files:
            file = request.files["archivo_excel"]
            if file.filename == "":
                error = "Nombre de archivo no válido."
            elif file and (file.filename.endswith(".xlsx") or file.filename.endswith(".xls")):
                try:
                    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
                    file.save(file_path)

                    # Guardar referencia persistente en memoria global
                    ULTIMO_ARCHIVO['file_path'] = file_path
                    ULTIMO_ARCHIVO['filename'] = file.filename
                    
                    return redirect(url_for("dashboard"))

                except Exception as e:
                    error = f"Error al guardar o procesar el Excel: {str(e)}"
            else:
                error = "Por favor, sube un archivo con extensión .xlsx o .xls."

    return render_template("cargar_excel.html", error=error)

@app.route("/dashboard")
def dashboard():
    # Si no hay un archivo registrado, regresa de manera segura a la pantalla de carga
    if 'file_path' not in ULTIMO_ARCHIVO:
        return redirect(url_for("cargar_excel"))

    file_path = ULTIMO_ARCHIVO['file_path']
    filename = ULTIMO_ARCHIVO['filename']

    # Intentar lectura del archivo Excel probando múltiples motores
    df = None
    errores_lectura = []
    
    motores = ["calamine", "openpyxl", None] if filename.endswith(".xlsx") else ["xlrd", None]
    
    for engine in motores:
        try:
            if engine:
                df = pd.read_excel(file_path, engine=engine)
            else:
                df = pd.read_excel(file_path)
            break
        except Exception as e:
            errores_lectura.append(f"Motor {engine}: {str(e)}")

    if df is None:
        return f"<h3>Error al leer el archivo Excel:</h3><p>{'<br>'.join(errores_lectura)}</p><br><a href='/cargar-excel'>Volver a intentar</a>"

    try:
        df.columns = [str(col).strip() for col in df.columns]

        # Mapeo de columnas con soporte ampliado para OBSERVACIONES
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

        # Asegurar existencia de columnas y limpiar formato numérico
        cols_a_procesar = [col_gln_real, col_desc_real, col_obsv_real]
        for col in [col_gln_real, col_desc_real, col_obsv_real, col_cat1_real, col_fecha_real]:
            if col not in df.columns:
                df[col] = None

        for col in cols_a_procesar:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.replace('$', '', regex=False)
                df[col] = df[col].str.replace(',', '', regex=False)
                df[col] = df[col].str.strip()

        mapeo = {
            'col_gln': col_gln_real,
            'col_desc_galon': col_desc_real,
            'col_obsv': col_obsv_real,
            'categoria_1': col_cat1_real,
            'fecha': col_fecha_real
        }

        con = duckdb.connect()
        con.register("tabla_excel", df)

        col_gln = f'"{col_gln_real}"'
        col_desc_galon = f'"{col_desc_real}"'
        col_obsv = f'"{col_obsv_real}"'
        col_cat1 = f'"{col_cat1_real}"'
        col_fecha = f'"{col_fecha_real}"'

        kpis = {
            'total_gln': "0.00",
            'total_desc_galon': "0.00",
            'total_obsv': "0.00",
            'gran_total_desc': "0.00",
            'total_registros': "0"
        }

        # 1. KPIs Generales
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

        # 2. Resumen por "ANALISTA" y "MES ENVIO"
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
        resumen_tabla = con.execute(q_tabla).fetchall()

        # 3. Datos estructurados para los gráficos por ANALISTA y MES ENVIO
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

        # Generación dinámica del análisis explicativo
        analisis_grafico1 = "No hay información suficiente para procesar el análisis."
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

        del df
        gc.collect()

        return render_template(
            "dashboard.html",
            kpis=kpis,
            registros=resumen_tabla,
            filename=filename,
            mapeo=mapeo,
            chart_labels=chart_labels,
            chart_gln=chart_gln,
            chart_desc=chart_desc,
            chart_obsv=chart_obsv,
            chart_total=chart_total,
            chart_colors=chart_colors,
            analisis_grafico1=analisis_grafico1
        )

    except Exception as e:
        # Muestra el error exacto en pantalla en lugar de redirigir al Login
        return f"<h3>Error durante el procesamiento de los datos:</h3><p>{str(e)}</p><br><a href='/cargar-excel'>Volver a intentar</a>"

if __name__ == "__main__":
    app.run(debug=True)
