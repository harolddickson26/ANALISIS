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

    # Asegurar existencia de columnas y limpiar formato numérico (quitar $, comas, espacios)
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
    DATA_STORE[user_key]['mapeo_columnas'] = mapeo

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

    # 2. Resumen por "ANALISTA" y "MES ENVIO"
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

    # 3. Datos estructurados para los 4 gráficos por Analista
    q_analistas = f"""
        SELECT 
            CAST({col_cat1} AS VARCHAR) as analista,
            COALESCE(SUM(TRY_CAST({col_gln} AS DOUBLE)), 0) as total_gln,
            COALESCE(SUM(TRY_CAST({col_desc_galon} AS DOUBLE)), 0) as sum_desc_galon,
            COALESCE(SUM(TRY_CAST({col_obsv} AS DOUBLE)), 0) as sum_obsv,
            (COALESCE(SUM(TRY_CAST({col_desc_galon} AS DOUBLE)), 0) + COALESCE(SUM(TRY_CAST({col_obsv} AS DOUBLE)), 0)) as total_descuento
        FROM tabla_excel
        WHERE {col_cat1} IS NOT NULL
        GROUP BY analista
        ORDER BY total_descuento DESC
        LIMIT 10
    """
    
    chart_labels = []
    chart_gln = []
    chart_desc = []
    chart_obsv = []
    chart_total = []

    try:
        res_analistas = con.execute(q_analistas).fetchall()
        for r in res_analistas:
            chart_labels.append(str(r[0]) if r[0] is not None else "N/A")
            chart_gln.append(float(r[1]) if r[1] else 0.0)
            chart_desc.append(float(r[2]) if r[2] else 0.0)
            chart_obsv.append(float(r[3]) if r[3] else 0.0)
            chart_total.append(float(r[4]) if r[4] else 0.0)
    except Exception:
        pass

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
        chart_total=chart_total
    )

if __name__ == "__main__":
    app.run(debug=True)