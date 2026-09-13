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

# PASO 1: Carga ligera de estructura
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
                    
                    return redirect(url_for("paso2_columnas"))

                except Exception as e:
                    error = f"Error al leer la estructura del Excel: {str(e)}"
            else:
                error = "Por favor, sube un archivo con extensión .xlsx o .xls."

    return render_template("cargar_excel.html", error=error)

# PASO 2: Selección dinámica con nombres exactos requeridos
@app.route("/paso2-columnas", methods=["GET", "POST"])
def paso2_columnas():
    if not session.get("usuario"):
        return redirect(url_for("login"))

    user_key = session.get("usuario", "DICKSON")
    if user_key not in DATA_STORE:
        return redirect(url_for("cargar_excel"))

    data = DATA_STORE[user_key]

    if request.method == "POST":
        DATA_STORE[user_key]['mapeo_columnas'] = {
            'col_gln': request.form.get("col_gln") or "CANT GLN",
            'col_desc_galon': request.form.get("col_desc_galon") or "DESC X GALON",
            'col_obsv': request.form.get("col_obsv") or "OBSV ADICIONAL",
            'categoria_1': request.form.get("col_categoria_1") or "ANALISTA",
            'fecha': request.form.get("col_fecha") or "MES ENVIO"
        }
        return redirect(url_for("dashboard"))

    return render_template(
        "Lista_despegable_columna.html",
        columnas=data['columnas'],
        filename=data['filename'],
        total_columnas=len(data['columnas'])
    )

# PASO 3: Dashboard dinámico con DuckDB
@app.route("/dashboard")
def dashboard():
    if not session.get("usuario"):
        return redirect(url_for("login"))

    user_key = session.get("usuario", "DICKSON")
    if user_key not in DATA_STORE or 'mapeo_columnas' not in DATA_STORE[user_key]:
        # Mapeo por defecto con los nombres exactos si no se han enviado por POST
        DATA_STORE[user_key] = DATA_STORE.get(user_key, {})
        DATA_STORE[user_key]['mapeo_columnas'] = {
            'col_gln': "CANT GLN",
            'col_desc_galon': "DESC X GALON",
            'col_obsv': "OBSV ADICIONAL",
            'categoria_1': "ANALISTA",
            'fecha': "MES ENVIO"
        }

    mapeo = DATA_STORE[user_key]['mapeo_columnas']
    file_path = DATA_STORE[user_key]['file_path']
    filename = DATA_STORE[user_key]['filename']

    # Asignación explícita de nombres de columnas exactos entre comillas
    nombre_col_gln = "CANT GLN"
    nombre_col_desc_galon = "DESC X GALON"
    nombre_col_obsv = "OBSV ADICIONAL"
    nombre_col_cat1 = "ANALISTA"
    nombre_col_fecha = "MES ENVIO"

    cols_a_cargar = [
        nombre_col_gln,
        nombre_col_desc_galon,
        nombre_col_obsv,
        nombre_col_cat1,
        nombre_col_fecha
    ]

    try:
        engine_type = "calamine" if filename.endswith(".xlsx") else "xlrd"
        try:
            df = pd.read_excel(file_path, engine=engine_type)
        except Exception:
            df = pd.read_excel(file_path)
        
        # Limpieza ligera de espacios en las cabeceras
        df.columns = [str(col).strip() for col in df.columns]

    except Exception as e:
        return f"Error al procesar el archivo: {str(e)}"

    con = duckdb.connect()
    con.register("tabla_excel", df)

    # Definición de columnas entre comillas dobles para DuckDB
    col_gln = f'"{nombre_col_gln}"'
    col_desc_galon = f'"{nombre_col_desc_galon}"'
    col_obsv = f'"{nombre_col_obsv}"'
    col_cat1 = f'"{nombre_col_cat1}"'
    col_fecha = f'"{nombre_col_fecha}"'

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

    kpis = {
        'total_gln': f"{res_kpis[0]:,.2f}",
        'total_desc_galon': f"{res_kpis[1]:,.2f}",
        'total_obsv': f"{res_kpis[2]:,.2f}",
        'gran_total_desc': f"{res_kpis[3]:,.2f}",
        'total_registros': f"{res_kpis[4]:,}"
    }

    # 2. Resumen por "ANALISTA" y "MES ENVIO"
    resumen_tabla = []
    q_tabla = f"""
        SELECT 
            CAST({col_cat1} AS VARCHAR) as analista,
            STRFTIME(TRY_CAST({col_fecha} AS DATE), '%Y-%m') as mes_envio,
            SUM(TRY_CAST({col_gln} AS DOUBLE)) as total_gln,
            SUM(TRY_CAST({col_desc_galon} AS DOUBLE)) as sum_desc_galon,
            SUM(TRY_CAST({col_obsv} AS DOUBLE)) as sum_obsv,
            (COALESCE(SUM(TRY_CAST({col_desc_galon} AS DOUBLE)), 0) + COALESCE(SUM(TRY_CAST({col_obsv} AS DOUBLE)), 0)) as total_descuento
        FROM tabla_excel
        WHERE {col_fecha} IS NOT NULL
        GROUP BY analista, mes_envio
        ORDER BY mes_envio ASC, analista ASC
    """
    try:
        resumen_tabla = con.execute(q_tabla).fetchall()
    except Exception:
        # Reintento si la fecha viene en formato string largo
        q_tabla_fallback = f"""
            SELECT 
                CAST({col_cat1} AS VARCHAR) as analista,
                STRFTIME(TRY_CAST(STRPTIME(CAST({col_fecha} AS VARCHAR), '%Y-%m-%d %H:%M:%S') AS DATE), '%Y-%m') as mes_envio,
                SUM(TRY_CAST({col_gln} AS DOUBLE)) as total_gln,
                SUM(TRY_CAST({col_desc_galon} AS DOUBLE)) as sum_desc_galon,
                SUM(TRY_CAST({col_obsv} AS DOUBLE)) as sum_obsv,
                (COALESCE(SUM(TRY_CAST({col_desc_galon} AS DOUBLE)), 0) + COALESCE(SUM(TRY_CAST({col_obsv} AS DOUBLE)), 0)) as total_descuento
            FROM tabla_excel
            GROUP BY analista, mes_envio
            ORDER BY mes_envio ASC, analista ASC
        """
        try:
            resumen_tabla = con.execute(q_tabla_fallback).fetchall()
        except Exception:
            resumen_tabla = []

    # 3. Datos para Gráfica (Top 10 ANALISTA por Total Descuento)
    chart_cat_labels = []
    chart_cat_data = []

    q_cat = f"""
        SELECT 
            CAST({col_cat1} AS VARCHAR) as categoria, 
            (COALESCE(SUM(TRY_CAST({col_desc_galon} AS DOUBLE)), 0) + COALESCE(SUM(TRY_CAST({col_obsv} AS DOUBLE)), 0)) as total
        FROM tabla_excel
        WHERE {col_cat1} IS NOT NULL
        GROUP BY categoria
        ORDER BY total DESC
        LIMIT 10
    """
    try:
        res_cat = con.execute(q_cat).fetchall()
        chart_cat_labels = [str(r[0]) for r in res_cat]
        chart_cat_data = [float(r[1]) if r[1] else 0.0 for r in res_cat]
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
        chart_cat_labels=chart_cat_labels,
        chart_cat_data=chart_cat_data
    )

if __name__ == "__main__":
    app.run(debug=True)