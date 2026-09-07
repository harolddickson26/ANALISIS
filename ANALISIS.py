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
                    df_preview = pd.read_excel(file_path, engine=engine_type, nrows=5)
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

# PASO 2: Selección dinámicas de métricas y categorías
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
            'kpi_1': request.form.get("col_kpi_1"),
            'kpi_2': request.form.get("col_kpi_2"),
            'categoria_1': request.form.get("col_categoria_1"),
            'fecha': request.form.get("col_fecha")
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
        return redirect(url_for("cargar_excel"))

    mapeo = DATA_STORE[user_key]['mapeo_columnas']
    file_path = DATA_STORE[user_key]['file_path']
    filename = DATA_STORE[user_key]['filename']

    cols_a_cargar = list({v for v in mapeo.values() if v and v != "None"})

    try:
        engine_type = "calamine" if filename.endswith(".xlsx") else "xlrd"
        if cols_a_cargar:
            df = pd.read_excel(file_path, engine=engine_type, usecols=cols_a_cargar)
        else:
            df = pd.read_excel(file_path, engine=engine_type)
        
        df.columns = [str(col).strip() for col in df.columns]
    except Exception as e:
        return f"Error al procesar el archivo: {str(e)}"

    con = duckdb.connect()
    con.register("tabla_excel", df)

    col_kpi1 = f'"{mapeo["kpi_1"]}"' if mapeo.get("kpi_1") else "0"
    col_kpi2 = f'"{mapeo["kpi_2"]}"' if mapeo.get("kpi_2") else "0"
    col_cat1 = f'"{mapeo["categoria_1"]}"' if mapeo.get("categoria_1") else "NULL"
    col_fecha = f'"{mapeo["fecha"]}"' if mapeo.get("fecha") else "NULL"

    # 1. KPIs Generales
    query_kpis = f"""
        SELECT 
            COALESCE(SUM(TRY_CAST({col_kpi1} AS DOUBLE)), 0) as total_kpi1,
            COALESCE(SUM(TRY_CAST({col_kpi2} AS DOUBLE)), 0) as total_kpi2,
            COUNT(*) as total_registros
        FROM tabla_excel
    """
    res_kpis = con.execute(query_kpis).fetchone()

    kpis = {
        'nombre_kpi1': mapeo.get("kpi_1") or "Métrica 1",
        'val_kpi1': f"{res_kpis[0]:,.2f}",
        'nombre_kpi2': mapeo.get("kpi_2") or "Métrica 2",
        'val_kpi2': f"{res_kpis[1]:,.2f}",
        'total_registros': f"{res_kpis[2]:,}"
    }

    # 2. Agrupación por Categoría / Dimensión seleccionada (Top 10)
    chart_cat_labels = []
    chart_cat_data = []

    if mapeo.get("categoria_1") and mapeo.get("kpi_1"):
        q_cat = f"""
            SELECT 
                CAST({col_cat1} AS VARCHAR) as categoria, 
                SUM(TRY_CAST({col_kpi1} AS DOUBLE)) as total
            FROM tabla_excel
            WHERE {col_cat1} IS NOT NULL
            GROUP BY categoria
            ORDER BY total DESC
            LIMIT 10
        """
        res_cat = con.execute(q_cat).fetchall()
        chart_cat_labels = [str(r[0]) for r in res_cat]
        chart_cat_data = [float(r[1]) if r[1] else 0.0 for r in res_cat]

    # 3. Resumen por Mes
    resumen_meses = []
    chart_meses_labels = []
    chart_meses_data = []

    if mapeo.get("fecha") and mapeo.get("kpi_1"):
        q_mes = f"""
            SELECT 
                STRFTIME(TRY_CAST({col_fecha} AS DATE), '%Y-%m') as mes,
                SUM(TRY_CAST({col_kpi1} AS DOUBLE)) as total_kpi1,
                SUM(TRY_CAST({col_kpi2} AS DOUBLE)) as total_kpi2,
                COUNT(*) as registros
            FROM tabla_excel
            WHERE TRY_CAST({col_fecha} AS DATE) IS NOT NULL
            GROUP BY mes
            ORDER BY mes ASC
        """
        try:
            resumen_meses = con.execute(q_mes).fetchall()
            chart_meses_labels = [str(r[0]) if r[0] else 'Sin Fecha' for r in resumen_meses]
            chart_meses_data = [float(r[1]) if r[1] else 0.0 for r in resumen_meses]
        except Exception:
            q_mes_fallback = f"""
                SELECT 
                    STRFTIME(TRY_CAST(STRPTIME(CAST({col_fecha} AS VARCHAR), '%Y-%m-%d %H:%M:%S') AS DATE), '%Y-%m') as mes,
                    SUM(TRY_CAST({col_kpi1} AS DOUBLE)) as total_kpi1,
                    SUM(TRY_CAST({col_kpi2} AS DOUBLE)) as total_kpi2,
                    COUNT(*) as registros
                FROM tabla_excel
                GROUP BY mes
                ORDER BY mes ASC
            """
            try:
                resumen_meses = con.execute(q_mes_fallback).fetchall()
                chart_meses_labels = [str(r[0]) if r[0] else 'Sin Fecha' for r in resumen_meses]
                chart_meses_data = [float(r[1]) if r[1] else 0.0 for r in resumen_meses]
            except Exception:
                resumen_meses = []

    del df
    gc.collect()

    return render_template(
        "dashboard.html",
        kpis=kpis,
        meses=resumen_meses,
        filename=filename,
        mapeo=mapeo,
        chart_cat_labels=chart_cat_labels,
        chart_cat_data=chart_cat_data,
        chart_meses_labels=chart_meses_labels,
        chart_meses_data=chart_meses_data
    )

if __name__ == "__main__":
    app.run(debug=True)