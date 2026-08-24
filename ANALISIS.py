import os
import io
import gc
import duckdb
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)
app.secret_key = "clave_secreta_analisis_app"

USUARIO_CORRECTO = "DICKSON"
PASSWORD_CORRECTO = "1234"

# Carpeta para archivos temporales
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
                    # Guardar archivo en disco local para no agotar la memoria RAM
                    file_path = os.path.join(UPLOAD_FOLDER, f"{user_key}_{file.filename}")
                    file.save(file_path)

                    # Leer solo las primeras 5 filas para extraer encabezados super rápido
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

# PASO 2: Selección de columnas
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
            'monto': request.form.get("col_monto"),
            'cantidad': request.form.get("col_cantidad"),
            'incremento': request.form.get("col_incremento"),
            'fecha': request.form.get("col_fecha"),
            'estacion': request.form.get("col_estacion"),
            'producto': request.form.get("col_producto")
        }
        return redirect(url_for("dashboard"))

    return render_template(
        "Lista_despegable_columna.html",
        columnas=data['columnas'],
        filename=data['filename'],
        total_columnas=len(data['columnas'])
    )

# PASO 3: Dashboard con DuckDB leyendo directo del disco
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

    # Cargar solo columnas necesarias en Pandas para la consulta
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

    col_monto = f'"{mapeo["monto"]}"' if mapeo.get("monto") else "0"
    col_cantidad = f'"{mapeo["cantidad"]}"' if mapeo.get("cantidad") else "0"
    col_estacion = f'"{mapeo["estacion"]}"' if mapeo.get("estacion") else "NULL"
    col_producto = f'"{mapeo["producto"]}"' if mapeo.get("producto") else "NULL"
    col_fecha = f'"{mapeo["fecha"]}"' if mapeo.get("fecha") else "NULL"

    # KPIs
    query_kpis = f"""
        SELECT 
            COALESCE(SUM(TRY_CAST({col_monto} AS DOUBLE)), 0) as total_monto,
            COALESCE(SUM(TRY_CAST({col_cantidad} AS DOUBLE)), 0) as total_cantidad,
            COUNT(*) as total_registros
        FROM tabla_excel
    """
    res_kpis = con.execute(query_kpis).fetchone()

    # Resumen por Mes
    resumen_meses = []
    if mapeo.get("fecha"):
        q_mes = f"""
            SELECT 
                STRFTIME(TRY_CAST({col_fecha} AS DATE), '%Y-%m') as mes,
                SUM(TRY_CAST({col_monto} AS DOUBLE)) as monto,
                SUM(TRY_CAST({col_cantidad} AS DOUBLE)) as cantidad,
                COUNT(*) as registros
            FROM tabla_excel
            WHERE TRY_CAST({col_fecha} AS DATE) IS NOT NULL
            GROUP BY mes
            ORDER BY mes ASC
        """
        try:
            resumen_meses = con.execute(q_mes).fetchall()
        except Exception:
            q_mes_fallback = f"""
                SELECT 
                    STRFTIME(TRY_CAST(STRPTIME(CAST({col_fecha} AS VARCHAR), '%Y-%m-%d %H:%M:%S') AS DATE), '%Y-%m') as mes,
                    SUM(TRY_CAST({col_monto} AS DOUBLE)) as monto,
                    SUM(TRY_CAST({col_cantidad} AS DOUBLE)) as cantidad,
                    COUNT(*) as registros
                FROM tabla_excel
                GROUP BY mes
                ORDER BY mes ASC
            """
            resumen_meses = con.execute(q_mes_fallback).fetchall()

    # Top 5 Estaciones
    resumen_estaciones = []
    if mapeo.get("estacion"):
        q_est = f"""
            SELECT {col_estacion} as estacion, 
                   SUM(TRY_CAST({col_monto} AS DOUBLE)) as monto, 
                   SUM(TRY_CAST({col_cantidad} AS DOUBLE)) as cantidad
            FROM tabla_excel
            GROUP BY {col_estacion}
            ORDER BY monto DESC
            LIMIT 5
        """
        resumen_estaciones = con.execute(q_est).fetchall()

    # Top 5 Productos
    resumen_productos = []
    if mapeo.get("producto"):
        q_prod = f"""
            SELECT {col_producto} as producto, 
                   SUM(TRY_CAST({col_monto} AS DOUBLE)) as monto, 
                   SUM(TRY_CAST({col_cantidad} AS DOUBLE)) as cantidad
            FROM tabla_excel
            GROUP BY {col_producto}
            ORDER BY monto DESC
            LIMIT 5
        """
        resumen_productos = con.execute(q_prod).fetchall()

    kpis = {
        'total_monto': f"${res_kpis[0]:,.2f}",
        'total_cantidad': f"{res_kpis[1]:,.2f}",
        'total_registros': f"{res_kpis[2]:,}"
    }

    # Limpieza explicita de RAM
    del df
    gc.collect()

    return render_template(
        "dashboard.html",
        kpis=kpis,
        meses=resumen_meses,
        estaciones=resumen_estaciones,
        productos=resumen_productos,
        filename=filename,
        mapeo=mapeo
    )

if __name__ == "__main__":
    app.run(debug=True)