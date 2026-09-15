import os
import re
import duckdb
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)
app.secret_key = "super_secret_key_analisis_web"

# Directorio de subidas
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Diccionario global de sesión temporal
DATA_STORE = {}


def buscar_col(df_columns, lista_opciones, fallback_idx):
    """Busca dinámicamente el nombre de columna en el DataFrame basándose en coincidencias."""
    cols_existentes = {str(col).upper().strip(): col for col in df_columns}
    for opcion in lista_opciones:
        for k, v in cols_existentes.items():
            if opcion.upper() in k:
                return v
    if len(df_columns) > fallback_idx:
        return df_columns[fallback_idx]
    return df_columns[0] if len(df_columns) > 0 else None


@app.route("/", methods=["GET", "POST"])
def cargar_excel():
    if request.method == "POST":
        if 'file' not in request.files:
            return render_template("index.html", error="No se seleccionó ningún archivo.")
        
        file = request.files['file']
        if file.filename == '':
            return render_template("index.html", error="El archivo no tiene nombre.")
        
        if file:
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            
            # Asignar usuario de sesión por defecto si no existe
            user_key = session.get("usuario", "DICKSON")
            session["usuario"] = user_key
            
            DATA_STORE[user_key] = {
                'file_path': filepath,
                'filename': file.filename
            }
            return redirect(url_for("dashboard"))
            
    return render_template("index.html")


@app.route("/dashboard")
def dashboard():
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
        
        # Normalizar nombres de columnas eliminando espacios extremos
        df.columns = [str(col).strip() for col in df.columns]
    except Exception as e:
        return f"Error al procesar el archivo Excel: {str(e)}"

    if df.empty:
        return "El archivo cargado está vacío."

    # Identificación segura de columnas
    col_gln_real = buscar_col(df.columns, ["CANT GLN", "CANTIDAD", "GALON", "GLN"], 0)
    col_desc_real = buscar_col(df.columns, ["DESC X GALON", "DESCUENTO", "DESC"], 1)
    col_obsv_real = buscar_col(df.columns, ["OBSERVACION", "OBSERVACIONES", "OBSV", "ADICIONAL"], 2)
    col_cat1_real = buscar_col(df.columns, ["ANALISTA", "USUARIO", "RESPONSABLE"], 3)
    col_fecha_real = buscar_col(df.columns, ["MES ENVIO", "FECHA", "MES"], 4)
    col_cc_real = buscar_col(df.columns, ["CENTRO DE COSTO", "CENTRO COSTO", "CENTRO_COSTO", "COSTO", "CC", "CENTRO"], 5)

    # Crear columnas internas estandarizadas e inmunes a errores SQL/Formatos
    df['c_gln'] = pd.to_numeric(df[col_gln_real].astype(str).str.replace(r'[\$,]', '', regex=True), errors='coerce').fillna(0)
    df['c_desc'] = pd.to_numeric(df[col_desc_real].astype(str).str.replace(r'[\$,]', '', regex=True), errors='coerce').fillna(0)
    df['c_obsv'] = pd.to_numeric(df[col_obsv_real].astype(str).str.replace(r'[\$,]', '', regex=True), errors='coerce').fillna(0)
    df['c_analista'] = df[col_cat1_real].astype(str).str.strip().str.upper()
    df['c_fecha'] = df[col_fecha_real].astype(str).str.strip()
    df['c_cc'] = df[col_cc_real].astype(str).str.strip()

    # Registro en DuckDB
    con = duckdb.connect()
    con.register("tabla_excel", df)

    # 1. KPIs Generales
    query_kpis = """
        SELECT 
            COALESCE(SUM(c_gln), 0) as total_gln,
            COALESCE(SUM(c_desc), 0) as total_desc_galon,
            COALESCE(SUM(c_obsv), 0) as total_obsv,
            (COALESCE(SUM(c_desc), 0) + COALESCE(SUM(c_obsv), 0)) as gran_total_desc,
            COUNT(*) as total_registros
        FROM tabla_excel
    """
    res_kpis = con.execute(query_kpis).fetchone()
    kpis = {
        'total_gln': f"{res_kpis[0]:,.2f}",
        'total_desc_galon': f"${res_kpis[1]:,.2f}",
        'total_obsv': f"${res_kpis[2]:,.2f}",
        'gran_total_desc': f"${res_kpis[3]:,.2f}",
        'total_registros': f"{res_kpis[4]:,}"
    }

    # 2. Resumen por Analista / Responsable
    query_analistas = """
        SELECT 
            c_analista,
            SUM(c_gln) as galones,
            SUM(c_desc + c_obsv) as descuento_total
        FROM tabla_excel
        GROUP BY c_analista
        ORDER BY descuento_total DESC
    """
    df_analistas = con.execute(query_analistas).df()

    # 3. Resumen por Centro de Costo
    query_cc = """
        SELECT 
            c_cc,
            SUM(c_gln) as galones,
            SUM(c_desc + c_obsv) as descuento_total
        FROM tabla_excel
        GROUP BY c_cc
        ORDER BY galones DESC
        LIMIT 10
    """
    df_cc = con.execute(query_cc).df()

    con.close()

    return render_template(
        "dashboard.html",
        filename=filename,
        kpis=kpis,
        labels_analista=df_analistas['c_analista'].tolist(),
        data_analista=df_analistas['descuento_total'].tolist(),
        labels_cc=df_cc['c_cc'].tolist(),
        data_cc=df_cc['galones'].tolist()
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)