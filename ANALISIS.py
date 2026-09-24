import os
import gc
import duckdb
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)
app.secret_key = "clave_secreta_fija_para_analisis_app_12345"

USUARIO_CORRECTO = "DICKSON"
PASSWORD_CORRECTO = "1234"

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

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
            session.permanent = True
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

                    ULTIMO_ARCHIVO['file_path'] = file_path
                    ULTIMO_ARCHIVO['filename'] = file.filename
                    
                    return procesar_y_renderizar_dashboard(file_path, file.filename)

                except Exception as e:
                    error = f"Error al guardar el archivo: {str(e)}"
            else:
                error = "Por favor, sube un archivo con extensión .xlsx o .xls."

    return render_template("cargar_excel.html", error=error)

@app.route("/dashboard")
def dashboard():
    if 'file_path' not in ULTIMO_ARCHIVO:
        return redirect(url_for("cargar_excel"))

    return procesar_y_renderizar_dashboard(ULTIMO_ARCHIVO['file_path'], ULTIMO_ARCHIVO['filename'])


def procesar_y_renderizar_dashboard(file_path, filename):
    try:
        try:
            df = pd.read_excel(file_path, engine="calamine")
        except Exception:
            df = pd.read_excel(file_path)
    except Exception as e:
        return f"<h3>Error al leer el archivo Excel:</h3><p>{str(e)}</p><br><a href='/cargar-excel'>Volver a intentar</a>"

    df.columns = [str(col).strip() for col in df.columns]
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
    col_cc_real = buscar_col(["CENTRO DE COSTO", "CENTRO DE COSTOS", "CENTRO COSTO", "CC", "COSTO"]) or "CENTRO DE COSTO"
    col_fact_real = buscar_col(["FACTURAS HD", "FACTURA HD", "FACTURA", "FACTURAS", "NRO FACTURA", "NUMERO FACTURA"]) or "FACTURAS HD"

    cols_a_procesar = [col_gln_real, col_desc_real, col_obsv_real]
    for col in [col_gln_real, col_desc_real, col_obsv_real, col_cat1_real, col_fecha_real, col_cc_real, col_fact_real]:
        if col not in df.columns:
            df[col] = None

    for col in cols_a_procesar:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.replace('$', '', regex=False)
            df[col] = df[col].str.replace(',', '', regex=False)
            df[col] = df[col].str.strip()

    con = duckdb.connect()
    con.register("tabla_excel", df)

    col_gln = f'"{col_gln_real}"'
    col_desc_galon = f'"{col_desc_real}"'
    col_obsv = f'"{col_obsv_real}"'
    col_cat1 = f'"{col_cat1_real}"'
    col_fecha = f'"{col_fecha_real}"'
    col_cc = f'"{col_cc_real}"'
    col_fact = f'"{col_fact_real}"'

    kpis = {
        'total_gln': "0.00",
        'total_desc_galon': "0.00",
        'total_obsv': "0.00",
        'gran_total_desc': "0.00",
        'total_registros': "0"
    }

    # 1. KPIs
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

    # 2. Tabla Resumen con CENTRO DE COSTOS y NÚMERO DE FACTURAS (Únicas por analista y mes)
    q_tabla = f"""
        SELECT 
            CAST({col_cat1} AS VARCHAR) as analista,
            SUBSTRING(CAST({col_fecha} AS VARCHAR), 1, 7) as mes_envio,
            COUNT(DISTINCT CAST({col_cc} AS VARCHAR)) as total_centros_costo,
            COUNT(DISTINCT CAST({col_fact} AS VARCHAR)) as total_facturas,
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

    del df
    gc.collect()

    return render_template(
        "dashboard.html",
        kpis=kpis,
        registros=resumen_tabla,
        filename=filename
    )

if __name__ == "__main__":
    app.run(debug=True)
