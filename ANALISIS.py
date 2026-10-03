import os
import gc
import json
import duckdb
import pandas as pd
import numpy as np
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


# --- NUEVA FUNCIÓN PARA LA TABLA DE CENTROS DE COSTOS ---
def generar_tabla_centros_costos(df_raw, col_cc_real, col_cat1_real, col_fecha_real, col_gln_real):
    df = df_raw.copy()
    df['CENTRO DE COSTOS NUM'] = pd.to_numeric(df[col_cc_real], errors='coerce')
    df['MES_STR'] = pd.to_datetime(df[col_fecha_real], errors='coerce').dt.strftime('%Y-%m')
    
    # Obtener el analista predominante por centro de costos
    analyst_mapping = df.groupby('CENTRO DE COSTOS NUM')[col_cat1_real].agg(
        lambda x: x.mode()[0] if not x.mode().empty else x.iloc[0]
    ).reset_index()
    analyst_mapping.columns = ['CENTRO DE COSTOS NUM', 'ANALISTA']
    
    pivot_df = pd.pivot_table(
        df,
        index='CENTRO DE COSTOS NUM',
        columns='MES_STR',
        values=col_gln_real,
        aggfunc='sum',
        fill_value=0
    ).sort_index()
    
    months = sorted([str(m) for m in pivot_df.columns.tolist() if pd.notna(m)])
    combined_df = pd.DataFrame(index=pivot_df.index)
    
    prev_col = None
    for i, col in enumerate(months):
        if i == 0:
            combined_df[col] = pivot_df[col].apply(lambda x: f"{x:,.2f}")
        else:
            old_val = pivot_df[prev_col]
            new_val = pivot_df[col]
            pct_change = np.where(
                old_val == 0,
                np.where(new_val == 0, 0.0, 100.0),
                ((new_val - old_val) / old_val) * 100
            )
            
            cell_contents = []
            for val, p, o, n in zip(new_val, pct_change, old_val, new_val):
                val_str = f"{val:,.2f}"
                if o == 0 and n == 0:
                    cell_contents.append(f"{val_str} (0.0% ➖)")
                elif p > 0:
                    cell_contents.append(f"{val_str} (+{p:.1f}% 🟢)")
                elif p < 0:
                    cell_contents.append(f"{val_str} ({p:.1f}% 🔴)")
                else:
                    cell_contents.append(f"{val_str} (0.0% ➖)")
            combined_df[col] = cell_contents
        prev_col = col

    combined_df['TOTAL GENERAL'] = pivot_df.sum(axis=1).apply(lambda x: f"{x:,.2f}")
    combined_df = combined_df.reset_index()
    combined_df = pd.merge(combined_df, analyst_mapping, on='CENTRO DE COSTOS NUM', how='left')
    
    cols = ['CENTRO DE COSTOS NUM', 'ANALISTA'] + [c for c in combined_df.columns if c not in ['CENTRO DE COSTOS NUM', 'ANALISTA']]
    final_df = combined_df[cols]
    
    return final_df.to_dict(orient='records'), final_df.columns.tolist()


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
    col_cc_real = buscar_col(["CENTRO DE COSTO", "CENTRO DE COSTOS", "CENTRO COSTO", "EDS", "CEN", "CC", "COSTO"]) or "CENTRO DE COSTO"
    col_fact_real = buscar_col(["FACTURAS HD", "FACTURA HD", "FACTURA", "FACTURAS", "NRO FACTURA", "NUMERO FACTURA"]) or "FACTURAS HD"
    col_prod_real = buscar_col(["PRODUCTO", "COMBUSTIBLE", "DESCRIPCION PRODUCTO", "PROD"]) or "PRODUCTO"

    cols_a_procesar = [col_gln_real, col_desc_real, col_obsv_real]
    for col in [col_gln_real, col_desc_real, col_obsv_real, col_cat1_real, col_fecha_real, col_cc_real, col_fact_real, col_prod_real]:
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
    col_prod = f'"{col_prod_real}"'

    kpis = {
        'total_gln': "0.00",
        'total_desc_galon': "0.00",
        'total_obsv': "0.00",
        'gran_total_desc': "0.00",
        'total_registros': "0"
    }

    # 1. KPIs Generales Totales
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

    # 2. Tabla Resumen Principal Estática
    q_tabla = f"""
        SELECT 
            SUBSTRING(CAST({col_fecha} AS VARCHAR), 1, 7) as mes_envio,
            CAST({col_cat1} AS VARCHAR) as analista,
            COUNT(DISTINCT CAST({col_cc} AS VARCHAR)) as total_centros_costo,
            COUNT(DISTINCT CAST({col_fact} AS VARCHAR)) as total_facturas,
            COALESCE(SUM(TRY_CAST({col_gln} AS DOUBLE)), 0) as total_gln,
            (COALESCE(SUM(TRY_CAST({col_desc_galon} AS DOUBLE)), 0) + COALESCE(SUM(TRY_CAST({col_obsv} AS DOUBLE)), 0)) as total_descuento,
            COALESCE(SUM(TRY_CAST({col_desc_galon} AS DOUBLE)), 0) as sum_desc_galon,
            COALESCE(SUM(TRY_CAST({col_obsv} AS DOUBLE)), 0) as sum_obsv
        FROM tabla_excel
        GROUP BY mes_envio, analista
        ORDER BY mes_envio ASC, analista ASC
    """
    try:
        resumen_tabla = con.execute(q_tabla).fetchall()
    except Exception:
        resumen_tabla = []

    # Generar los datos para la nueva tabla de Centros de Costos
    try:
        cc_tabla_registros, cc_columnas = generar_tabla_centros_costos(df, col_cc_real, col_cat1_real, col_fecha_real, col_gln_real)
    except Exception:
        cc_tabla_registros, cc_columnas = [], []

    # 3. Registros para Filtros Dinámicos
    q_raw_data = f"""
        SELECT 
            UPPER(TRIM(COALESCE(CAST({col_cat1} AS VARCHAR), 'DESCONOCIDO'))) as analista,
            SUBSTRING(CAST({col_fecha} AS VARCHAR), 1, 7) as mes_envio,
            UPPER(TRIM(COALESCE(CAST({col_cc} AS VARCHAR), 'SIN CC/EDS'))) as centro_costos,
            UPPER(TRIM(COALESCE(CAST({col_prod} AS VARCHAR), 'GENERAL'))) as producto,
            CAST({col_fact} AS VARCHAR) as factura_hd,
            COALESCE(TRY_CAST({col_gln} AS DOUBLE), 0) as gln,
            COALESCE(TRY_CAST({col_desc_galon} AS DOUBLE), 0) as desc_galon,
            COALESCE(TRY_CAST({col_obsv} AS DOUBLE), 0) as obsv
        FROM tabla_excel
    """
    try:
        raw_rows = con.execute(q_raw_data).fetchall()
        raw_data = [
            {
                "analista": r[0],
                "mes": r[1] if r[1] else "N/A",
                "cc": r[2],
                "producto": r[3],
                "factura": r[4],
                "gln": r[5],
                "desc": r[6],
                "obsv": r[7],
                "total_desc": r[6] + r[7]
            }
            for r in raw_rows
        ]
    except Exception:
        raw_data = []

    del df
    gc.collect()

    return render_template(
        "dashboard.html",
        kpis=kpis,
        registros=resumen_tabla,
        filename=filename,
        raw_data_json=json.dumps(raw_data),
        cc_tabla_registros=cc_tabla_registros,
        cc_columnas=cc_columnas
    )

if __name__ == "__main__":
    app.run(debug=True)
