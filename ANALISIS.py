import os
import duckdb
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)
app.secret_key = "clave_secreta_analisis"

UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Credenciales de usuario fijas
USUARIOS_VALIDOS = {
    "DICKSON": "12345",
    "FABIAN": "12345",
    "admin": "admin123"
}

DATA_STORE = {}

def buscar_col(df_columns, lista_opciones, fallback_idx):
    cols_existentes = {str(col).upper().strip(): col for col in df_columns}
    for opcion in lista_opciones:
        for k, v in cols_existentes.items():
            if opcion.upper() in k:
                return v
    return df_columns[fallback_idx] if len(df_columns) > fallback_idx else df_columns[0]

@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip().upper()
        password = request.form.get("password", "").strip()
        
        if usuario in USUARIOS_VALIDOS and USUARIOS_VALIDOS[usuario] == password:
            session["usuario"] = usuario
            return redirect(url_for("cargar_excel"))
        else:
            return render_template("login.html", error="Usuario o contraseña incorrectos")
            
    return render_template("login.html")

@app.route("/cargar", methods=["GET", "POST"])
def cargar_excel():
    if not session.get("usuario"):
        return redirect(url_for("login"))
        
    if request.method == "POST":
        if 'file' not in request.files:
            return render_template("cargar.html", error="No se seleccionó ningún archivo.")
        
        file = request.files['file']
        if file.filename != '':
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            
            user_key = session["usuario"]
            DATA_STORE[user_key] = {
                'file_path': filepath,
                'filename': file.filename
            }
            return redirect(url_for("dashboard"))
            
    return render_template("cargar.html", usuario=session.get("usuario"))

@app.route("/dashboard")
def dashboard():
    usuario = session.get("usuario")
    if not usuario or usuario not in DATA_STORE:
        return redirect(url_for("login"))

    file_path = DATA_STORE[usuario]['file_path']
    filename = DATA_STORE[usuario]['filename']

    try:
        engine_type = "calamine" if filename.endswith(".xlsx") else "xlrd"
        try:
            df = pd.read_excel(file_path, engine=engine_type)
        except Exception:
            df = pd.read_excel(file_path)
        df.columns = [str(col).strip() for col in df.columns]
    except Exception as e:
        return f"Error al procesar el archivo Excel: {str(e)}"

    # Identificación dinámica de columnas
    col_gln_real = buscar_col(df.columns, ["CANT GLN", "CANTIDAD", "GALON", "GLN"], 0)
    col_desc_real = buscar_col(df.columns, ["DESC X GALON", "DESCUENTO", "DESC"], 1)
    col_obsv_real = buscar_col(df.columns, ["OBSERVACION", "OBSERVACIONES", "OBSV", "ADICIONAL"], 2)
    col_cat1_real = buscar_col(df.columns, ["ANALISTA", "USUARIO", "RESPONSABLE"], 3)
    col_cc_real = buscar_col(df.columns, ["CENTRO DE COSTO", "CENTRO COSTO", "CC"], 5)

    # Limpieza segura para DuckDB
    df['c_gln'] = pd.to_numeric(df[col_gln_real].astype(str).str.replace(r'[\$,]', '', regex=True), errors='coerce').fillna(0)
    df['c_desc'] = pd.to_numeric(df[col_desc_real].astype(str).str.replace(r'[\$,]', '', regex=True), errors='coerce').fillna(0)
    df['c_obsv'] = pd.to_numeric(df[col_obsv_real].astype(str).str.replace(r'[\$,]', '', regex=True), errors='coerce').fillna(0)
    df['c_analista'] = df[col_cat1_real].astype(str).str.strip().str.upper()
    df['c_cc'] = df[col_cc_real].astype(str).str.strip()

    con = duckdb.connect()
    con.register("tabla_excel", df)

    # KPIs
    res_kpis = con.execute("""
        SELECT 
            COALESCE(SUM(c_gln), 0),
            COALESCE(SUM(c_desc), 0),
            COALESCE(SUM(c_obsv), 0),
            (COALESCE(SUM(c_desc), 0) + COALESCE(SUM(c_obsv), 0)),
            COUNT(*)
        FROM tabla_excel
    """).fetchone()

    kpis = {
        'total_gln': f"{res_kpis[0]:,.2f}",
        'total_desc_galon': f"${res_kpis[1]:,.2f}",
        'total_obsv': f"${res_kpis[2]:,.2f}",
        'gran_total_desc': f"${res_kpis[3]:,.2f}",
        'total_registros': f"{res_kpis[4]:,}"
    }

    # Gráfico 1: Total Descuento por Analista
    df_analistas = con.execute("""
        SELECT c_analista, SUM(c_desc + c_obsv) as descuento_total
        FROM tabla_excel
        GROUP BY c_analista
        ORDER BY descuento_total DESC
    """).df()

    # Tabla Resumen detallada
    df_tabla = con.execute("""
        SELECT 
            c_analista as Analista,
            COUNT(*) as Registros,
            SUM(c_gln) as Galones,
            SUM(c_desc) as Desc_Galon,
            SUM(c_obsv) as Observaciones,
            SUM(c_desc + c_obsv) as Total_Descuento
        FROM tabla_excel
        GROUP BY c_analista
        ORDER BY Total_Descuento DESC
    """).df()

    # Formateo de la tabla para HTML
    tabla_html = []
    for _, r in df_tabla.iterrows():
        tabla_html.append({
            'analista': r['Analista'],
            'registros': f"{int(r['Registros']):,}",
            'galones': f"{r['Galones']:,.2f}",
            'desc_galon': f"${r['Desc_Galon']:,.2f}",
            'observaciones': f"${r['Observaciones']:,.2f}",
            'total_descuento': f"${r['Total_Descuento']:,.2f}"
        })

    # Conclusiones automatizadas
    top_analista = df_analistas.iloc[0]['c_analista'] if not df_analistas.empty else "N/A"
    top_monto = f"${df_analistas.iloc[0]['descuento_total']:,.2f}" if not df_analistas.empty else "$0"
    
    conclusiones = [
        f"El volumen total procesado alcanza los {kpis['total_gln']} galones en un total de {kpis['total_registros']} registros.",
        f"El analista con mayor participación en descuentos es **{top_analista}** con un consolidado de **{top_monto}**.",
        f"El Gran Total de Descuentos otorgados suma **{kpis['gran_total_desc']}**, compuesto por {kpis['total_desc_galon']} de descuento directo y {kpis['total_obsv']} de adicionales."
    ]

    con.close()

    return render_template(
        "dashboard.html",
        filename=filename,
        usuario=usuario,
        kpis=kpis,
        labels_analista=df_analistas['c_analista'].tolist(),
        data_analista=df_analistas['descuento_total'].tolist(),
        tabla=tabla_html,
        conclusiones=conclusiones
    )

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

if __name__ == "__main__":
    app.run(debug=True, port=5000)