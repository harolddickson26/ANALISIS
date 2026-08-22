from flask import Flask, render_template, request, redirect, url_for, session, jsonify
import pandas as pd
import duckdb
import traceback

app = Flask(__name__)
app.secret_key = "clave_secreta_analisis_app"

USUARIO_CORRECTO = "DICKSON"
PASSWORD_CORRECTO = "1234"

DATA_STORE = {}

@app.route("/")
def inicio():
    if session.get("usuario"):
        return render_template("index.html", usuario=session["usuario"])
    return redirect(url_for("login"))

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        usuario_ingresado = request.form.get("username", "")
        password_ingresado = request.form.get("password", "")

        if usuario_ingresado == USUARIO_CORRECTO and password_ingresado == PASSWORD_CORRECTO:
            session["usuario"] = usuario_ingresado
            return redirect(url_for("inicio"))
        else:
            error = "Usuario o contraseña incorrectos."
    
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.pop("usuario", None)
    return redirect(url_for("login"))

# --- RUTA PARA SUBIR Y ANALIZAR EXCEL CON BAJO CONSUMO DE MEMORIA ---
@app.route("/cargar-excel", methods=["GET", "POST"])
def cargar_excel():
    if not session.get("usuario"): 
        return redirect(url_for("login"))
    
    error = None
    kpis = {}
    estaciones = []
    anios = []

    if request.method == "POST":
        if "archivo_excel" not in request.files:
            error = "No se seleccionó ningún archivo."
        else:
            file = request.files["archivo_excel"]
            if file.filename == "":
                error = "Nombre de archivo no válido."
            elif file and (file.filename.endswith(".xlsx") or file.filename.endswith(".xls")):
                try:
                    # Carga optimizada sin formateo visual para ahorrar memoria RAM
                    if file.filename.endswith(".xlsx"):
                        df = pd.read_excel(file, engine="openpyxl")
                    else:
                        df = pd.read_excel(file, engine="xlrd")
                    
                    # Limpiar nombres de columnas
                    df.columns = [str(col).strip().lower() for col in df.columns]
                    
                    user_key = session.get("usuario", "DICKSON")
                    DATA_STORE[user_key] = df
                    
                    cols = df.columns.tolist()

                    col_estacion = next((c for c in cols if any(k in c for k in ['estacion', 'sede', 'zona', 'centro'])), None)
                    if col_estacion:
                        estaciones = sorted([str(x) for x in df[col_estacion].dropna().unique().tolist()])

                    col_fecha = next((c for c in cols if any(k in c for k in ['fecha', 'date', 'dia'])), None)
                    if col_fecha:
                        df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce')
                        anios = sorted([str(int(x)) for x in df[col_fecha].dt.year.dropna().unique().tolist()], reverse=True)

                    num_cols = df.select_dtypes(include=['number']).columns.tolist()
                    col_monto = next((c for c in num_cols if any(k in c for k in ['monto', 'total', 'venta', 'valor', 'precio', 'importe'])), num_cols[-1] if num_cols else None)

                    if not col_monto:
                        for col in cols:
                            try:
                                df[col] = pd.to_numeric(df[col], errors='coerce')
                                if df[col].notna().sum() > 0:
                                    col_monto = col
                                    break
                            except:
                                pass

                    if not col_monto:
                        raise Exception("No se encontró ninguna columna numérica para realizar los cálculos.")

                    # Consultas ultra rápidas usando memoria de DuckDB
                    res = duckdb.query(f"SELECT SUM({col_monto}) as total_v, COUNT(*) as total_r FROM df").fetchone()
                    total_v = float(res[0]) if res[0] is not None else 0.0
                    total_r = int(res[1]) if res[1] is not None else 0
                    prom = total_v / total_r if total_r > 0 else 0
                    
                    kpis = {
                        'total_ventas': f"${total_v:,.2f}",
                        'total_remisiones': f"{total_r:,}",
                        'promedio_venta': f"${prom:,.2f}"
                    }

                except Exception as e:
                    error = f"Error al procesar el archivo Excel: {str(e)}"
            else:
                error = "Por favor, sube un archivo con extensión .xlsx o .xls."

    return render_template(
        "cargar_excel.html", 
        error=error, 
        kpis=kpis, 
        estaciones=estaciones, 
        anios=anios
    )

# --- API DE FILTRADO PARA DESPLEGABLES ---
@app.route("/api/filtrar-analisis", methods=["POST"])
def filtrar_analisis():
    user = session.get("usuario", "DICKSON")
    if user not in DATA_STORE:
        return jsonify({'error': 'No hay datos cargados'}), 400

    df = DATA_STORE[user]
    data = request.get_json() or {}
    estacion_sel = data.get('estacion')
    anio_sel = data.get('anio')

    cols = df.columns.tolist()
    col_estacion = next((c for c in cols if any(k in c for k in ['estacion', 'sede', 'zona', 'centro'])), None)
    col_fecha = next((c for c in cols if any(k in c for k in ['fecha', 'date', 'dia'])), None)

    # Filtrado mediante DuckDB para evitar duplicados en memoria RAM
    query = "SELECT * FROM df WHERE 1=1"
    if col_estacion and estacion_sel and estacion_sel != 'Todas':
        query += f" AND CAST({col_estacion} AS VARCHAR) = '{estacion_sel}'"
        
    if col_fecha and anio_sel and anio_sel != 'Todos':
        query += f" AND STRFTIME({col_fecha}, '%Y') = '{anio_sel}'"

    df_filtered = duckdb.query(query).df()

    num_cols = df_filtered.select_dtypes(include=['number']).columns.tolist()
    col_monto = next((c for c in num_cols if any(k in c for k in ['monto', 'total', 'venta', 'valor', 'precio', 'importe'])), num_cols[-1] if num_cols else cols[0])
    
    text_cols = df_filtered.select_dtypes(include=['object']).columns.tolist()
    col_prod = next((c for c in text_cols if any(k in c for k in ['producto', 'categoria', 'descripcion', 'item'])), text_cols[0] if text_cols else cols[0])

    total_v = float(df_filtered[col_monto].sum()) if col_monto in df_filtered else 0.0
    total_r = len(df_filtered)
    prom = total_v / total_r if total_r > 0 else 0.0

    labels_prod, valores_prod = [], []
    if col_prod in df_filtered and col_monto in df_filtered:
        grp_prod = df_filtered.groupby(col_prod)[col_monto].sum().head(5)
        labels_prod = [str(x) for x in grp_prod.index.tolist()]
        valores_prod = [float(x) for x in grp_prod.values.tolist()]

    labels_mes, valores_mes = [], []
    if col_fecha in df_filtered and col_monto in df_filtered:
        df_filtered['mes_str'] = pd.to_datetime(df_filtered[col_fecha]).dt.strftime('%m-%b')
        grp_mes = df_filtered.groupby('mes_str')[col_monto].sum()
        labels_mes = [str(x) for x in grp_mes.index.tolist()]
        valores_mes = [float(x) for x in grp_mes.values.tolist()]

    return jsonify({
        'kpi_total': f"${total_v:,.2f}",
        'kpi_remisiones': f"{total_r:,}",
        'kpi_promedio': f"${prom:,.2f}",
        'labels_prod': labels_prod,
        'valores_prod': valores_prod,
        'labels_mes': labels_mes,
        'valores_mes': valores_mes
    })

if __name__ == "__main__":
    app.run(debug=True)