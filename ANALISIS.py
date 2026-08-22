from flask import Flask, render_template, request, redirect, url_for, session, jsonify
import pandas as pd
import duckdb
import json

app = Flask(__name__)
app.secret_key = "clave_secreta_analisis_app"

USUARIO_CORRECTO = "DICKSON"
PASSWORD_CORRECTO = "1234"

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
    kpis = {'total_ventas': '$0.00', 'total_remisiones': '0', 'promedio_venta': '$0.00'}
    estaciones = []
    anios = []
    productos = []
    labels_mes, valores_mes = [], []
    labels_prod, valores_prod = [], []

    if request.method == "POST":
        if "archivo_excel" not in request.files:
            error = "No se seleccionó ningún archivo."
        else:
            file = request.files["archivo_excel"]
            if file.filename == "":
                error = "Nombre de archivo no válido."
            elif file and (file.filename.endswith(".xlsx") or file.filename.endswith(".xls")):
                try:
                    if file.filename.endswith(".xlsx"):
                        df = pd.read_excel(file, engine="calamine")
                    else:
                        df = pd.read_excel(file, engine="xlrd")
                    
                    df.columns = [str(col).strip().lower() for col in df.columns]
                    
                    user_key = session.get("usuario", "DICKSON")
                    DATA_STORE[user_key] = df
                    
                    cols = df.columns.tolist()

                    col_estacion = next((c for c in cols if any(k in c for k in ['estacion', 'estación', 'eds', 'sede', 'zona', 'centro', 'local', 'punto', 'cliente'])), None)
                    if col_estacion:
                        estaciones = sorted([str(x) for x in df[col_estacion].dropna().unique().tolist() if str(x).strip() != ''])

                    col_fecha = next((c for c in cols if any(k in c for k in ['fecha', 'date', 'dia', 'día', 'fec'])), None)
                    if col_fecha:
                        df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce')
                        anios = sorted([str(int(x)) for x in df[col_fecha].dt.year.dropna().unique().tolist()], reverse=True)

                    text_cols = df.select_dtypes(include=['object']).columns.tolist()
                    col_prod = next((c for c in text_cols if any(k in c for k in ['producto', 'categoria', 'categoría', 'descripcion', 'descripción', 'item', 'artículo', 'articulo'])), None)
                    if col_prod:
                        productos = sorted([str(x) for x in df[col_prod].dropna().unique().tolist() if str(x).strip() != ''])

                    num_cols = df.select_dtypes(include=['number']).columns.tolist()
                    col_monto = next((c for c in num_cols if any(k in c for k in ['monto', 'total', 'venta', 'valor', 'precio', 'importe', 'neto'])), num_cols[-1] if num_cols else None)

                    if not col_monto:
                        raise Exception("No se encontró ninguna columna numérica para realizar los cálculos.")

                    res = duckdb.query(f"SELECT SUM({col_monto}) as total_v, COUNT(*) as total_r FROM df").fetchone()
                    total_v = float(res[0]) if res[0] is not None else 0.0
                    total_r = int(res[1]) if res[1] is not None else 0
                    prom = total_v / total_r if total_r > 0 else 0
                    
                    kpis = {
                        'total_ventas': f"${total_v:,.2f}",
                        'total_remisiones': f"{total_r:,}",
                        'promedio_venta': f"${prom:,.2f}"
                    }

                    if col_fecha and col_fecha in df and col_monto in df:
                        df_mes = df.dropna(subset=[col_fecha]).copy()
                        df_mes['mes_num'] = df_mes[col_fecha].dt.month
                        df_mes['mes_nombre'] = df_mes[col_fecha].dt.strftime('%b')
                        grp_mes = df_mes.groupby(['mes_num', 'mes_nombre'])[col_monto].sum().reset_index().sort_values('mes_num')
                        labels_mes = grp_mes['mes_nombre'].tolist()
                        valores_mes = grp_mes[col_monto].tolist()

                    if col_prod and col_prod in df and col_monto in df:
                        grp_p = df.groupby(col_prod)[col_monto].sum().head(5)
                        labels_prod = [str(x) for x in grp_p.index.tolist()]
                        valores_prod = [float(x) for x in grp_p.values.tolist()]

                except Exception as e:
                    error = f"Error al procesar el archivo Excel: {str(e)}"
            else:
                error = "Por favor, sube un archivo con extensión .xlsx o .xls."

    return render_template(
        "cargar_excel.html", 
        error=error, 
        kpis=kpis, 
        estaciones=estaciones, 
        anios=anios,
        productos=productos,
        json_labels_mes=json.dumps(labels_mes),
        json_valores_mes=json.dumps(valores_mes),
        json_labels_prod=json.dumps(labels_prod),
        json_valores_prod=json.dumps(valores_prod)
    )

@app.route("/api/filtrar-analisis", methods=["POST"])
def filtrar_analisis():
    user = session.get("usuario", "DICKSON")
    if user not in DATA_STORE:
        return jsonify({'error': 'No hay datos cargados'}), 400

    df = DATA_STORE[user]
    data = request.get_json() or {}
    estacion_sel = data.get('estacion')
    anio_sel = data.get('anio')
    prod_sel = data.get('producto')

    cols = df.columns.tolist()
    col_estacion = next((c for c in cols if any(k in c for k in ['estacion', 'estación', 'eds', 'sede', 'zona', 'centro', 'local', 'punto', 'cliente'])), None)
    col_fecha = next((c for c in cols if any(k in c for k in ['fecha', 'date', 'dia', 'día', 'fec'])), None)
    text_cols = df.select_dtypes(include=['object']).columns.tolist()
    col_prod = next((c for c in text_cols if any(k in c for k in ['producto', 'categoria', 'categoría', 'descripcion', 'descripción', 'item', 'artículo', 'articulo'])), None)

    query = "SELECT * FROM df WHERE 1=1"
    if col_estacion and estacion_sel and estacion_sel not in ['Todas', 'Todas las Estaciones']:
        query += f" AND CAST({col_estacion} AS VARCHAR) = '{estacion_sel}'"
        
    if col_fecha and anio_sel and anio_sel not in ['Todos', 'Todos los Años']:
        query += f" AND STRFTIME({col_fecha}, '%Y') = '{anio_sel}'"

    if col_prod and prod_sel and prod_sel not in ['Todos', 'Todos los Productos']:
        query += f" AND CAST({col_prod} AS VARCHAR) = '{prod_sel}'"

    df_filtered = duckdb.query(query).df()

    num_cols = df_filtered.select_dtypes(include=['number']).columns.tolist()
    col_monto = next((c for c in num_cols if any(k in c for k in ['monto', 'total', 'venta', 'valor', 'precio', 'importe', 'neto'])), num_cols[-1] if num_cols else cols[0])

    total_v = float(df_filtered[col_monto].sum()) if col_monto in df_filtered else 0.0
    total_r = len(df_filtered)
    prom = total_v / total_r if total_r > 0 else 0.0

    labels_mes, valores_mes = [], []
    if col_fecha and col_fecha in df_filtered and col_monto in df_filtered:
        df_mes = df_filtered.dropna(subset=[col_fecha]).copy()
        df_mes['mes_num'] = pd.to_datetime(df_mes[col_fecha]).dt.month
        df_mes['mes_nombre'] = pd.to_datetime(df_mes[col_fecha]).dt.strftime('%b')
        grp_mes = df_mes.groupby(['mes_num', 'mes_nombre'])[col_monto].sum().reset_index().sort_values('mes_num')
        labels_mes = grp_mes['mes_nombre'].tolist()
        valores_mes = grp_mes[col_monto].tolist()

    labels_prod, valores_prod = [], []
    if col_prod and col_prod in df_filtered and col_monto in df_filtered:
        grp_p = df_filtered.groupby(col_prod)[col_monto].sum().head(5)
        labels_prod = [str(x) for x in grp_p.index.tolist()]
        valores_prod = [float(x) for x in grp_p.values.tolist()]

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