from flask import Flask, render_template, request, redirect, url_for, session, jsonify
import pandas as pd
import duckdb
import os

app = Flask(__name__)
app.secret_key = "clave_secreta_super_segura"

USUARIO_CORRECTO = "DICKSON"
PASSWORD_CORRECTO = "1234"

df_global = None

@app.route("/")
def inicio():
    if "usuario" in session:
        return render_template("index.html", usuario=session["usuario"])
    return redirect(url_for("login"))

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        usuario_ingresado = request.form["username"]
        password_ingresado = request.form["password"]

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

# --- RUTA PARA SUBIR Y ANALIZAR EXCEL ---
@app.route("/cargar-excel", methods=["GET", "POST"])
def cargar_excel():
    global df_global
    if "usuario" not in session: 
        return redirect(url_for("login"))
    
    error = None
    kpis = None
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
                    engine = "openpyxl" if file.filename.endswith(".xlsx") else "xlrd"
                    df_global = pd.read_excel(file, engine=engine)
                    
                    # Limpiar nombres de columnas
                    df_global.columns = [str(col).strip().lower() for col in df_global.columns]
                    
                    con = duckdb.connect(database=':memory:')
                    con.register('datos', df_global)
                    cols = df_global.columns.tolist()

                    # Detectar columna de estación/sede
                    col_estacion = next((c for c in cols if 'estacion' in c or 'sede' in c or 'zona' in c), None)
                    if col_estacion:
                        estaciones = [str(r[0]) for r in con.execute(f"SELECT DISTINCT {col_estacion} FROM datos WHERE {col_estacion} IS NOT NULL").fetchall()]

                    # Detectar columna de fecha
                    col_fecha = next((c for c in cols if 'fecha' in c or 'date' in c), None)
                    if col_fecha:
                        anios = [str(r[0]) for r in con.execute(f"SELECT DISTINCT YEAR(CAST({col_fecha} AS DATE)) FROM datos WHERE {col_fecha} IS NOT NULL ORDER BY 1 DESC").fetchall()]

                    # Detectar columna de monto
                    num_cols = df_global.select_dtypes(include=['number']).columns.tolist()
                    col_monto = next((c for c in num_cols if 'monto' in c or 'total' in c or 'venta' in c or 'valor' in c), num_cols[-1] if num_cols else None)

                    if not col_monto:
                        raise Exception("No se encontró ninguna columna numérica para calcular los indicadores.")

                    total_v = con.execute(f"SELECT COALESCE(SUM({col_monto}), 0) FROM datos").fetchone()[0]
                    total_r = len(df_global)
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

    # Siempre pasar estaciones, anios y kpis (incluso vacíos en GET) para evitar Server Error
    return render_template("cargar_excel.html", error=error, kpis=kpis, estaciones=estaciones, anios=anios)

# --- API DE FILTRADO PARA DESPLEGABLES ---
@app.route("/api/filtrar-analisis", methods=["POST"])
def filtrar_analisis():
    global df_global
    if df_global is None:
        return jsonify({'error': 'No hay datos cargados'}), 400

    data = request.get_json()
    estacion_sel = data.get('estacion')
    anio_sel = data.get('anio')

    con = duckdb.connect(database=':memory:')
    con.register('datos', df_global)
    cols = df_global.columns.tolist()

    col_estacion = next((c for c in cols if 'estacion' in c or 'sede' in c or 'zona' in c), None)
    col_fecha = next((c for c in cols if 'fecha' in c or 'date' in c), None)

    condiciones = []
    if col_estacion and estacion_sel and estacion_sel != 'Todas':
        condiciones.append(f"{col_estacion} = '{estacion_sel}'")
    if col_fecha and anio_sel and anio_sel != 'Todos':
        condiciones.append(f"YEAR(CAST({col_fecha} AS DATE)) = {anio_sel}")

    where_clause = " WHERE " + " AND ".join(condiciones) if condiciones else ""

    num_cols = df_global.select_dtypes(include=['number']).columns.tolist()
    col_monto = next((c for c in num_cols if 'monto' in c or 'total' in c or 'venta' in c or 'valor' in c), num_cols[-1] if num_cols else cols[0])
    
    text_cols = df_global.select_dtypes(include=['object']).columns.tolist()
    col_prod = next((c for c in text_cols if 'producto' in c or 'categoria' in c or 'descripcion' in c), text_cols[0] if text_cols else cols[0])

    total_v = con.execute(f"SELECT COALESCE(SUM({col_monto}), 0) FROM datos {where_clause}").fetchone()[0]
    total_r = con.execute(f"SELECT COUNT(*) FROM datos {where_clause}").fetchone()[0]
    prom = total_v / total_r if total_r > 0 else 0

    res_prod = con.execute(f"SELECT {col_prod}, SUM({col_monto}) FROM datos {where_clause} GROUP BY {col_prod} LIMIT 5").fetchall()
    
    if col_fecha:
        res_mes = con.execute(f"SELECT STRFTIME(CAST({col_fecha} AS DATE), '%m-%b'), SUM({col_monto}) FROM datos {where_clause} GROUP BY 1 ORDER BY 1").fetchall()
    else:
        res_mes = []

    return jsonify({
        'kpi_total': f"${total_v:,.2f}",
        'kpi_remisiones': f"{total_r:,}",
        'kpi_promedio': f"${prom:,.2f}",
        'labels_prod': [str(r[0]) for r in res_prod],
        'valores_prod': [r[1] for r in res_prod],
        'labels_mes': [str(r[0]) for r in res_mes],
        'valores_mes': [r[1] for r in res_mes]
    })

if __name__ == "__main__":
    app.run(debug=True)