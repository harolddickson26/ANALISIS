from flask import Flask, render_template, request, redirect, url_for, session, jsonify
import pandas as pd
import duckdb
import os

app = Flask(__name__)
app.secret_key = "clave_secreta_super_segura"

USUARIO_CORRECTO = "DICKSON"
PASSWORD_CORRECTO = "1234"

# Variable global para mantener el DataFrame en sesión/memoria durante el análisis
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
                    
                    # Normalizar nombres de columnas a minúsculas para consultas DuckDB limpias
                    df_global.columns = [str(col).strip().lower() for col in df_global.columns]
                    
                    # Consulta inicial de listas desplegables con DuckDB
                    con = duckdb.connect(database=':memory:')
                    con.register('datos', df_global)
                    
                    # Extraer estaciones únicas (si existe la columna 'estacion' o similar)
                    cols = df_global.columns.tolist()
                    if 'estacion' in cols:
                        estaciones = [r[0] for r in con.execute("SELECT DISTINCT estacion FROM datos WHERE estacion IS NOT NULL").fetchall()]
                    
                    # Extraer años únicos si existe columna de fecha
                    if 'fecha' in cols:
                        anios = [r[0] for r in con.execute("SELECT DISTINCT YEAR(CAST(fecha AS DATE)) FROM datos WHERE fecha IS NOT NULL ORDER BY 1 DESC").fetchall()]
                    
                    # KPIs iniciales
                    col_monto = 'monto' if 'monto' in cols else ('venta' if 'venta' in cols else cols[-1])
                    total_v = con.execute(f"SELECT SUM({col_monto}) FROM datos").fetchone()[0] or 0
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

    return render_template("cargar_excel.html", error=error, kpis=kpis, estaciones=estaciones, anios=anios)

# --- API PARA FILTRAR DESDE LAS LISTAS DESPLEGABLES ---
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

    condiciones = []
    if 'estacion' in cols and estacion_sel and estacion_sel != 'Todas':
        condiciones.append(f"estacion = '{estacion_sel}'")
    if 'fecha' in cols and anio_sel and anio_sel != 'Todos':
        condiciones.append(f"YEAR(CAST(fecha AS DATE)) = {anio_sel}")

    where_clause = " WHERE " + " AND ".join(condiciones) if condiciones else ""

    col_monto = 'monto' if 'monto' in cols else ('venta' if 'venta' in cols else cols[-1])
    col_prod = 'producto' if 'producto' in cols else ('categoria' if 'categoria' in cols else cols[0])

    # Consultas filtradas
    total_v = con.execute(f"SELECT COALESCE(SUM({col_monto}), 0) FROM datos {where_clause}").fetchone()[0]
    total_r = con.execute(f"SELECT COUNT(*) FROM datos {where_clause}").fetchone()[0]
    prom = total_v / total_r if total_r > 0 else 0

    # Datos Gráfico de Dona
    res_prod = con.execute(f"SELECT {col_prod}, SUM({col_monto}) FROM datos {where_clause} GROUP BY {col_prod} LIMIT 5").fetchall()
    
    # Datos Gráfico de Líneas (Tendencia)
    if 'fecha' in cols:
        res_mes = con.execute(f"SELECT STRFTIME(CAST(fecha AS DATE), '%m-%b'), SUM({col_monto}) FROM datos {where_clause} GROUP BY 1 ORDER BY 1").fetchall()
    else:
        res_mes = []

    return jsonify({
        'kpi_total': f"${total_v:,.2f}",
        'kpi_remisiones': f"{total_r:,}",
        'kpi_promedio': f"${prom:,.2f}",
        'labels_prod': [r[0] for r in res_prod],
        'valores_prod': [r[1] for r in res_prod],
        'labels_mes': [r[0] for r in res_mes],
        'valores_mes': [r[1] for r in res_mes]
    })

if __name__ == "__main__":
    app.run(debug=True)