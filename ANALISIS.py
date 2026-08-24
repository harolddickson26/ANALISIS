from flask import Flask, render_template, request, redirect, url_for, session, jsonify
import pandas as pd
import duckdb

app = Flask(__name__)
app.secret_key = "clave_secreta_analisis_app"

USUARIO_CORRECTO = "DICKSON"
PASSWORD_CORRECTO = "1234"

# Memoria global para el DataFrame y metadatos por usuario
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

# PASO 1: Carga del Archivo Excel
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
                    # Carga ligera y limpia de encabezados
                    if file.filename.endswith(".xlsx"):
                        df = pd.read_excel(file, engine="calamine")
                    else:
                        df = pd.read_excel(file, engine="xlrd")
                    
                    df.columns = [str(col).strip() for col in df.columns]
                    
                    # Guardar DataFrame y estructura
                    DATA_STORE[user_key] = {
                        'df': df,
                        'filename': file.filename,
                        'columnas': df.columns.tolist(),
                        'total_filas': len(df)
                    }
                    
                    # Redirigir al Paso 2: Selección con listas desplegables
                    return redirect(url_for("paso2_columnas"))

                except Exception as e:
                    error = f"Error al procesar el archivo Excel: {str(e)}"
            else:
                error = "Por favor, sube un archivo con extensión .xlsx o .xls."

    return render_template("cargar_excel.html", error=error)

# PASO 2: Selección y confirmación de columnas por listas desplegables
@app.route("/paso2-columnas", methods=["GET", "POST"])
def paso2_columnas():
    if not session.get("usuario"):
        return redirect(url_for("login"))

    user_key = session.get("usuario", "DICKSON")
    if user_key not in DATA_STORE:
        return redirect(url_for("cargar_excel"))

    data = DATA_STORE[user_key]

    if request.method == "POST":
        # Guardar las selecciones de listas desplegables para usar en los KPIs del Paso 3
        DATA_STORE[user_key]['mapeo_columnas'] = {
            'monto': request.form.get("col_monto"),
            'cantidad': request.form.get("col_cantidad"),
            'incremento': request.form.get("col_incremento"),
            'fecha': request.form.get("col_fecha"),
            'estacion': request.form.get("col_estacion"),
            'producto': request.form.get("col_producto")
        }
        return redirect(url_for("paso2_columnas"))

    # Renderizado apuntando a la nueva plantilla
    return render_template(
        "Lista_despegable_columna.html",
        columnas=data['columnas'],
        filename=data['filename'],
        total_filas=f"{data['total_filas']:,}",
        total_columnas=len(data['columnas'])
    )

if __name__ == "__main__":
    app.run(debug=True)