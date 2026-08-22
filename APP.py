from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)
# Clave secreta requerida por Flask para manejar sesiones seguras
app.secret_key = "clave_secreta_super_segura"

# Credenciales de prueba
USUARIO_CORRECTO = "DianaPaulina"
PASSWORD_CORRECTO = "3127835548"

@app.route("/")
def inicio():
    # Si el usuario ya inició sesión, muestra la página principal
    if "usuario" in session:
        return render_template("index.html", usuario=session["usuario"])
    # Si no ha iniciado sesión, redirige al login
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

# --- RUTAS DE LAS SECCIONES DEL MENÚ ---

@app.route("/inventario")
def inventario():
    if "usuario" not in session: return redirect(url_for("login"))
    return render_template("seccion.html", titulo="Inventario", contenido="Gestión e historial del inventario de productos.")

@app.route("/ingresar-entrada")
def ingresar_entrada():
    if "usuario" not in session: return redirect(url_for("login"))
    return render_template("seccion.html", titulo="Ingresar Entrada", contenido="Registro de nuevas entradas de mercancía.")

@app.route("/ingresar-venta")
def ingresar_venta():
    if "usuario" not in session: return redirect(url_for("login"))
    return render_template("seccion.html", titulo="Ingresar Venta", contenido="Registro de nuevas ventas.")

@app.route("/utilidad")
def utilidad():
    if "usuario" not in session: return redirect(url_for("login"))
    return render_template("seccion.html", titulo="Utilidad", contenido="Reporte de ganancias y utilidades.")

@app.route("/total-ventas")
def total_ventas():
    if "usuario" not in session: return redirect(url_for("login"))
    return render_template("seccion.html", titulo="Total Ventas", contenido="Resumen total del volumen de ventas.")

@app.route("/total-entradas")
def total_entradas():
    if "usuario" not in session: return redirect(url_for("login"))
    return render_template("seccion.html", titulo="Total Entradas", contenido="Resumen total de las entradas registradas.")

if __name__ == "__main__":
    app.run(debug=True)