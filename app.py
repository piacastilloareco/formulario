from flask import Flask, render_template, request, send_file
import pandas as pd
import io
from datetime import datetime

app = Flask(__name__)

# Servicios fijos por nivel
SERVICIOS_FIJOS_POR_NIVEL = {
    "360": ["Datos Generales", "Datos de contacto", "Número de registro / Información fiscal", "Sector de actividad", "Centros", "Términos y Condiciones", "Configuración IT"],
    "180": ["Datos Generales", "Datos de contacto", "Número de registro / Información fiscal", "Sector de actividad", "Centros", "Términos y Condiciones", "Configuración IT"],
    "basic": ["Datos Generales", "Datos de contacto", "Número de registro / Información fiscal", "Sector de actividad", "Centros", "Términos y Condiciones", "Configuración IT"],
    "elementary": ["Datos Generales", "Datos de contacto", "Número de registro / Información fiscal", "Sector de actividad", "Centros", "Términos y Condiciones", "Configuración IT"],
    "digital": ["Datos Generales", "Número de registro / Información fiscal", "Sector de actividad", "Configuración IT"]
}

@app.route('/')
def index():
    return render_template('formulario.html')

@app.route('/submit', methods=['POST'])
def submit():
    # Obtener el nivel seleccionado
    nivel = request.form.get('nivel_cuestionario', '').lower()

    # Obtener los servicios seleccionados del formulario
    servicios_seleccionados = request.form.getlist('servicios')

    # Agregar los servicios fijos correspondientes al nivel
    servicios_fijos = SERVICIOS_FIJOS_POR_NIVEL.get(nivel, [])
    servicios_finales = list(set(servicios_seleccionados + servicios_fijos))  # Eliminar duplicados

    # Crear un DataFrame con los servicios finales
    servicios_df = pd.DataFrame({'Check Name': servicios_finales})

    # Obtener datos del cliente
    nombre_cliente = request.form.get('nombre_cliente', 'N/A')
    liderado_por = request.form.get('liderado_por', 'N/A')
    fecha_generacion = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    # Crear un DataFrame con los datos adicionales
    metadata_df = pd.DataFrame({
        'Concepto': ['Nombre del cliente', 'Liderado por', 'Fecha de generación'],
        'Data': [nombre_cliente, liderado_por, fecha_generacion]
    })

    # Crear un buffer para generar el archivo Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        servicios_df.to_excel(writer, sheet_name='Lista de Servicios', index=False)
        metadata_df.to_excel(writer, sheet_name='Información General', index=False)

    output.seek(0)

    # Enviar el archivo Excel como respuesta para descarga
    return send_file(output, download_name="lista_servicios.xlsx", as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
