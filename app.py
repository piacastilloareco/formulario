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

# Mapeo de servicios con sus códigos
MAPEO_SERVICIOS_CODIGOS = {
    "Datos Generales": "F",
    "Datos de contacto": "G",
    "Número de registro / Información fiscal": "I",
    "Sector de actividad": "J",
    "Centros": "L",
    "Términos y Condiciones": "M",
    "Configuración IT": "N",
    "Modelo Completo Enriquecido (Con Documento)": "A",
    "Modelo Reducido Enriquecido (Con documento)": "C",
    "Modelo Reducido No Enriquecido (Sin Documento)": "AZ",
    "Modelo Mínimo No Enriquecido (Sin documento)": "E",
    "Operacional con documento": "AM",
    "Operacional sin documento": "AN",
    "Accidentabilidad": "AO",
    "Ciberseguridad Completo": "AP",
    "Ciberseguridad Basico": "AQ",
    "Newsclipping": "BA",
    "Geopolítico": "T",
    "Observaciones": "AU",
    "ESG Predictivo": "BA",
    "Personas de contacto para licitaciones": "H",
    "ESG Transversal": "AB",
    "Obligaciones Tributarias con documento":"AR",
    "Obligaciones Tributarias sin documento": "AS",
    "ESG Intermedio": "AG",
    "ESG Completo": "AC",
    "ESG Basico":"AJ",
    "Datos bancarios sin documento": "AW",
    "Datos bancarios con documento": "R",
    "Poliza de seguro con documento" : "Q",
    "Poliza de seguro sin documento" : "AV",
    "Onlycompany" : "T",
    "Onlycompany + Politicas": "U",
    "Stakeholders": "V",
    "Stakeholders + Politicas" : "W", 
    "Stakeholders + Peps y Sips" : "X",
    "Stakeholders + Politicas + Peps y Sips": "Y",

 
    # Agrega el resto de servicios y códigos aquí
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

    # Crear un DataFrame con los servicios finales y sus códigos
    servicios_df = pd.DataFrame({
        'Check Name': servicios_finales,
        'Codigo': [MAPEO_SERVICIOS_CODIGOS.get(servicio, "N/A") for servicio in servicios_finales]
    })

    # Obtener datos del cliente
    nombre_cliente = request.form.get('nombre_cliente', 'N/A')
    liderado_por = request.form.get('liderado_por', 'N/A')
    fecha_generacion = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    # Crear un DataFrame con los datos adicionales
    metadata_df = pd.DataFrame({
        'Concepto': ['Nombre del cliente', 'Liderado por', 'Fecha de generación', 'Nivel del cuestionario'],
        'Data': [nombre_cliente, liderado_por, fecha_generacion, nivel.capitalize()]
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
