from flask import Flask, request, render_template, send_file, after_this_request
from docxtpl import DocxTemplate, InlineImage
from io import BytesIO
from docx.shared import Inches
import os

app = Flask(__name__)

@app.route('/')
def index():
    # Muestra el formulario HTML al usuario
    return render_template('index.html')

@app.route('/generate_report', methods=['POST'])
def generate_report():
    # Obtener los datos del formulario
    datos_formulario = {
        'NombreCliente': request.form.get('nombre_cliente'),
        'DireccionEmpresa': request.form.get('direccion_empresa'),
        'CiudadEmpresa': request.form.get('ciudad_empresa'),
        'bateria': request.form.get('bateria'),
        'alternador': request.form.get('alternador'),
        'mbateria': request.form.get('mbateria'),
        'vl1': request.form.get('vl1'),
        'vl2': request.form.get('vl2'),
        'vl3': request.form.get('vl3'),
        'vmono': request.form.get('vmono'),
        'frecuencia': request.form.get('frecuencia'),
        'rpm': request.form.get('rpm'),
        'aceite': request.form.get('aceite'),
        'temp': request.form.get('temp'),
        'potenciacarga': request.form.get('potenciacarga'),
        'intensidad': request.form.get('intensidad'),
        'calefactor': request.form.get('calefactor'),
        'nivelcombustible': request.form.get('nivelcombustible'),
        'hfuncionamiento': request.form.get('hfuncionamiento'),
        'MARCAMODELOGE': request.form.get('MARCAMODELOGE'),
        'MARCAMOTOR': request.form.get('MARCAMOTOR'),
        'MODELOMOTOR': request.form.get('MODELOMOTOR'),
        'POTENCIAM': request.form.get('POTENCIAM'),
        'SERIE': request.form.get('SERIE'),
        'tecnico': request.form.get('tecnico'),
        'fecha': request.form.get('fecha'),
        'trabajorealizado': request.form.get('trabajorealizado'),
        'sugerencias': request.form.get('sugerencias')
    }

    plantilla_path = "C:\\Users\\Programador\\Desktop\\python\\daniela\\plantilla_base.docx"
    plantilla = DocxTemplate(plantilla_path)

    # Manejo de la imagen principal
    imagen_reporte = request.files.get('imagen_reporte')
    if imagen_reporte:
        image_stream = BytesIO(imagen_reporte.read())
        imagen = InlineImage(plantilla, image_stream, width=Inches(6), height=Inches(8.5))
        datos_formulario['ImagenReporte'] = imagen

    # Manejo de imágenes adicionales
    imagenes_adicionales = request.files.getlist('imagenes_adicionales')
    imagenes_contexto = []
    for imagen in imagenes_adicionales:
        if imagen:
            image_stream = BytesIO(imagen.read())
            imagen_inline = InlineImage(plantilla, image_stream, width=Inches(4.8), height=Inches(4.5))
            imagenes_contexto.append(imagen_inline)

    datos_formulario['ImagenesAdicionales'] = imagenes_contexto

    # Renderizar la plantilla con el contexto
    plantilla.render(datos_formulario)

    # Guardar el informe generado
    informe_generado_path = 'C:\\Users\\Programador\\Desktop\\python\\daniela\\informe_generado.docx'
    plantilla.save(informe_generado_path)

    # Eliminar el archivo después de enviarlo
    @after_this_request
    def remove_file(response):
        try:
            os.remove(informe_generado_path)
        except Exception as error:
            app.logger.error(f"Error eliminando archivo generado: {error}")
        return response

    # Envía el archivo generado para su descarga
    return send_file(informe_generado_path, as_attachment=True, download_name='Informe_Generado.docx')

if __name__ == '__main__':
    app.run(debug=True)
