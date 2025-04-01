#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Aplicación para generar un cuadro de mando a partir de datos en Excel
"""
import argparse
import os
import sys
import webbrowser
from pathlib import Path
from datetime import datetime  # Añadir esta importación
import json  # Añadir importación de json

# Importar los módulos de nuestra aplicación
from excel_extractor import ExcelDataExtractor
from data_processor_full import DataProcessor
from html_generator_full import HtmlGenerator

# Configuración predeterminada específica para tu proyecto
DEFAULT_EXCEL_PATH = "/Users/didac/Downloads/StoryMac/DashBTracker/PruebasCdM/Tchart_V06.xlsm"
DEFAULT_HISTORIC_SHEET = "FrmBB_2"  # Actualizado para usar la hoja correcta
DEFAULT_KPI_SHEET = "FrmBB_3"       # Actualizado para usar la hoja correcta
DEFAULT_OUTPUT_HTML = "dashboard_output.html"
DEFAULT_TEMPLATE = "dashboard_template_fixed.html"  # Usar la nueva plantilla

def main():
    """Función principal de la aplicación"""
    parser = argparse.ArgumentParser(description='Genera un cuadro de mando a partir de un archivo Excel')
    parser.add_argument('--excel', '-e', default=DEFAULT_EXCEL_PATH, 
                      help=f'Ruta al archivo Excel (por defecto: {DEFAULT_EXCEL_PATH})')
    parser.add_argument('--historic', '-i', default=DEFAULT_HISTORIC_SHEET, 
                      help=f'Nombre de la hoja con datos históricos (por defecto: {DEFAULT_HISTORIC_SHEET})')
    parser.add_argument('--kpi', '-k', default=DEFAULT_KPI_SHEET, 
                      help=f'Nombre de la hoja con datos de KPIs (por defecto: {DEFAULT_KPI_SHEET})')
    parser.add_argument('--template', '-t', default=DEFAULT_TEMPLATE, 
                      help=f'Plantilla HTML a utilizar (por defecto: {DEFAULT_TEMPLATE})')
    parser.add_argument('--output', '-o', default=DEFAULT_OUTPUT_HTML,
                      help=f'Nombre del archivo HTML de salida (por defecto: {DEFAULT_OUTPUT_HTML})')
    parser.add_argument('--open', '-b', action='store_true', help='Abrir el cuadro de mando en el navegador')
    
    args = parser.parse_args()
    
    # Verificar que el archivo Excel existe
    if not os.path.exists(args.excel):
        print(f"Error: El archivo {args.excel} no existe.")
        sys.exit(1)
    
    # Verificar que la plantilla HTML existe
    template_path = args.template
    if not os.path.exists(template_path):
        # Buscar en el directorio de la aplicación
        app_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(app_dir, args.template)
        if not os.path.exists(template_path):
            print(f"Error: La plantilla HTML {args.template} no existe.")
            sys.exit(1)
    
    try:
        # Extraer datos históricos del Excel
        print(f"Extrayendo datos históricos de {args.excel}, hoja: {args.historic}...")
        excel_extractor = ExcelDataExtractor(args.excel)
        historic_data = excel_extractor.extract_data(args.historic)
        
        # Extraer datos de KPIs del Excel
        print(f"Extrayendo datos de KPIs de {args.excel}, hoja: {args.kpi}...")
        kpi_data = excel_extractor.extract_data(args.kpi)
        
        # Procesar los datos históricos
        print("Procesando datos históricos...")
        historic_processor = DataProcessor(historic_data)
        historic_processed = historic_processor.process()
        
        # Procesar los datos de KPIs
        print("Procesando datos de KPIs...")
        kpi_processor = DataProcessor(kpi_data, is_kpi_data=True)
        kpi_processed = kpi_processor.process()
        
        # Combinar los datos procesados
        print("Combinando datos históricos y KPIs...")
        combined_data = combine_data(historic_processed, kpi_processed)
        
        # Generar el HTML
        print(f"Generando HTML utilizando la plantilla {template_path}...")
        generator = HtmlGenerator(template_path)
        html_content = generator.generate(combined_data)
        
        # Verificar que el HTML generado no contiene duplicados o código basura
        if html_content.count('<div class="dashboard-container">') > 1:
            print("Advertencia: Se detectaron elementos duplicados en el HTML generado. Limpiando...")
            # Limpiar manualmente si es necesario
            first_container = html_content.find('<div class="dashboard-container">')
            if first_container != -1:
                second_container = html_content.find('<div class="dashboard-container">', first_container + 1)
                if second_container != -1:
                    # Encontrar el final del documento
                    end_html = html_content.find('</html>')
                    if end_html != -1:
                        html_content = html_content[:second_container] + html_content[end_html:]
                    else:
                        html_content = html_content[:second_container] + "</div></body></html>"
        
        # Verificar y reemplazar variables de plantilla no sustituidas
        placeholders = ['${TITLE}', '${GENERATED_DATE}', '{{DASHBOARD_TITLE}}', '{{GENERATED_DATE}}', '{{FILTERS}}', '{{INFO}}', '{{CELLS}}']
        for placeholder in placeholders:
            if placeholder in html_content:
                print(f"Advertencia: La variable de plantilla {placeholder} no fue reemplazada.")
                if placeholder == '${TITLE}' or placeholder == '{{DASHBOARD_TITLE}}':
                    html_content = html_content.replace(placeholder, 'Cuadro de Mando Excel')
                elif placeholder == '${GENERATED_DATE}' or placeholder == '{{GENERATED_DATE}}':
                    html_content = html_content.replace(placeholder, f"Generado el: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
                else:
                    html_content = html_content.replace(placeholder, '')
        
        # Guardar el HTML generado
        output_path = args.output
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"Cuadro de mando generado correctamente: {output_path}")
        
        # Abrir en el navegador si se solicitó
        if args.open:
            # Convertir a ruta absoluta
            abs_path = os.path.abspath(output_path)
            url = f"file://{abs_path}"
            print(f"Abriendo en el navegador: {url}")
            webbrowser.open(url)
        
        return 0  # Éxito
        
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1  # Error

# Modificar la función combine_data para aclarar la relación entre KPIs e históricos
def combine_data(historic_data, kpi_data):
    """
    Combina los datos históricos y los datos de KPIs manteniendo su independencia
    
    Args:
        historic_data (dict): Datos históricos procesados
        kpi_data (dict): Datos de KPIs procesados
        
    Returns:
        dict: Datos combinados
    """
    # Crear una estructura de datos combinada
    combined = {
        'categories': historic_data.get('categories', []),
        'subcategories': historic_data.get('subcategories', []),
        'cellData': {},
        'hasKpis': True,
        'historicData': historic_data.get('cellData', {}),
        'kpiData': kpi_data.get('cellData', {})
    }
    
    # Extraer y contar proyectos y compañías
    projects = set()
    companies = set()
    
    # Extraer de datos históricos
    for cell_key, cell_data in historic_data.get('cellData', {}).items():
        if 'PRJID' in cell_data:
            projects.add(str(cell_data['PRJID']))
        if 'CIA' in cell_data:
            companies.add(str(cell_data['CIA']))
    
    # Extraer de datos KPI
    for cell_key, cell_data in kpi_data.get('cellData', {}).items():
        if 'PRJID' in cell_data:
            projects.add(str(cell_data['PRJID']))
        if 'CIA' in cell_data:
            companies.add(str(cell_data['CIA']))
    
    # Guardar en la estructura combinada
    combined['projects'] = sorted(list(projects))
    combined['companies'] = sorted(list(companies))
    combined['hasProjects'] = len(projects) > 0
    combined['hasCompanies'] = len(companies) > 0
    
    # Imprimir información de proyectos y compañías
    print(f"Total de proyectos encontrados: {len(projects)}")
    if projects:
        print(f"Proyectos: {', '.join(sorted(projects))}")
    
    print(f"Total de compañías encontradas: {len(companies)}")
    if companies:
        print(f"Compañías: {', '.join(sorted(companies))}")
    
    # Combinar categorías y subcategorías de ambas fuentes
    if 'categories' in kpi_data:
        for category in kpi_data['categories']:
            if category not in combined['categories']:
                combined['categories'].append(category)
    
    if 'subcategories' in kpi_data:
        for subcategory in kpi_data['subcategories']:
            if subcategory not in combined['subcategories']:
                combined['subcategories'].append(subcategory)
    
    # Crear un mapa de correspondencia entre claves de KPI y claves de histórico
    # basado únicamente en la jerarquía (categoría y subcategoría)
    kpi_to_historic_map = {}
    
    # Mapear por categoría y subcategoría (usando los nombres exactos)
    for kpi_key, kpi_cell in combined['kpiData'].items():
        if 'Category' not in kpi_cell or 'Subcategory' not in kpi_cell:
            continue
            
        kpi_category = kpi_cell['Category']
        kpi_subcategory = kpi_cell['Subcategory']
        
        for historic_key, historic_cell in combined['historicData'].items():
            if 'Category' not in historic_cell or 'Subcategory' not in historic_cell:
                continue
                
            historic_category = historic_cell['Category']
            historic_subcategory = historic_cell['Subcategory']
            
            if kpi_category == historic_category and kpi_subcategory == historic_subcategory:
                kpi_to_historic_map[kpi_key] = historic_key
                break
    
    # Crear celdas combinadas a partir de datos históricos
    for historic_key, historic_cell in combined['historicData'].items():
        cell_data = historic_cell.copy()
        cell_data['source'] = 'historic'
        
        # Asegurarse de que la celda tiene una serie temporal
        if 'timeSeries' not in cell_data:
            cell_data['timeSeries'] = []
        
        # Verificar que la serie temporal tiene datos
        if not cell_data['timeSeries'] and 'PREV' in cell_data:
            # Si no hay datos históricos pero hay valores, crear una serie temporal básica
            # Usamos prefijo 'h' para PREV para distinguirlo del PREV de KPIs
            cell_data['timeSeries'] = [
                {'period': 'Anterior', 'value': cell_data.get('PREV', 0), 'source': 'hPREV'},
                {'period': 'Actual', 'value': cell_data.get('PPTO', 0), 'source': 'PPTO'},
                {'period': 'Proyección', 'value': cell_data.get('PDTE', 0), 'source': 'PDTE'}
            ]
    
            # Asegurarse de que los valores de la serie temporal son números
            if 'timeSeries' in cell_data and cell_data['timeSeries']:
                for item in cell_data['timeSeries']:
                    if 'value' in item:
                        try:
                            item['value'] = float(item['value'])
                        except (ValueError, TypeError):
                            item['value'] = 0.0
        
        combined['cellData'][historic_key] = cell_data
    
    # Añadir celdas de KPI que no tienen correspondencia en histórico
    for kpi_key, kpi_cell in combined['kpiData'].items():
        if kpi_key not in kpi_to_historic_map and kpi_key not in combined['cellData']:
            cell_data = kpi_cell.copy()
            cell_data['source'] = 'kpi'
            
            # Los KPIs no tienen series temporales, son datos instantáneos
            cell_data['timeSeries'] = []
            
            # Guardar los valores KPI en la celda usando los nombres originales
            # Usamos prefijo 'k' para PREV para distinguirlo del PREV histórico
            kpi_values = {
                'kPREV': kpi_cell.get('PREV', 0),
                'REALPREV': kpi_cell.get('REALPREV', 0),
                'PPTOPREV': kpi_cell.get('PPTOPREV', 0),
                'PDTE': kpi_cell.get('PDTE', 0)
            }
            
            # Convertir a números
            for k, v in kpi_values.items():
                try:
                    kpi_values[k] = float(v)
                except (ValueError, TypeError):
                    kpi_values[k] = 0.0
            
            # Guardar los valores KPI en la celda
            cell_data['kpiValues'] = kpi_values  # Cambiado de 'kpis' a 'kpiValues' para mayor claridad
            cell_data['hasKpiValues'] = True     # Cambiado de 'hasKpis' a 'hasKpiValues'
                
            combined['cellData'][kpi_key] = cell_data
    
    # Añadir datos de KPI a las celdas históricas correspondientes por jerarquía
    for kpi_key, historic_key in kpi_to_historic_map.items():
        kpi_cell = combined['kpiData'][kpi_key]
        
        if historic_key in combined['cellData']:
            # Añadir datos de KPI a la celda histórica como datos complementarios
            # pero independientes de los datos históricos
            kpi_values = {
                'kPREV': kpi_cell.get('PREV', 0),
                'REALPREV': kpi_cell.get('REALPREV', 0),
                'PPTOPREV': kpi_cell.get('PPTOPREV', 0),
                'PDTE': kpi_cell.get('PDTE', 0)
            }
            
            # Asegurarse de que los valores de KPI son números
            for k, v in kpi_values.items():
                try:
                    kpi_values[k] = float(v)
                except (ValueError, TypeError):
                    kpi_values[k] = 0.0
            
            # Guardar los valores KPI en la celda histórica
            combined['cellData'][historic_key]['kpiValues'] = kpi_values  # Cambiado de 'kpis' a 'kpiValues'
            combined['cellData'][historic_key]['hasKpiValues'] = True     # Cambiado de 'hasKpis' a 'hasKpiValues'
            combined['cellData'][historic_key]['kpiSource'] = kpi_key
    
    # Imprimir información de depuración
    print(f"Total de celdas combinadas: {len(combined['cellData'])}")
    print(f"Celdas con datos históricos: {sum(1 for cell in combined['cellData'].values() if cell.get('source') == 'historic')}")
    print(f"Celdas solo con KPIs: {sum(1 for cell in combined['cellData'].values() if cell.get('source') == 'kpi')}")
    print(f"Celdas con ambos tipos de datos: {sum(1 for cell in combined['cellData'].values() if cell.get('source') == 'historic' and cell.get('hasKpiValues', False))}")
    
    # Modificar la sección de depuración para que sea más clara
    # Imprimir ejemplo de una celda con datos históricos para depuración
    for key, cell in combined['cellData'].items():
        if 'timeSeries' in cell and cell['timeSeries']:
            print(f"Ejemplo de celda con histórico: {key}")
            print(f"  Series temporales: {cell['timeSeries'][:3]}...")
            if 'hasKpiValues' in cell and cell['hasKpiValues']:
                print(f"  También tiene valores KPI complementarios (no vinculados funcionalmente)")
            
            # Mostrar los campos disponibles sin separar entre originales y derivados
            print("  Campos disponibles:")
            for field_name, field_value in cell.items():
                if not isinstance(field_value, (dict, list)):
                    print(f"    {field_name}: {field_value}")
            break
    
    # Verificar la estructura final de los datos
    print("\nVerificación final de datos combinados:")
    
    # Verificar KPIs
    kpi_count = 0
    for key, cell in combined['cellData'].items():
        if 'kpiValues' in cell:  # Cambiado de 'kpis' a 'kpiValues'
            kpi_count += 1
            if kpi_count == 1:  # Solo mostrar el primer ejemplo
                print(f"Ejemplo de celda con valores KPI: {key}")
                print(f"  Valores KPI disponibles: {list(cell['kpiValues'].keys())}")
                print(f"  Valores: {cell['kpiValues']}")
    
    print(f"Total de celdas con valores KPI: {kpi_count}")
    
    # Verificar series temporales
    series_count = 0
    for key, cell in combined['cellData'].items():
        if 'timeSeries' in cell and cell['timeSeries'] and len(cell['timeSeries']) > 0:
            series_count += 1
            if series_count == 1:  # Solo mostrar el primer ejemplo
                print(f"Ejemplo de celda con series temporales: {key}")
                print(f"  Número de puntos: {len(cell['timeSeries'])}")
                print(f"  Periodos: {[item.get('period', '') for item in cell['timeSeries']]}")
                print(f"  Valores: {[item.get('value', 0) for item in cell['timeSeries']]}")
    
    print(f"Total de celdas con series temporales: {series_count}")
    
    return combined

if __name__ == "__main__":
    sys.exit(main())