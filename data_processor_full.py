#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Módulo para procesar y transformar los datos para el cuadro de mando
"""
import pandas as pd
import numpy as np
import json
from datetime import datetime

class DataProcessor:
    """Procesa los datos extraídos del Excel para prepararlos para el dashboard"""
    
    def __init__(self, data, is_kpi_data=False):
        """
        Inicializa el procesador de datos
        
        Args:
            data: Datos extraídos del Excel (puede ser DataFrame o lista de diccionarios)
            is_kpi_data (bool): Indica si los datos son de KPIs
        """
        self.is_kpi_data = is_kpi_data
        
        # Convertir lista a DataFrame si es necesario
        import pandas as pd
        if isinstance(data, list):
            self.df = pd.DataFrame(data)
        else:
            self.df = data
            
        # Continuar con la inicialización normal
        self.categories = []
        self.subcategories = []
        
    def process(self):
        """
        Procesa los datos para generar la estructura necesaria para el cuadro de mando
        
        Returns:
            dict: Datos procesados en formato adecuado para la visualización
        """
        try:
            # Verificar que se detectaron las columnas mínimas necesarias
            required_cols = ['ROW', 'COLUMN']
            
            if not self.is_kpi_data:
                required_cols.extend(['WKS', 'REAL'])
            else:
                # Para datos de KPIs, necesitamos al menos uno de estos campos
                kpi_fields = ['PREV', '%REAL/PREV', '%PPTO/PREV', 'PDTE']
                if not any(field in self.df.columns for field in kpi_fields):
                    raise ValueError(f"No se pudo detectar ninguna columna de KPI: {', '.join(kpi_fields)}")
            
            for col_type in required_cols:
                if col_type not in self.df.columns:
                    raise ValueError(f"No se pudo detectar la columna '{col_type}'")
            
            # Extraer columnas con nombres exactos
            cat_col = 'ROW'
            subcat_col = 'COLUMN'
            period_col = 'WKS'
            value_col = 'REAL'
            project_col = 'PRJID'
            company_col = 'CIA'
            ppto_col = 'PPTO'
            prev_col = 'PREV'
            
            # Columnas específicas para KPIs
            prev_value_col = 'PREV'
            real_prev_percent_col = '%REAL/PREV'
            ppto_prev_percent_col = '%PPTO/PREV'
            pending_value_col = 'PDTE'
            
            # Crear copia trabajar con ella
            df = self.df.copy()
            
            # Asegurar que las columnas necesarias existen
            required_cols_actual = [cat_col, subcat_col]
            if not self.is_kpi_data:
                required_cols_actual.extend([period_col, value_col])
            
            missing_cols = [col for col in required_cols_actual if col not in df.columns]
            if missing_cols:
                raise ValueError(f"Faltan columnas requeridas: {', '.join(missing_cols)}")
            
            # Eliminar filas con valores faltantes en columnas críticas
            df = df.dropna(subset=required_cols_actual)
            
            # Convertir valores a tipos adecuados (solo para datos históricos)
            if not self.is_kpi_data and value_col in df.columns:
                df[value_col] = pd.to_numeric(df[value_col], errors='coerce')
                # Eliminar filas con valores no numéricos
                df = df.dropna(subset=[value_col])
            
            # Convertir valores numéricos para KPIs
            if self.is_kpi_data:
                for col in [prev_value_col, real_prev_percent_col, ppto_prev_percent_col, pending_value_col]:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce')
            
            # Extraer listas únicas ordenadas
            categories = sorted(df[cat_col].unique().tolist())
            subcategories = sorted(df[subcat_col].unique().tolist())
            
            # Solo para datos históricos
            periods = []
            if not self.is_kpi_data and period_col in df.columns:
                periods = sorted(df[period_col].unique().tolist())
            
            # Si existe columna de proyecto, extraer proyectos únicos
            projects = []
            if project_col in df.columns:
                projects = sorted(df[project_col].unique().tolist())
            
            # Organizar datos para la visualización
            self.processed_data = {
                'categories': categories,
                'subcategories': subcategories,
                'periods': periods,
                'projects': projects,
                'hasProjects': len(projects) > 0,
                'hasPpto': ppto_col in df.columns,
                'hasPrev': prev_col in df.columns,
                'cellData': {}
            }
            
            # Configuración específica para KPIs
            if self.is_kpi_data:
                self.processed_data['isKpiData'] = True
                self.processed_data['hasPrevValue'] = prev_value_col in df.columns
                self.processed_data['hasRealPrevPercent'] = real_prev_percent_col in df.columns
                self.processed_data['hasPptoPrevPercent'] = ppto_prev_percent_col in df.columns
                self.processed_data['hasPendingValue'] = pending_value_col in df.columns
            
            # Añadir información de fechas y generación
            self.processed_data['generatedAt'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            # Procesar datos de celdas
            if not self.is_kpi_data:
                # Procesamiento para datos históricos
                self._process_historic_cell_data(df, cat_col, subcat_col, period_col, value_col, project_col, company_col, ppto_col, prev_col)
            else:
                # Procesamiento para datos de KPIs
                self._process_kpi_cell_data(df, cat_col, subcat_col, project_col, company_col, prev_value_col, real_prev_percent_col, ppto_prev_percent_col, pending_value_col)
            
            # Calcular totales
            self._calculate_totals()
            
            return self.processed_data
            
        except Exception as e:
            raise Exception(f"Error al procesar datos: {str(e)}")
    
    def _process_historic_cell_data(self, df, cat_col, subcat_col, period_col, value_col, project_col, company_col, ppto_col, prev_col):
        """
        Procesa los datos históricos de celdas para preparar la estructura del cuadro de mando
        
        Args:
            df (pandas.DataFrame): DataFrame a procesar
            cat_col, subcat_col, period_col, value_col: Nombres de columnas
            project_col, company_col, ppto_col, prev_col: Nombres de columnas opcionales
        """
        # Agrupar por las columnas de identificación y período
        group_cols = []
        
        # Añadir proyecto y compañía si existen
        if project_col in df.columns:
            group_cols.append(project_col)
        
        if company_col in df.columns:
            group_cols.append(company_col)
        
        # Añadir categoría, subcategoría y período
        group_cols.extend([cat_col, subcat_col, period_col])
        
        # Preparar diccionario para agrupar
        agg_dict = {value_col: 'sum'}
        
        if ppto_col in df.columns:
            agg_dict[ppto_col] = 'sum'
        
        if prev_col in df.columns:
            agg_dict[prev_col] = 'sum'
        
        # Agrupar con todas las columnas
        grouped = df.groupby(group_cols).agg(agg_dict).reset_index()
        
        # Procesar cada fila agrupada
        for _, row in grouped.iterrows():
            # Crear clave de celda usando todas las columnas de identificación
            cell_components = []
            
            if project_col in df.columns:
                cell_components.append(str(row[project_col]))
            
            if company_col in df.columns:
                cell_components.append(str(row[company_col]))
            
            cell_components.append(str(row[cat_col]))
            cell_components.append(str(row[subcat_col]))
            
            cell_key = "|".join(cell_components)
            
            # Inicializar datos de celda si no existe
            if cell_key not in self.processed_data['cellData']:
                initial_data = {
                    'Category': row[cat_col],
                    'Subcategory': row[subcat_col],
                    'timeSeries': [],
                    'periodsWithData': 0,
                    'lastValue': 0
                }
                
                if project_col in df.columns:
                    initial_data['PRJID'] = row[project_col]
                
                if company_col in df.columns:
                    initial_data['CIA'] = row[company_col]
                    
                self.processed_data['cellData'][cell_key] = initial_data
            
            # Preparar punto de serie temporal
            time_point = {
                'period': row[period_col],
                'value': float(row[value_col]) if not pd.isna(row[value_col]) else 0,
                'source': 'hREAL'
            }
            
            # Añadir valores adicionales si existen
            if ppto_col in df.columns:
                time_point['ppto'] = float(row[ppto_col]) if not pd.isna(row[ppto_col]) else 0
                time_point['source'] = 'hPPTO'
            
            if prev_col in df.columns:
                time_point['prev'] = float(row[prev_col]) if not pd.isna(row[prev_col]) else 0
                time_point['source'] = 'hPREV'
            
            # Añadir punto de serie temporal
            self.processed_data['cellData'][cell_key]['timeSeries'].append(time_point)
            
            # Incrementar contador de períodos
            self.processed_data['cellData'][cell_key]['periodsWithData'] += 1
        
        # Procesar cada celda para calcular últimos valores y métricas comparativas
        for cell_key, cell_data in self.processed_data['cellData'].items():
            # Ordenar series temporales por período (para tener el último período al final)
            cell_data['timeSeries'] = sorted(cell_data['timeSeries'], key=lambda x: x['period'])
            
            # Actualizar último valor
            if cell_data['timeSeries']:
                cell_data['lastValue'] = cell_data['timeSeries'][-1]['value']
                
                # Si hay PPTO y PREV, añadir métricas comparativas
                has_ppto = ppto_col in df.columns
                has_prev = prev_col in df.columns
                
                if has_ppto or has_prev:
                    last_point = cell_data['timeSeries'][-1]
                    
                    # Inicializar valores comparativos
                    cell_data['PPTO'] = last_point.get('ppto', 0) if has_ppto else 0
                    cell_data['PREV'] = last_point.get('prev', 0) if has_prev else 0
                    
                    # Calcular métricas comparativas
                    prvPtoPercent = 0
                    if has_ppto and has_prev and cell_data['PPTO'] != 0:
                        prvPtoPercent = (cell_data['PREV'] / cell_data['PPTO']) * 100
                    
                    realPrvPercent = 0
                    if has_prev and cell_data['PREV'] != 0:
                        realPrvPercent = (cell_data['lastValue'] / cell_data['PREV']) * 100
                    
                    pending = 0
                    if has_prev:
                        pending = cell_data['PREV'] - cell_data['lastValue']
                    
                    cell_data['comparative'] = {
                        'prvPtoPercent': prvPtoPercent,
                        'realPrvPercent': realPrvPercent,
                        'pending': pending
                    }
    
    def _process_kpi_cell_data(self, df, cat_col, subcat_col, project_col, company_col, prev_value_col, real_prev_percent_col, ppto_prev_percent_col, pending_value_col):
        """
        Procesa los datos de KPIs para preparar la estructura del cuadro de mando
        
        Args:
            df (pandas.DataFrame): DataFrame a procesar
            cat_col, subcat_col: Nombres de columnas obligatorias
            project_col, company_col, prev_value_col, real_prev_percent_col, ppto_prev_percent_col, pending_value_col: Nombres de columnas opcionales
        """
        # Agrupar por categoría y subcategoría
        group_cols = [cat_col, subcat_col]
        
        # Añadir proyecto y compañía si existen
        if project_col in df.columns:
            group_cols = [project_col] + group_cols
        
        if company_col in df.columns:
            group_cols = [company_col] + group_cols
        
        # Preparar diccionario para agrupar por las columnas de KPIs disponibles
        agg_dict = {}
        
        if prev_value_col in df.columns:
            agg_dict[prev_value_col] = 'sum'
        
        if real_prev_percent_col in df.columns:
            agg_dict[real_prev_percent_col] = 'mean'
        
        if ppto_prev_percent_col in df.columns:
            agg_dict[ppto_prev_percent_col] = 'mean'
        
        if pending_value_col in df.columns:
            agg_dict[pending_value_col] = 'sum'
        
        # Si no hay columnas de KPIs, no procesar
        if not agg_dict:
            return
        
        # Agrupar con todas las columnas
        grouped = df.groupby(group_cols).agg(agg_dict).reset_index()
        
        # Procesar cada fila agrupada
        for _, row in grouped.iterrows():
            # Crear clave de celda
            cell_components = []
            
            if project_col in df.columns:
                cell_components.append(str(row[project_col]))
            
            if company_col in df.columns:
                cell_components.append(str(row[company_col]))
            
            cell_components.append(str(row[cat_col]))
            cell_components.append(str(row[subcat_col]))
            
            cell_key = "|".join(cell_components)
            
            # Inicializar datos de celda si no existe
            if cell_key not in self.processed_data['cellData']:
                initial_data = {
                    'Category': row[cat_col],
                    'Subcategory': row[subcat_col]
                }
                
                if project_col in df.columns:
                    initial_data['PRJID'] = row[project_col]
                
                if company_col in df.columns:
                    initial_data['CIA'] = row[company_col]
                    
                self.processed_data['cellData'][cell_key] = initial_data
            
            # Añadir valores de KPIs si existen
            if prev_value_col in df.columns:
                self.processed_data['cellData'][cell_key]['PREV'] = float(row[prev_value_col]) if not pd.isna(row[prev_value_col]) else 0
            
            if real_prev_percent_col in df.columns:
                self.processed_data['cellData'][cell_key]['REALPREV'] = float(row[real_prev_percent_col]) if not pd.isna(row[real_prev_percent_col]) else 0
            
            if ppto_prev_percent_col in df.columns:
                self.processed_data['cellData'][cell_key]['PPTOPREV'] = float(row[ppto_prev_percent_col]) if not pd.isna(row[ppto_prev_percent_col]) else 0
            
            if pending_value_col in df.columns:
                self.processed_data['cellData'][cell_key]['PDTE'] = float(row[pending_value_col]) if not pd.isna(row[pending_value_col]) else 0
    
    def _calculate_totals(self):
        """Calcula los totales por categoría y subcategoría"""
        # Crear diccionarios para almacenar totales
        cat_totals = {}
        subcat_totals = {}
        
        # Calcular totales basados en el valor absoluto del último valor
        for _, cell_data in self.processed_data['cellData'].items():
            category = cell_data['Category']
            subcategory = cell_data['Subcategory']
            
            # Para datos históricos, usar el último valor
            if not self.is_kpi_data and 'lastValue' in cell_data:
                last_value = abs(cell_data['lastValue'])
                
                # Acumular por categoría
                if category not in cat_totals:
                    cat_totals[category] = 0
                cat_totals[category] += last_value
                
                # Acumular por subcategoría
                if subcategory not in subcat_totals:
                    subcat_totals[subcategory] = 0
                subcat_totals[subcategory] += last_value
            
            # Para datos de KPIs, usar el valor de previsión si está disponible
            elif self.is_kpi_data and 'PREV' in cell_data:
                prev_value = abs(cell_data['PREV'])
                
                # Acumular por categoría
                if category not in cat_totals:
                    cat_totals[category] = 0
                cat_totals[category] += prev_value
                
                # Acumular por subcategoría
                if subcategory not in subcat_totals:
                    subcat_totals[subcategory] = 0
                subcat_totals[subcategory] += prev_value
        
        # Añadir totales a los datos procesados
        self.processed_data['categoryTotals'] = cat_totals
        self.processed_data['subcategoryTotals'] = subcat_totals
    
    def get_processed_data(self):
        """
        Devuelve los datos procesados
        
        Returns:
            dict: Datos procesados
        """
        return self.processed_data
    
    def to_json(self, pretty=True):
        """
        Convierte los datos procesados a formato JSON
        
        Args:
            pretty (bool): Si se formatea el JSON para legibilidad
            
        Returns:
            str: Datos en formato JSON
        """
        if pretty:
            return json.dumps(self.processed_data, indent=2)
        else:
            return json.dumps(self.processed_data)


def _process_kpi_data(self):
    """
    Procesa los datos de KPIs
    
    Returns:
        dict: Datos de KPIs procesados
    """
    result = {'cellData': {}}
    
    # Usar los nombres exactos de las columnas KPI
    kpi_columns = {
        'prev_value': 'PREV',
        'real_prev_percent': '%REAL/PREV',
        'ppto_prev_percent': '%PPTO/PREV',  # Sin espacio al inicio
        'pending_value': 'PDTE'  # Sin punto final
    }
    
    # Columnas adicionales necesarias para identificar las celdas
    required_columns = ['Column', 'PrjId', 'Cia', 'Row']  # Corregido: 'Column' en lugar de 'cColumn'
    
    # Imprimir columnas disponibles para depuración
    print(f"Columnas disponibles en la hoja KPI: {list(self.data.columns)}")
    
    # Verificar que las columnas KPI existen
    missing_columns = []
    for key, col_name in kpi_columns.items():
        if col_name not in self.data.columns:
            missing_columns.append(f"{key} ({col_name})")
    
    if missing_columns:
        raise Exception(f"No se encontraron las siguientes columnas KPI: {', '.join(missing_columns)}")
    
    # Verificar que las columnas de identificación existen
    missing_required = [col for col in required_columns if col not in self.data.columns]
    if missing_required:
        raise Exception(f"No se encontraron las siguientes columnas requeridas: {', '.join(missing_required)}")
    
    # Procesar los datos con las columnas específicas
    for _, row in self.data.iterrows():
        # Obtener el identificador de la celda usando las columnas adicionales
        cell_id = self._get_cell_id(row)
        if not cell_id:
            continue
            
        # Extraer los valores KPI
        result['cellData'][cell_id] = {
            'prevValue': float(row[kpi_columns['prev_value']]) if pd.notna(row[kpi_columns['prev_value']]) else 0,
            'realPrevPercent': float(row[kpi_columns['real_prev_percent']]) if pd.notna(row[kpi_columns['real_prev_percent']]) else 0,
            'pptoPrevPercent': float(row[kpi_columns['ppto_prev_percent']]) if pd.notna(row[kpi_columns['ppto_prev_percent']]) else 0,
            'pendingValue': float(row[kpi_columns['pending_value']]) if pd.notna(row[kpi_columns['pending_value']]) else 0
        }
    
    return result

def _get_cell_id(self, row):
    """
    Obtiene el identificador único de una celda basado en sus datos
    
    Args:
        row: Fila de datos
        
    Returns:
        str: Identificador único de la celda
    """
    # Verificar si tenemos las columnas necesarias para identificar la celda
    if 'Column' in row and 'PrjId' in row and 'Cia' in row and 'Row' in row:  # Corregido: 'Column' en lugar de 'cColumn'
        # Crear un identificador único usando las columnas de identificación
        return f"{row['PrjId']}|{row['Cia']}|{row['Column']}|{row['Row']}"  # Corregido: 'Column' en lugar de 'cColumn'
    
    # Si no tenemos las columnas específicas, intentar con category y subcategory
    elif 'category' in row and 'subcategory' in row:
        return f"{row['category']}|{row['subcategory']}"
    
    # Si tenemos project, category y subcategory
    elif 'project' in row and 'category' in row and 'subcategory' in row:
        return f"{row['project']}|{row['category']}|{row['subcategory']}"
    
    return None


def _process_row(self, row):
    """
    Procesa una fila de datos y la convierte al formato requerido
    
    Args:
        row (dict): Fila de datos
        
    Returns:
        tuple: (clave jerárquica, datos procesados)
    """
    # Usar directamente los nombres de columnas originales sin mapeo
    processed = {}
    
    # Copiar todos los campos originales
    for key, value in row.items():
        processed[key] = value
    
    # Generar clave jerárquica usando los campos originales con nombres correctos en mayúsculas
    company = str(row.get('CIA', row.get('cia', '')))
    project = str(row.get('PRJID', row.get('prjid', '')))
    category = str(row.get('ROW', row.get('row', '')))
    subcategory = str(row.get('COLUMN', row.get('column', '')))
    
    # Asegurar que los valores son strings para la clave
    hierarchy_key = f"{company}|{project}|{category}|{subcategory}"
    
    # Añadir campos de categoría y subcategoría para facilitar el filtrado
    processed['Category'] = category
    processed['Subcategory'] = subcategory
    
    return hierarchy_key, processed

def _detect_columns(self):
    """
    Detecta las columnas disponibles en el DataFrame y crea un mapeo
    
    Returns:
        dict: Mapeo de columnas
    """
    # Verificar si el DataFrame tiene columnas
    if not hasattr(self.df, 'columns'):
        print("Advertencia: Los datos no tienen columnas definidas.")
        return {}
    
    # Crear un mapeo directo sin transformación de nombres
    column_mappings = {}
    
    # Buscar columnas clave con diferentes variantes de mayúsculas/minúsculas
    for col in self.df.columns:
        col_upper = col.upper() if isinstance(col, str) else col
        
        # Mapear columnas clave con sus variantes de mayúsculas/minúsculas
        if col_upper in ['CIA', 'PRJID', 'ROW', 'COLUMN', 'PREV', 'PPTO', 'PDTE', 'REALPREV', 'PPTOPREV']:
            column_mappings[col_upper] = col
    
    # Imprimir mapeo para depuración
    print(f"Mapeo de columnas detectado: {column_mappings}")
    
    return column_mappings

def _calculate_totals(self):
    """Calcula los totales por categoría y subcategoría"""
    # Crear diccionarios para almacenar totales
    cat_totals = {}
    subcat_totals = {}
    
    # Calcular totales basados en el valor absoluto del último valor
    for _, cell_data in self.processed_data['cellData'].items():
        category = cell_data['category']
        subcategory = cell_data['subcategory']
        
        # Para datos históricos, usar el último valor
        if not self.is_kpi_data and 'lastValue' in cell_data:
            last_value = abs(cell_data['lastValue'])
            
            # Acumular por categoría
            if category not in cat_totals:
                cat_totals[category] = 0
            cat_totals[category] += last_value
            
            # Acumular por subcategoría
            if subcategory not in subcat_totals:
                subcat_totals[subcategory] = 0
            subcat_totals[subcategory] += last_value
        
        # Para datos de KPIs, usar el valor de previsión si está disponible
        elif self.is_kpi_data and 'prevValue' in cell_data:
            prev_value = abs(cell_data['prevValue'])
            
            # Acumular por categoría
            if category not in cat_totals:
                cat_totals[category] = 0
            cat_totals[category] += prev_value
            
            # Acumular por subcategoría
            if subcategory not in subcat_totals:
                subcat_totals[subcategory] = 0
            subcat_totals[subcategory] += prev_value
    
        # Añadir totales a los datos procesados
        self.processed_data['categoryTotals'] = cat_totals
        self.processed_data['subcategoryTotals'] = subcat_totals
    
    def get_processed_data(self):
        """
        Devuelve los datos procesados
        
        Returns:
            dict: Datos procesados
        """
        return self.processed_data
    
    def to_json(self, pretty=True):
        """
        Convierte los datos procesados a formato JSON
        
        Args:
            pretty (bool): Si se formatea el JSON para legibilidad
            
        Returns:
            str: Datos en formato JSON
        """
        if pretty:
            return json.dumps(self.processed_data, indent=2)
        else:
            return json.dumps(self.processed_data)
