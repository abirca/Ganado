import os
from datetime import datetime
import openpyxl
from django.contrib import messages
from django.shortcuts import render, redirect, get_object_or_404, Http404
from django.db.models import Sum
from .forms import MovimientoForm, ProveedorForm, MovimientoClienteForm
from django.db import transaction
from openpyxl import load_workbook
from decimal import Decimal
import re
import json
from django.utils.safestring import mark_safe
from django.core.paginator import Paginator
from django.http import HttpResponse
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle

RUTA_EXCEL = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'Financiero.xlsx')
RUTA_EXCEL_SEGUNDO = r'D:\OneDrive - Olimpia IT S.A.S\Descargas\Financiero.xlsx'

# Configuración para tipos de entidad (Proveedores/Clientes)
ENTITY_CONFIG = {
    'proveedor': {
        'form': MovimientoForm,
        'movimiento_form_template': 'formulario.html',
        'resumen_template': 'resumen.html',
        'movimientos_template': 'movimientos.html',
        'agregar_persona_template': 'agregar_persona.html',
        'dashboard_template': 'dashboard.html',
        'editar_template': 'editar.html',
        'sheet_movimientos': 'Proveedores',
        'sheet_resumen': 'Resumen',
        'url_index': 'MovimientoProveedor',
        'url_resumen': 'resumen',
    },
    'cliente': {
        'form': MovimientoClienteForm,
        'movimiento_form_template': 'formularioCliente.html',
        'resumen_template': 'resumenCliente.html',
        'movimientos_template': 'movimientosCliente.html',
        'agregar_persona_template': 'agregar_personaCliente.html',
        'dashboard_template': 'dashboardCliente.html',
        'editar_template': 'editarCliente.html',
        'sheet_movimientos': 'ProveedoresCliente',
        'sheet_resumen': 'ResumenCliente',
        'url_index': 'MovimientoCliente',
        'url_resumen': 'resumenCliente',
    }
}

# Funciones genéricas reutilizables
def cargar_datos_excel(sheet_name):
    """Carga datos desde una hoja de Excel específica"""
    if not os.path.exists(RUTA_EXCEL):
        return []
    
    try:
        wb = openpyxl.load_workbook(RUTA_EXCEL)
        if sheet_name not in wb.sheetnames:
            return []
        
        ws = wb[sheet_name]
        return list(ws.iter_rows(min_row=2, values_only=True))
    except Exception as e:
        print(f"Error al cargar datos de Excel: {e}")
        return []

def obtener_ultimo_id(sheet_name):
    """Obtiene el último ID utilizado en una hoja de Excel"""
    datos = cargar_datos_excel(sheet_name)
    if not datos:
        return 0
    
    # Buscar el máximo ID en la primera columna
    max_id = 0
    for fila in datos:
        if fila and len(fila) > 0 and isinstance(fila[0], (int, float)):
            max_id = max(max_id, int(fila[0]))
    
    return max_id

def guardar_en_excel(sheet_name, datos, encabezados=None, modo='overwrite'):
    """Guarda datos en una hoja de Excel específica"""
    try:
        if os.path.exists(RUTA_EXCEL):
            wb = load_workbook(RUTA_EXCEL)
        else:
            wb = openpyxl.Workbook()
            # Eliminar la hoja por defecto si existe
            if 'Sheet' in wb.sheetnames:
                del wb['Sheet']
        
        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(sheet_name)
            if encabezados:
                ws.append(encabezados)
        else:
            ws = wb[sheet_name]
            # Si el modo es 'overwrite', limpiar la hoja
            if modo == 'overwrite':
                # Eliminar todas las filas excepto los encabezados
                if ws.max_row > 1:
                    ws.delete_rows(2, ws.max_row)
                # Si no hay encabezados, mantener los existentes
                if encabezados:
                    # Reemplazar encabezados existentes
                    for idx, encabezado in enumerate(encabezados, 1):
                        ws.cell(row=1, column=idx, value=encabezado)
        
        # Agregar datos
        for dato in datos:
            ws.append(dato)
        
        wb.save(RUTA_EXCEL)
        wb.save(RUTA_EXCEL_SEGUNDO)
        return True
    except Exception as e:
        print(f"Error al guardar en Excel: {e}")
        return False

def normalizar_total(total_raw):
    """Normaliza el valor total eliminando caracteres no numéricos"""
    try:
        total_str = re.sub(r'[^\d]', '', total_raw) if total_raw else '0'
        return Decimal(total_str) if total_str else Decimal('0')
    except Exception as e:
        print(f"Error al normalizar total: {e}")
        return Decimal('0')

def obtener_movimientos_filtrados(entity_type, proveedor_filtrado=None, fecha_filtrada=None):
    """Obtiene movimientos filtrados por proveedor y fecha"""
    config = ENTITY_CONFIG[entity_type]
    movimientos = []
    
    datos = cargar_datos_excel(config['sheet_movimientos'])
    
    for row in datos:
        if len(row) < 6:
            continue
            
        mov = {
            'id': row[0],
            'fecha': row[1],
            'proveedor': row[2],
            'detalle': row[3],
            'obs': row[4],
            'total': row[5]
        }
        
        # Aplicar filtros
        cumple_proveedor = True
        cumple_fecha = True
        
        if proveedor_filtrado:
            prov = mov['proveedor'] or ''
            cumple_proveedor = (str(prov).strip().lower() == proveedor_filtrado.strip().lower())
        
        if fecha_filtrada:
            fecha_str = ''
            if isinstance(mov['fecha'], datetime):
                fecha_str = mov['fecha'].strftime('%Y-%m-%d')
            elif isinstance(mov['fecha'], str):
                fecha_str = mov['fecha']
            cumple_fecha = (fecha_str == fecha_filtrada)
        
        if cumple_proveedor and cumple_fecha:
            movimientos.append(mov)
    
    return movimientos

def obtener_resumen_filtrado(entity_type, proveedor_filtrado=None):
    """Obtiene resumen filtrado por proveedor"""
    config = ENTITY_CONFIG[entity_type]
    datos = []
    
    resumen_data = cargar_datos_excel(config['sheet_resumen'])
    
    for row in resumen_data:
        if len(row) < 5:
            continue
        
        id = row[0]
        proveedor = row[1]
        factura = row[2] or 0
        ahorro = row[3] or 0
        saldo = row[4] or 0
        
        if proveedor_filtrado:
            if proveedor == proveedor_filtrado:
                datos.append((id, proveedor, factura, ahorro, saldo))
        else:
            datos.append((id, proveedor, factura, ahorro, saldo))
    
    return datos

def recalcular_resumen(entity_type):
    """Recalcula el resumen a partir de los movimientos"""
    config = ENTITY_CONFIG[entity_type]
    
    # Obtener todos los movimientos
    movimientos = cargar_datos_excel(config['sheet_movimientos'])
    
    # Diccionario para almacenar los totales por proveedor
    resumen_dict = {}
    
    for mov in movimientos:
        if len(mov) < 6:
            continue
            
        id, fecha, proveedor, detalle, obs, total = mov
        
        if proveedor not in resumen_dict:
            resumen_dict[proveedor] = {
                'facturas': Decimal('0'),
                'ahorros': Decimal('0'),
                'saldo': Decimal('0')
            }
        
        if detalle == 'Factura':
            resumen_dict[proveedor]['facturas'] += Decimal(str(total))
        elif detalle == 'Ahorro':
            resumen_dict[proveedor]['ahorros'] += Decimal(str(total))
        
        resumen_dict[proveedor]['saldo'] = resumen_dict[proveedor]['facturas'] - resumen_dict[proveedor]['ahorros']
    
    # Preparar datos para guardar
    resumen_data = []
    for proveedor, valores in resumen_dict.items():
        resumen_data.append([
            obtener_ultimo_id(config['sheet_resumen']) + 1,  # Nuevo ID
            proveedor,
            float(valores['facturas']),
            float(valores['ahorros']),
            float(valores['saldo'])
        ])
    
    # Guardar en Excel
    encabezados = ['Id', 'Proveedor', 'Total Facturas', 'Total Ahorro', 'Saldo']
    guardar_en_excel(config['sheet_resumen'], resumen_data, encabezados, modo='overwrite')
    
    return True

# Vistas genéricas
def movimiento_view(request, entity_type):
    """Vista genérica para movimientos de proveedores o clientes"""
    config = ENTITY_CONFIG[entity_type]
    form = config['form']()
    
    proveedor_filtrado = request.GET.get('proveedor', None)
    fecha_filtrada = request.GET.get('fecha', None)
    
    movimientos = obtener_movimientos_filtrados(entity_type, proveedor_filtrado, fecha_filtrada)
    resumen = cargar_datos_excel(config['sheet_resumen'])
    resumen_filtrado = obtener_resumen_filtrado(entity_type, proveedor_filtrado)
    
    paginator = Paginator(movimientos, 10)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    
    return render(request, config['movimiento_form_template'], {
        'form': form,
        'resumen': resumen,
        'movimientos': page_obj,
        'resumen_Filtrado': resumen_filtrado,
        'proveedor_filtrado': proveedor_filtrado,
        'fecha_filtrada': fecha_filtrada,
        'paginator': paginator,
        'page_obj': page_obj,
    })

def resumen_view(request, entity_type):
    """Vista genérica para resumen de proveedores o clientes"""
    config = ENTITY_CONFIG[entity_type]
    
    proveedor_filtrado = request.GET.get('proveedor', None)
    datos = obtener_resumen_filtrado(entity_type, proveedor_filtrado)
    
    return render(request, config['resumen_template'], {
        'resumen': datos,
        'proveedor_filtrado': proveedor_filtrado,
    })

def movimientos_list_view(request, entity_type):
    """Vista genérica para listar movimientos de proveedores o clientes con filtros mejorados"""
    config = ENTITY_CONFIG[entity_type]
    
    # Obtener todos los parámetros de filtro
    proveedor_filtrado = request.GET.get('proveedor', '').strip()
    fecha_filtrada = request.GET.get('fecha', '').strip()
    fecha_inicio = request.GET.get('fecha_inicio', '').strip()
    fecha_fin = request.GET.get('fecha_fin', '').strip()
    movimientos = []
    
    datos = cargar_datos_excel(config['sheet_movimientos'])
    resumen = cargar_datos_excel(config['sheet_resumen'])
    
    for mov_data in datos:
        if len(mov_data) < 6:
            continue
            
        mov_id, fecha_raw, proveedor, detalle, obs, total = mov_data
        
        # Normalizar fecha
        fecha_mov = None
        if isinstance(fecha_raw, datetime):
            fecha_mov = fecha_raw
        elif isinstance(fecha_raw, str):
            try:
                fecha_mov = datetime.strptime(fecha_raw, "%Y-%m-%d")
            except:
                try:
                    fecha_mov = datetime.strptime(fecha_raw, "%d/%m/%Y")
                except:
                    fecha_mov = None
        
        # Filtrar por proveedor
        if proveedor_filtrado and (not proveedor or str(proveedor).strip().lower() != proveedor_filtrado.lower()):
            continue
        
        # Si hay fecha específica, ignorar los filtros de rango
        if fecha_filtrada:
            try:
                filtro_fecha = datetime.strptime(fecha_filtrada, "%Y-%m-%d").date()
                if not fecha_mov or fecha_mov.date() != filtro_fecha:
                    continue
            except:
                # Si hay error en el formato, intentar con otro formato
                try:
                    filtro_fecha = datetime.strptime(fecha_filtrada, "%d/%m/%Y").date()
                    if not fecha_mov or fecha_mov.date() != filtro_fecha:
                        continue
                except:
                    continue
        else:
            # Filtrar por rango de fechas (fecha_inicio y fecha_fin)
            if fecha_inicio and fecha_mov:
                try:
                    fecha_inicio_dt = datetime.strptime(fecha_inicio, "%Y-%m-%d").date()
                    if fecha_mov.date() < fecha_inicio_dt:
                        continue
                except:
                    # Si hay error en el formato, intentar con otro formato
                    try:
                        fecha_inicio_dt = datetime.strptime(fecha_inicio, "%d/%m/%Y").date()
                        if fecha_mov.date() < fecha_inicio_dt:
                            continue
                    except:
                        continue
            
            if fecha_fin and fecha_mov:
                try:
                    fecha_fin_dt = datetime.strptime(fecha_fin, "%Y-%m-%d").date()
                    if fecha_mov.date() > fecha_fin_dt:
                        continue
                except:
                    # Si hay error en el formato, intentar con otro formato
                    try:
                        fecha_fin_dt = datetime.strptime(fecha_fin, "%d/%m/%Y").date()
                        if fecha_mov.date() > fecha_fin_dt:
                            continue
                    except:
                        continue
        
        # Si llegamos aquí, el movimiento pasa todos los filtros
        movimientos.append({
            'id': mov_id,
            'fecha': fecha_mov.date() if fecha_mov else None,
            'proveedor': proveedor,
            'detalle': detalle,
            'obs': obs,
            'total': total,
        })
    
    # Ordenar movimientos por fecha (más reciente primero)
    movimientos.sort(key=lambda x: x['fecha'] if x['fecha'] else datetime.min.date(), reverse=True)
    
    # Paginación
    paginator = Paginator(movimientos, 10)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    
    return render(request, config['movimientos_template'], {
        'resumen': resumen,
        'movimientos': page_obj,
        'proveedor_filtrado': proveedor_filtrado,
        'fecha_filtrada': fecha_filtrada,
        'fecha_inicio': fecha_inicio,
        'fecha_fin': fecha_fin,
        'paginator': paginator,
        'page_obj': page_obj,
        'all_params': request.GET.urlencode()  # Para mantener todos los parámetros en los enlaces de paginación
    })

def agregar_persona_view(request, entity_type):
    """Vista genérica para agregar personas (proveedores o clientes)"""
    config = ENTITY_CONFIG[entity_type]
    
    if request.method == 'POST':
        form = ProveedorForm(request.POST)
        if form.is_valid():
            nombre = form.cleaned_data['nombre'].strip()
            
            # Validar si ya existe en Excel
            resumen_data = cargar_datos_excel(config['sheet_resumen'])
            
            existe = False
            for row in resumen_data:
                if len(row) > 1 and row[1] == nombre:
                    existe = True
                    break
            
            if existe:
                messages.info(request, "La persona ya existe en el archivo Excel. ❌")
                form.add_error('nombre', 'La persona ya existe en el archivo Excel.')
            else:
                # Obtener el último ID y generar uno nuevo
                nuevo_id = obtener_ultimo_id(config['sheet_resumen']) + 1
                
                # Crear nueva fila para el resumen
                nueva_fila = [nuevo_id, nombre, 0, 0, 0]
                
                # Agregar a los datos existentes
                resumen_data.append(nueva_fila)
                
                # Guardar en Excel
                encabezados = ['Id', 'Proveedor', 'Total Facturas', 'Total Ahorro', 'Saldo']
                if guardar_en_excel(config['sheet_resumen'], resumen_data, encabezados, modo='overwrite'):
                    messages.success(request, f'Persona {nombre} agregada correctamente. ✔️')
                    return redirect(config['url_resumen'])
                else:
                    messages.error(request, 'Error al guardar la persona. Inténtelo de nuevo. ❌')
                
                return redirect(config['url_resumen'])
    else:
        form = ProveedorForm()
    
    return render(request, config['agregar_persona_template'], {'form': form})

def dashboard_view(request, entity_type):
    """Vista genérica para dashboard de proveedores o clientes"""
    config = ENTITY_CONFIG[entity_type]
    
    proveedor_filtrado = request.GET.get('proveedor', None)
    proveedores = []
    facturas_totales = []
    ahorros_totales = []
    saldos_totales = []
    
    facturas_por_mes = {}
    ahorros_por_mes = {}
    
    # Cargar datos de resumen
    resumen_data = cargar_datos_excel(config['sheet_resumen'])
    for row in resumen_data:
        if len(row) < 5:
            continue
            
        id, proveedor, factura, ahorro, saldo = row[0], row[1], row[2] or 0, row[3] or 0, row[4] or 0
        
        proveedores.append(proveedor)
        facturas_totales.append(factura)
        ahorros_totales.append(ahorro)
        saldos_totales.append(saldo)
    
    # Cargar datos de movimientos
    movimientos_data = cargar_datos_excel(config['sheet_movimientos'])
    for row in movimientos_data:
        if len(row) < 6:
            continue
            
        id_, fecha, prov, detalle, obs, total = row
        
        if proveedor_filtrado and prov != proveedor_filtrado:
            continue
        
        if not fecha:
            continue
        
        # Convertir fecha a mes-año
        if isinstance(fecha, datetime):
            mes_ano = fecha.strftime('%Y-%m')
        else:
            try:
                dt = datetime.strptime(str(fecha), '%Y-%m-%d')
                mes_ano = dt.strftime('%Y-%m')
            except:
                try:
                    dt = datetime.strptime(str(fecha), '%d/%m/%Y')
                    mes_ano = dt.strftime('%Y-%m')
                except:
                    continue
        
        # Facturas (detalle == 'Factura')
        if detalle == 'Factura':
            facturas_por_mes.setdefault(prov, {})
            facturas_por_mes[prov][mes_ano] = facturas_por_mes[prov].get(mes_ano, 0) + (total or 0)
        
        # Ahorros (detalle == 'Ahorro')
        if detalle == 'Ahorro':
            ahorros_por_mes.setdefault(prov, {})
            ahorros_por_mes[prov][mes_ano] = ahorros_por_mes[prov].get(mes_ano, 0) + (total or 0)
    
    # Meses únicos combinando facturas y ahorros
    meses = set()
    for prov_data in facturas_por_mes.values():
        meses.update(prov_data.keys())
    for prov_data in ahorros_por_mes.values():
        meses.update(prov_data.keys())
    meses = sorted(meses)
    
    # Formatear datos para Chart.js
    facturas_linea = []
    ahorros_linea = []
    proveedores_filtrados = sorted(set(list(facturas_por_mes.keys()) + list(ahorros_por_mes.keys())))
    
    for prov in proveedores_filtrados:
        # Facturas
        data_fact = [facturas_por_mes.get(prov, {}).get(mes, 0) for mes in meses]
        facturas_linea.append({'proveedor': prov, 'datos': data_fact})
        # Ahorros
        data_ahorro = [ahorros_por_mes.get(prov, {}).get(mes, 0) for mes in meses]
        ahorros_linea.append({'proveedor': prov, 'datos': data_ahorro})
    
    context = {
        'proveedores': proveedores,
        'facturas': mark_safe(json.dumps(facturas_totales)),
        'ahorros': mark_safe(json.dumps(ahorros_totales)),
        'saldos': mark_safe(json.dumps(saldos_totales)),
        'proveedores_filtrados': proveedores_filtrados,
        'meses': mark_safe(json.dumps(meses)),
        'facturas_linea': mark_safe(json.dumps(facturas_linea)),
        'ahorros_linea': mark_safe(json.dumps(ahorros_linea)),
        'proveedor_filtrado': proveedor_filtrado or '',
    }
    return render(request, config['dashboard_template'], context)

# En tus funciones de vista, agrega mensajes según sea necesario
def guardar_movimiento_view(request, entity_type):
    """Vista genérica para guardar movimientos de proveedores o clientes"""
    config = ENTITY_CONFIG[entity_type]
    
    if request.method == 'POST':
        form = config['form'](request.POST)
        if form.is_valid():
            data = form.cleaned_data
            proveedor = data['proveedor']
            detalle = data['detalle']
            obs = data['obs']
            fecha = data['fecha']
            
            # Normalizar total
            total = normalizar_total(request.POST.get('total', ''))
            
            # Obtener el último ID y generar uno nuevo
            nuevo_id = obtener_ultimo_id(config['sheet_movimientos']) + 1
            
            # Crear nueva fila para el movimiento
            nueva_fila = [nuevo_id, fecha, proveedor, detalle, obs, float(total)]
            
            # Cargar datos existentes y agregar la nueva fila
            movimientos_data = cargar_datos_excel(config['sheet_movimientos'])
            movimientos_data.append(nueva_fila)
            
            # Guardar en Excel
            encabezados = ['Id', 'Fecha', 'Proveedor', 'Detalle', 'Obs', 'Total']
            if guardar_en_excel(config['sheet_movimientos'], movimientos_data, encabezados, modo='overwrite'):
                # Recalcular el resumen
                recalcular_resumen(entity_type)
                
                # Mensaje de éxito
                messages.success(request, f'Movimiento de {detalle} para {proveedor} guardado correctamente.')
                return redirect(config['url_index'])
            else:
                # Mensaje de error
                messages.error(request, 'Error al guardar el movimiento. Inténtelo de nuevo.')
    
    else:
        form = config['form']()
        messages.info(request, 'Por favor, complete el formulario para agregar un nuevo movimiento.')
    
    return render(request, config['movimiento_form_template'], {'form': form})

def editar_movimiento_view(request, entity_type, index):
    """Vista genérica para editar movimientos de proveedores o clientes"""
    config = ENTITY_CONFIG[entity_type]
    
    # Buscar el movimiento en Excel
    movimientos_data = cargar_datos_excel(config['sheet_movimientos'])
    mov = None
    
    for row in movimientos_data:
        if row and len(row) > 0 and row[0] == index:
            mov = {
                'id': row[0],
                'fecha': row[1],
                'proveedor': row[2],
                'detalle': row[3],
                'obs': row[4],
                'total': row[5]
            }
            break
    
    if not mov:
        messages.error(request, 'Movimiento no encontrado.')
        raise Http404("Movimiento no encontrado")
    
    if request.method == 'POST':
        form = config['form'](request.POST)
        if form.is_valid():
            data = form.cleaned_data
            
            # Actualizar el movimiento en la lista
            for i, row in enumerate(movimientos_data):
                if row and len(row) > 0 and row[0] == index:
                    movimientos_data[i] = [
                        index,
                        data['fecha'],
                        data['proveedor'],
                        data['detalle'],
                        data['obs'],
                        float(data['total'])
                    ]
                    break
            
            # Guardar en Excel
            encabezados = ['Id', 'Fecha', 'Proveedor', 'Detalle', 'Obs', 'Total']
            if guardar_en_excel(config['sheet_movimientos'], movimientos_data, encabezados, modo='overwrite'):
            
                # Recalcular el resumen
                recalcular_resumen(entity_type)
                # Mensaje de éxito
                messages.success(request, f'Se movimiento Modifico el registro correctamente.')

                return redirect(config['url_index'])
            else:
                 # Mensaje de error
                messages.error(request, 'Error al guardar el movimiento. Inténtelo de nuevo.')
    
    else:
        # Convertir fecha a formato string para el formulario
        if isinstance(mov['fecha'], datetime):
            fecha_str = mov['fecha'].strftime('%Y-%m-%d')
        else:
            fecha_str = mov['fecha']
            
        initial_data = {
            'fecha': fecha_str,
            'proveedor': mov['proveedor'],
            'detalle': mov['detalle'],
            'obs': mov['obs'],
            'total': mov['total'],
        }
        form = config['form'](initial=initial_data)

        
    
    return render(request, config['editar_template'], {'form': form, 'id': index})

def descargar_excel_entidad(request, entity_type):
    """Vista para descargar un archivo Excel con los movimientos de una entidad específica con filtros"""
    try:
        config = ENTITY_CONFIG[entity_type]
        
        # Obtener parámetros de filtro
        nombre_entidad = request.GET.get('proveedor', '')
        fecha_inicio = request.GET.get('fecha_inicio', '')
        fecha_fin = request.GET.get('fecha_fin', '')
        fecha_especifica = request.GET.get('fecha', '')
        
        # Obtener todos los movimientos y luego filtrar
        movimientos_data = cargar_datos_excel(config['sheet_movimientos'])
        movimientos_filtrados = []
        
        for row in movimientos_data:
            if len(row) < 6:
                continue
                
            mov = {
                'id': row[0],
                'fecha': row[1],
                'proveedor': row[2],
                'detalle': row[3],
                'obs': row[4],
                'total': row[5]
            }
            
            # Aplicar filtros
            cumple_proveedor = True
            cumple_fecha_especifica = True
            cumple_fecha_inicio = True
            cumple_fecha_fin = True
            
            # Filtro por proveedor
            if nombre_entidad:
                prov = mov['proveedor'] or ''
                cumple_proveedor = (str(prov).strip().lower() == nombre_entidad.strip().lower())
            
            # Filtro por fecha específica
            if fecha_especifica:
                try:
                    fecha_mov = None
                    if isinstance(mov['fecha'], datetime):
                        fecha_mov = mov['fecha']
                    elif isinstance(mov['fecha'], str):
                        try:
                            fecha_mov = datetime.strptime(mov['fecha'], "%Y-%m-%d")
                        except:
                            try:
                                fecha_mov = datetime.strptime(mov['fecha'], "%d/%m/%Y")
                            except:
                                pass
                    
                    if fecha_mov:
                        fecha_especifica_dt = datetime.strptime(fecha_especifica, "%Y-%m-%d")
                        cumple_fecha_especifica = (fecha_mov.date() == fecha_especifica_dt.date())
                    else:
                        cumple_fecha_especifica = False
                except:
                    cumple_fecha_especifica = False
            else:
                # Filtro por fecha inicio
                if fecha_inicio:
                    try:
                        fecha_mov = None
                        if isinstance(mov['fecha'], datetime):
                            fecha_mov = mov['fecha']
                        elif isinstance(mov['fecha'], str):
                            try:
                                fecha_mov = datetime.strptime(mov['fecha'], "%Y-%m-%d")
                            except:
                                try:
                                    fecha_mov = datetime.strptime(mov['fecha'], "%d/%m/%Y")
                                except:
                                    pass
                        
                        if fecha_mov:
                            fecha_inicio_dt = datetime.strptime(fecha_inicio, "%Y-%m-%d")
                            cumple_fecha_inicio = (fecha_mov.date() >= fecha_inicio_dt.date())
                        else:
                            cumple_fecha_inicio = False
                    except:
                        cumple_fecha_inicio = False
                
                # Filtro por fecha fin
                if fecha_fin:
                    try:
                        fecha_mov = None
                        if isinstance(mov['fecha'], datetime):
                            fecha_mov = mov['fecha']
                        elif isinstance(mov['fecha'], str):
                            try:
                                fecha_mov = datetime.strptime(mov['fecha'], "%Y-%m-%d")
                            except:
                                try:
                                    fecha_mov = datetime.strptime(mov['fecha'], "%d/%m/%Y")
                                except:
                                    pass
                        
                        if fecha_mov:
                            fecha_fin_dt = datetime.strptime(fecha_fin, "%Y-%m-%d")
                            cumple_fecha_fin = (fecha_mov.date() <= fecha_fin_dt.date())
                        else:
                            cumple_fecha_fin = False
                    except:
                        cumple_fecha_fin = False
            
            # Determinar qué filtros aplicar
            if fecha_especifica:
                # Solo usar filtro de fecha específica
                if cumple_proveedor and cumple_fecha_especifica:
                    movimientos_filtrados.append(mov)
            else:
                # Usar filtros de rango de fechas
                if cumple_proveedor and cumple_fecha_inicio and cumple_fecha_fin:
                    movimientos_filtrados.append(mov)
        
        # Crear un nuevo libro de Excel
        wb = openpyxl.Workbook()
        ws = wb.active
        
        # Crear un estilo para formato de moneda
        moneda_style = NamedStyle(name="moneda_style")
        moneda_style.number_format = '"$"#,##0.00'
        
        # Si el estilo ya existe, eliminarlo para evitar conflictos
        if "moneda_style" in wb.named_styles:
            del wb.named_styles[wb.named_styles.index("moneda_style")]
        
        # Agregar el estilo al libro
        wb.add_named_style(moneda_style)
        
        # Título basado en los filtros aplicados
        titulo = f"Movimientos de {entity_type.capitalize()}"
        if nombre_entidad:
            titulo += f" - {nombre_entidad}"
        if fecha_especifica:
            titulo += f" - Fecha: {fecha_especifica}"
        elif fecha_inicio or fecha_fin:
            titulo += " - Período: "
            if fecha_inicio and fecha_fin:
                titulo += f"{fecha_inicio} al {fecha_fin}"
            elif fecha_inicio:
                titulo += f"{fecha_inicio} en adelante"
            elif fecha_fin:
                titulo += f"hasta {fecha_fin}"
        
        ws.title = titulo[:31]  # Limitar a 31 caracteres (máximo de Excel)
        
        # Agregar título
        ws.cell(row=1, column=1, value=titulo)
        ws.cell(row=1, column=1).font = Font(bold=True, size=14)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
        
        # Definir encabezados
        encabezados = ['Id', 'Fecha', 'Proveedor/Cliente', 'Detalle Factura', 'Detalle Abono', 'Observaciones', 'Total']
        for col_num, encabezado in enumerate(encabezados, 1):
            celda = ws.cell(row=3, column=col_num, value=encabezado)
            celda.font = Font(bold=True)
            celda.alignment = Alignment(horizontal='center')
        
        # Llenar datos
        fila = 4
        total_facturas = 0
        total_abonos = 0
        
        for mov in movimientos_filtrados:
            # ID
            ws.cell(row=fila, column=1, value=mov['id'])
            
            # Fecha
            fecha = mov['fecha']
            if isinstance(fecha, datetime):
                fecha = fecha.strftime('%d/%m/%Y')
            elif isinstance(fecha, str):
                # Intentar convertir string a fecha y luego formatear
                try:
                    fecha_dt = datetime.strptime(fecha, "%Y-%m-%d")
                    fecha = fecha_dt.strftime('%d/%m/%Y')
                except:
                    try:
                        fecha_dt = datetime.strptime(fecha, "%d/%m/%Y")
                        fecha = fecha_dt.strftime('%d/%m/%Y')
                    except:
                        fecha = "Fecha inválida"
            ws.cell(row=fila, column=2, value=fecha)
            
            # Proveedor/Cliente
            ws.cell(row=fila, column=3, value=mov['proveedor'])
            
            # Detalle Factura/Abono
            if mov['detalle'] == 'Factura':
                ws.cell(row=fila, column=4, value=mov['detalle'])
                ws.cell(row=fila, column=5, value='')
                total_facturas += float(mov['total'] or 0)
            else:  # Ahorro/Abono
                ws.cell(row=fila, column=4, value='')
                ws.cell(row=fila, column=5, value=mov['detalle'])
                total_abonos += float(mov['total'] or 0)
            
            # Observaciones
            ws.cell(row=fila, column=6, value=mov['obs'] or '')
            
            # Total - aplicar formato de moneda
            total_valor = float(mov['total'] or 0)
            celda_total = ws.cell(row=fila, column=7, value=total_valor)
            celda_total.style = moneda_style
            
            fila += 1
        
        # Agregar fila de totales
        fila_total = fila + 1
        
        # Total Facturas
        ws.cell(row=fila_total, column=4, value="TOTAL FACTURAS:")
        ws.cell(row=fila_total, column=4).font = Font(bold=True)
        celda_total_facturas = ws.cell(row=fila_total, column=7, value=total_facturas)
        celda_total_facturas.font = Font(bold=True)
        celda_total_facturas.style = moneda_style
        
        # Total Abonos
        ws.cell(row=fila_total+1, column=5, value="TOTAL ABONOS:")
        ws.cell(row=fila_total+1, column=5).font = Font(bold=True)
        celda_total_abonos = ws.cell(row=fila_total+1, column=7, value=total_abonos)
        celda_total_abonos.font = Font(bold=True)
        celda_total_abonos.style = moneda_style
        
        # Saldo (Diferencia)
        saldo = total_abonos - total_facturas
        ws.cell(row=fila_total+2, column=6, value="SALDO:")
        ws.cell(row=fila_total+2, column=6).font = Font(bold=True)
        celda_saldo = ws.cell(row=fila_total+2, column=7, value=saldo)
        celda_saldo.font = Font(bold=True, color="FF0000" if saldo < 0 else "007500")
        celda_saldo.style = moneda_style
        
        # Ajustar el ancho de las columnas
        column_widths = [8, 12, 20, 15, 15, 30, 15]  # Anchos personalizados para cada columna
        for column, width in enumerate(column_widths, 1):
            col_letter = get_column_letter(column)
            ws.column_dimensions[col_letter].width = width
        
        # Preparar la respuesta HTTP
        response = HttpResponse(
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        filename = f"movimientos_{entity_type}_{nombre_entidad or 'todos'}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        response['Content-Disposition'] = f'attachment; filename={filename}'
        
        message.success(request,f"Archivo Excel generado correctamente: {filename}")
        # Guardar el libro en la respuesta
        wb.save(response)
        
        return response
        
    except Exception as e:
        # En caso de error, retornar una respuesta de error
        import traceback
        message.error(request,f"Error al generar el archivo Excel")
        error_msg = f"Error al generar el archivo Excel: {str(e)}\n{traceback.format_exc()}"
        return HttpResponse(error_msg, status=500)

def gestionar_persona_view(request, entity_type, id=None):
    """Vista genérica para agregar o editar personas (proveedores o clientes) buscando en Excel"""
    config = ENTITY_CONFIG[entity_type]
    
    # Si se proporciona un ID, estamos editando una persona existente
    if id:
        # Buscar la persona en el archivo Excel
        persona = None
        es_edicion = True
        
        # Cargar datos de Excel
        resumen_data = cargar_datos_excel(config['sheet_resumen'])
        
        # Buscar la persona por ID
        for row in resumen_data:
            if len(row) > 0 and row[0] == id:
                persona = {
                    'id': row[0],
                    'proveedor': row[1],
                    'facturas': row[2] if len(row) > 2 else 0,
                    'ahorros': row[3] if len(row) > 3 else 0,
                    'saldo': row[4] if len(row) > 4 else 0
                }
                break
                
        if not persona:
            # Si no se encuentra en Excel, mostrar error 404
            messages.error(request, "La persona no existe en el archivo Excel.")
            raise Http404("La persona no existe")
    else:
        persona = None
        es_edicion = False
    
    if request.method == 'POST':
        form = ProveedorForm(request.POST)
        if form.is_valid():
            nombre = form.cleaned_data['nombre'].strip()
            
            # Validar si ya existe en Excel (excepto si estamos editando la misma persona)
            resumen_data = cargar_datos_excel(config['sheet_resumen'])
            existe_excel = False
            
            for row in resumen_data:
                if len(row) > 1 and row[1] == nombre:  # Nombre está en la segunda columna
                    if es_edicion and row[0] == id:
                        # Es la misma persona que estamos editando, no es un duplicado
                        continue
                    existe_excel = True
                    break
            
            if existe_excel:
                messages.info(request, "La persona ya existe en el archivo Excel. ❌")
                form.add_error('nombre', 'La persona ya existe en el archivo Excel.')
            else:
                # Actualizar el archivo Excel
                resumen_data = cargar_datos_excel(config['sheet_resumen'])
                
                if es_edicion:
                    # Actualizar la fila existente en los datos de Excel
                    for i, row in enumerate(resumen_data):
                        if len(row) > 0 and row[0] == id:
                            resumen_data[i] = [
                                id,
                                nombre,
                                persona['facturas'],
                                persona['ahorros'],
                                persona['saldo']
                            ]
                            break
                else:
                    # Crear nueva persona
                    nuevo_id = obtener_ultimo_id(config['sheet_resumen']) + 1
                    nueva_fila = [nuevo_id, nombre, 0, 0, 0]
                    resumen_data.append(nueva_fila)
                
                # Guardar los datos actualizados en Excel
                encabezados = ['Id', 'Proveedor', 'Total Facturas', 'Total Ahorro', 'Saldo']
                if  guardar_en_excel(config['sheet_resumen'], resumen_data, encabezados, modo='overwrite'):
                    messages.success(request, f'Persona {nombre} {"actualizada" if es_edicion else "agregada"} correctamente. ✔️')
                    return redirect(config['url_resumen'])
                else:
                    messages.error(request, 'Error al guardar la persona. Inténtelo de nuevo. ❌')
                return redirect(config['url_resumen'])
    else:
        # Si es edición, precargar el formulario con los datos existentes
        if es_edicion:
            form = ProveedorForm(initial={'nombre': persona['proveedor']})
        else:
            form = ProveedorForm()
    
    return render(request, config['agregar_persona_template'], {
        'form': form,
        'es_edicion': es_edicion,
        'persona': persona
    })

# Vistas específicas para proveedores y clientes
def descargar_excel_proveedor(request):
    return descargar_excel_entidad(request, 'proveedor')

def descargar_excel_cliente(request):
    return descargar_excel_entidad(request, 'cliente')

# Vistas específicas (ahora son simples wrappers de las vistas genéricas)
def MovimientoProveedor(request):
    return movimiento_view(request, 'proveedor')

def MovimientoCliente(request):
    return movimiento_view(request, 'cliente')

def resumen(request):
    return resumen_view(request, 'proveedor')

def resumenCliente(request):
    return resumen_view(request, 'cliente')

def movimientos(request):
    return movimientos_list_view(request, 'proveedor')

def movimientosCliente(request):
    return movimientos_list_view(request, 'cliente')

def agregar_persona(request):
    return agregar_persona_view(request, 'proveedor')

def agregar_persona_Cliente(request):
    return agregar_persona_view(request, 'cliente')

def index(request):
    return dashboard_view(request, 'proveedor')

def dashboardCliente(request):
    return dashboard_view(request, 'cliente')

def guardar_movimiento(request):
    return guardar_movimiento_view(request, 'proveedor')

def guardar_movimiento_cliente(request):
    return guardar_movimiento_view(request, 'cliente')

def editar_movimiento(request, index):
    return editar_movimiento_view(request, 'proveedor', index)

def editar_movimiento_Cliente(request, index):
    return editar_movimiento_view(request, 'cliente', index)

def editar_proveedor(request, id):
    return gestionar_persona_view(request, 'proveedor', id)

def editar_cliente(request, id):
    return gestionar_persona_view(request, 'cliente', id)