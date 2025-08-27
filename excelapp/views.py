import os
from datetime import datetime
import openpyxl
from django.shortcuts import render, redirect
from django.db.models import Sum
from .forms import MovimientoForm, ProveedorForm, MovimientoClienteForm
from .models import Movimiento, Resumen, Movimiento_Cliente, Resumen_Cliente
from django.db import transaction
from openpyxl import load_workbook
from decimal import Decimal
import re
import json
from django.utils.safestring import mark_safe
from django.core.paginator import Paginator
from django.shortcuts import get_object_or_404

RUTA_EXCEL = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'Financiero.xlsx')

#Proveedor

def MovimientoProveedor(request):
    form = MovimientoForm()
    resumen = []
    resumen_Filtrado = []
    movimientos = []

    proveedor_filtrado = request.GET.get('proveedor', None)
    fecha_filtrada = request.GET.get('fecha', None)

    if os.path.exists(RUTA_EXCEL):
        wb = openpyxl.load_workbook(RUTA_EXCEL)

        # Cargar resumen (igual)
        if 'Resumen' in wb.sheetnames:
            ws_resumen = wb['Resumen']
            for row in ws_resumen.iter_rows(min_row=2, values_only=True):
                resumen.append(row)

        # Filtrar resumen
        if 'Resumen' in wb.sheetnames:
            ws_resumen_Filtrado = wb['Resumen']
            for row in ws_resumen_Filtrado.iter_rows(min_row=2, values_only=True):
                if proveedor_filtrado:
                    if row[0] and row[0].strip().lower() == proveedor_filtrado.strip().lower():
                        resumen_Filtrado.append(row)
                else:
                    resumen_Filtrado.append(row)

        # Cargar movimientos con ID
        if 'Proveedores' in wb.sheetnames:
            ws_mov = wb['Proveedores']
            all_movs = []
            for row in ws_mov.iter_rows(min_row=2, values_only=True):
                # row = (ID, Fecha, Proveedor, Detalle, Obs, Total)
                if len(row) >= 6:
                    mov = {
                        'id': row[0],
                        'fecha': row[1],
                        'proveedor': row[2],
                        'detalle': row[3],
                        'obs': row[4],
                        'total': row[5]
                    }
                    all_movs.append(mov)

            # Aplicar filtros
            def filtrar_mov(m):
                cumple_proveedor = True
                cumple_fecha = True

                if proveedor_filtrado:
                    cumple_proveedor = (m['proveedor'].strip().lower() == proveedor_filtrado.strip().lower())

                if fecha_filtrada:
                    fecha_str = ''
                    if isinstance(m['fecha'], datetime):
                        fecha_str = m['fecha'].strftime('%Y-%m-%d')
                    elif isinstance(m['fecha'], str):
                        fecha_str = m['fecha']
                    cumple_fecha = (fecha_str == fecha_filtrada)

                return cumple_proveedor and cumple_fecha

            movimientos = list(filter(filtrar_mov, all_movs))
    else:
        movimientos = []

    paginator = Paginator(movimientos, 10)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    return render(request, 'formulario.html', {
        'form': form,
        'resumen': resumen,
        'movimientos': page_obj,
        'resumen_Filtrado': resumen_Filtrado,
        'proveedor_filtrado': proveedor_filtrado,
        'fecha_filtrada': fecha_filtrada,
        'paginator': paginator,
        'page_obj': page_obj,
    })

def guardar_movimiento_backup(request):
    if request.method == 'POST':
        form = MovimientoForm(request.POST)
        if form.is_valid():
            proveedor = form.cleaned_data['proveedor']
            detalle = form.cleaned_data['detalle']
            obs = form.cleaned_data['obs']
            fecha = form.cleaned_data['fecha']

            total_raw = request.POST.get('total', '')
            total_str = re.sub(r'[^\d]', '', total_raw)
            total = Decimal(total_str) if total_str else Decimal('0')

            if os.path.exists(RUTA_EXCEL):
                wb = openpyxl.load_workbook(RUTA_EXCEL)
            else:
                wb = openpyxl.Workbook()

            if 'Proveedores' not in wb.sheetnames:
                ws = wb.create_sheet('Proveedores')
                ws.append(['ID', 'Fecha', 'Proveedor', 'Detalle', 'Obs', 'Total'])
            else:
                ws = wb['Proveedores']

            # Generar nuevo ID (puede ser max ID + 1 o max_row - 1)
            ids = [row[0].value for row in ws.iter_rows(min_row=2) if row[0].value is not None]
            nuevo_id = max(ids) + 1 if ids else 1

            ws.append([nuevo_id, fecha, proveedor, detalle, obs, float(total)])

            # Obtener la fila recién agregada
            ultima_fila = ws.max_row

            # Aplicar formato de fecha personalizado a la celda de la columna de fecha (columna 2)
            ws.cell(row=ultima_fila, column=2).number_format = 'D/M/YYYY'

            # Actualizar resumen (igual que antes)
            if 'Resumen' not in wb.sheetnames:
                resumen = wb.create_sheet('Resumen')
                resumen.append(['Proveedor', 'Facturas', 'Ahorros', 'Saldo'])
            else:
                resumen = wb['Resumen']

            datos = {row[0].value: row for row in resumen.iter_rows(min_row=2)}
            if proveedor in datos:
                row = datos[proveedor]
                if detalle == 'Factura':
                    row[1].value += float(total)
                elif detalle == 'Ahorro':
                    row[2].value += float(total)
                row[3].value = row[1].value - row[2].value
            else:
                factura = float(total) if detalle == 'Factura' else 0
                ahorro = float(total) if detalle == 'Ahorro' else 0
                resumen.append([proveedor, factura, ahorro, factura - ahorro])

            wb.save(RUTA_EXCEL)

    return redirect('index')

def resumen(request):
    proveedor_filtrado = request.GET.get('proveedor', None)
    datos = []
    if os.path.exists(RUTA_EXCEL):
        wb = load_workbook(RUTA_EXCEL)
        if 'Resumen' in wb.sheetnames:
            ws = wb['Resumen']
            for row in ws.iter_rows(min_row=2, values_only=True):
                proveedor = row[0]
                factura = row[1] or 0
                ahorro = row[2] or 0
                saldo = row[3] or 0

                if proveedor_filtrado:
                    if proveedor == proveedor_filtrado:
                        datos.append((proveedor, factura, ahorro, saldo))
                else:
                    datos.append((proveedor, factura, ahorro, saldo))

    return render(request, 'resumen.html', {
        'resumen': datos,
        'proveedor_filtrado': proveedor_filtrado,
    })

def formato_pesos(valor):
    return f"${int(valor):,}".replace(",", ".")

def editar_movimiento_backup(request, index):
    wb = openpyxl.load_workbook(RUTA_EXCEL)
    ws = wb['Proveedores']

    # Obtener todas las filas para buscar la fila con ID == index
    rows = list(ws.iter_rows(min_row=2, values_only=True))

    fila_num = None
    for i, row in enumerate(rows, start=2):
        if row[0] == index:
            fila_num = i
            break

    if fila_num is None:
        return redirect('index')  # No encontrado

    row_data = rows[fila_num - 2]

    if request.method == 'POST':
        form = MovimientoForm(request.POST)
        if form.is_valid():
            data = form.cleaned_data

            # Actualizar fila en Excel
            ws.cell(row=fila_num, column=2).value = data['fecha']  # Fecha como texto dd/mm/yyyy
            ws.cell(row=fila_num, column=2).number_format = 'D/M/YYYY' 
            ws.cell(row=fila_num, column=3).value = data['proveedor']
            ws.cell(row=fila_num, column=4).value = data['detalle']
            ws.cell(row=fila_num, column=5).value = data['obs']
            ws.cell(row=fila_num, column=6).value = float(data['total'])

            # Recalcular resumen completo
            facturas_totales = {}
            ahorros_totales = {}

            for r in ws.iter_rows(min_row=2, values_only=True):
                prov = r[2]
                det = r[3]
                tot = float(r[5]) if r[5] else 0

                 # Validar que proveedor exista y no esté vacío o solo espacios
                if not prov or not str(prov).strip():
                    continue

                if prov not in facturas_totales:
                    facturas_totales[prov] = 0
                    ahorros_totales[prov] = 0

                if det == 'Factura':
                    facturas_totales[prov] += tot
                elif det == 'Ahorro':
                    ahorros_totales[prov] += tot

            # Encabezado fijo (solo si se crea la hoja)
            if 'Resumen' not in wb.sheetnames:
                resumen = wb.create_sheet('Resumen')
                resumen.append(['Proveedor', 'Facturas', 'Ahorros', 'Saldo'])
            else:
                resumen = wb['Resumen']

            # Reescribir los valores del resumen sin borrar todo
            fila_resumen = 2  # Comenzar después del encabezado

            for prov in facturas_totales:
                factura = facturas_totales[prov]
                ahorro = ahorros_totales.get(prov, 0)
                saldo = factura - ahorro

                resumen.cell(row=fila_resumen, column=1).value = prov
                resumen.cell(row=fila_resumen, column=2).value = factura
                resumen.cell(row=fila_resumen, column=3).value = ahorro
                resumen.cell(row=fila_resumen, column=4).value = saldo

                fila_resumen += 1

            wb.save(RUTA_EXCEL)
            return redirect('index')
    else:
        fecha_excel = row_data[1]
        fecha_excel = fecha_excel.date()
        initial_data = {
            'fecha': fecha_excel.strftime('%Y-%m-%d') if fecha_excel else '',
            'proveedor': row_data[2],
            'detalle': row_data[3],
            'obs': row_data[4],
            'total': row_data[5],
        }
        form = MovimientoForm(initial=initial_data)

    return render(request, 'editar.html', {'form': form, 'id': index})

def movimientos(request):
    proveedor_filtrado = request.GET.get('proveedor', '').strip()
    fecha_filtrada = request.GET.get('fecha', '').strip()
    movimientos = []
    resumen = []

    if os.path.exists(RUTA_EXCEL):
        wb = openpyxl.load_workbook(RUTA_EXCEL)

        # Cargar resumen para el dropdown de proveedores
        if 'Resumen' in wb.sheetnames:
            ws_resumen = wb['Resumen']
            for row in ws_resumen.iter_rows(min_row=2, values_only=True):
                resumen.append(row)

        # Cargar movimientos y aplicar filtros
        if 'Proveedores' in wb.sheetnames:
            ws_mov = wb['Proveedores']
            all_movs = list(ws_mov.iter_rows(min_row=2, values_only=True))

            for mov in all_movs:
                # mov expected: (ID, Fecha, Proveedor, Detalle, Obs, Total)
                mov_id = mov[0]
                fecha_raw = mov[1]
                proveedor = mov[2]
                detalle = mov[3]
                obs = mov[4]
                total = mov[5]

                # Normalizar fecha
                fecha_mov = None
                if isinstance(fecha_raw, datetime):
                    fecha_mov = fecha_raw.date()
                elif isinstance(fecha_raw, str):
                    try:
                        fecha_mov = datetime.strptime(fecha_raw, "%Y-%m-%d").date()
                    except:
                        try:
                            fecha_mov = datetime.strptime(fecha_raw, "%d/%m/%Y").date()
                        except:
                            fecha_mov = None

                # Filtrar por proveedor
                if proveedor_filtrado and (not proveedor or proveedor.strip().lower() != proveedor_filtrado.lower()):
                    continue

                # Filtrar por fecha
                if fecha_filtrada:
                    try:
                        filtro_fecha = datetime.strptime(fecha_filtrada, "%Y-%m-%d").date()
                        if fecha_mov != filtro_fecha:
                            continue
                    except:
                        pass

                movimientos.append({
                    'id': mov_id,
                    'fecha': fecha_mov,
                    'proveedor': proveedor,
                    'detalle': detalle,
                    'obs': obs,
                    'total': total,
                })

    # Paginación 10 por página
    paginator = Paginator(movimientos, 10)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    return render(request, 'movimientos.html', {
        'resumen': resumen,
        'movimientos': page_obj,
        'proveedor_filtrado': proveedor_filtrado,
        'fecha_filtrada': fecha_filtrada,
        'paginator': paginator,
        'page_obj': page_obj,
    })

def agregar_persona(request):
    if request.method == 'POST':
        form = ProveedorForm(request.POST)
        if form.is_valid():
            nombre = form.cleaned_data['nombre'].strip()

            # Validar si ya existe en la base de datos
            existe_bd = Resumen.objects.filter(proveedor__iexact=nombre).exists()

            if existe_bd:
                form.add_error('nombre', 'El proveedor ya existe en la base de datos.')
            else:
                # Guardar en la base de datos (Resumen)
                with transaction.atomic():
                    resumen_obj = Resumen.objects.create(
                        proveedor=nombre,
                        facturas=0,
                        ahorros=0,
                        saldo=0
                    )

                    # Guardar en Excel
                    if os.path.exists(RUTA_EXCEL):
                        wb = load_workbook(RUTA_EXCEL)
                    else:
                        from openpyxl import Workbook
                        wb = Workbook()

                    if 'Resumen' not in wb.sheetnames:
                        ws = wb.create_sheet('Resumen')
                        ws.append(['Proveedor', 'Total Facturas', 'Total Ahorro', 'Saldo'])
                    else:
                        ws = wb['Resumen']

                    proveedores_excel = [row[0].value for row in ws.iter_rows(min_row=2)]

                    if nombre not in proveedores_excel:
                        ws.append([nombre, 0, 0, 0])
                        wb.save(RUTA_EXCEL)

                return redirect('resumen')  # Redirige a donde quieras mostrar resumen

    else:
        form = ProveedorForm()

    return render(request, 'agregar_persona.html', {'form': form})

def index(request):
    proveedor_filtrado = request.GET.get('proveedor', None)
    proveedores = []
    facturas_totales = []
    ahorros_totales = []
    saldos_totales = []

    facturas_por_mes = {}
    ahorros_por_mes = {}

    if os.path.exists(RUTA_EXCEL):
        wb = load_workbook(RUTA_EXCEL)
        if 'Resumen' in wb.sheetnames:
            ws = wb['Resumen']
            for row in ws.iter_rows(min_row=2, values_only=True):
                proveedor = row[0]
                factura = row[1] or 0
                ahorro = row[2] or 0
                saldo = row[3] or 0

                proveedores.append(proveedor)
                facturas_totales.append(factura)
                ahorros_totales.append(ahorro)
                saldos_totales.append(saldo)

        if 'Proveedores' in wb.sheetnames:
            ws_prov = wb['Proveedores']
            for row in ws_prov.iter_rows(min_row=2, values_only=True):
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
                        dt = datetime.strptime(fecha, '%Y-%m-%d')
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
    return render(request, 'dashboard.html', context)

def guardar_movimiento(request):
    if request.method == 'POST':
        form = MovimientoForm(request.POST)
        if form.is_valid():
            data = form.cleaned_data
            proveedor = data['proveedor']
            detalle = data['detalle']
            obs = data['obs']
            fecha = data['fecha']

            # Normalizar total
            total_raw = request.POST.get('total', '')
            total_str = re.sub(r'[^\d]', '', total_raw)
            total = Decimal(total_str) if total_str else Decimal('0')

            # 1. Guardar en base de datos (modelo Django)
            with transaction.atomic():
                mov = Movimiento.objects.create(
                    fecha=fecha,
                    proveedor=proveedor,
                    detalle=detalle,
                    obs=obs,
                    total=total
                )

                # Actualizar resumen en base de datos
                resumen_obj, created = Resumen.objects.get_or_create(proveedor=proveedor)
                if detalle == 'Factura':
                    resumen_obj.facturas += total
                elif detalle == 'Ahorro':
                    resumen_obj.ahorros += total
                resumen_obj.saldo = resumen_obj.facturas - resumen_obj.ahorros
                resumen_obj.save()

                # 2. Guardar en Excel
                if os.path.exists(RUTA_EXCEL):
                    wb = load_workbook(RUTA_EXCEL)
                else:
                    from openpyxl import Workbook
                    wb = Workbook()

                if 'Proveedores' not in wb.sheetnames:
                    ws = wb.create_sheet('Proveedores')
                    ws.append(['ID', 'Fecha', 'Proveedor', 'Detalle', 'Obs', 'Total'])
                else:
                    ws = wb['Proveedores']

                # El ID que usaremos en Excel es el ID del modelo
                ws.append([mov.id, fecha, proveedor, detalle, obs, float(total)])
                ultima_fila = ws.max_row
                ws.cell(row=ultima_fila, column=2).number_format = 'D/M/YYYY'

                # Actualizar resumen en Excel
                if 'Resumen' not in wb.sheetnames:
                    resumen = wb.create_sheet('Resumen')
                    resumen.append(['Proveedor', 'Facturas', 'Ahorros', 'Saldo'])
                else:
                    resumen = wb['Resumen']

                # Cargar datos actuales del resumen en dict
                datos = {row[0].value: row for row in resumen.iter_rows(min_row=2)}

                if proveedor in datos:
                    row = datos[proveedor]
                    if detalle == 'Factura':
                        row[1].value += float(total)
                    elif detalle == 'Ahorro':
                        row[2].value += float(total)
                    row[3].value = row[1].value - row[2].value
                else:
                    factura = float(total) if detalle == 'Factura' else 0
                    ahorro = float(total) if detalle == 'Ahorro' else 0
                    resumen.append([proveedor, factura, ahorro, factura - ahorro])

                wb.save(RUTA_EXCEL)

            return redirect('index')

    else:
        form = MovimientoForm()

    return render(request, 'formulario.html', {'form': form})

def editar_movimiento(request, index):
    mov = get_object_or_404(Movimiento, pk=index)

    if request.method == 'POST':
        form = MovimientoForm(request.POST)
        if form.is_valid():
            data = form.cleaned_data
            with transaction.atomic():
                # Actualizar modelo
                mov.fecha = data['fecha']
                mov.proveedor = data['proveedor']
                mov.detalle = data['detalle']
                mov.obs = data['obs']
                mov.total = data['total']
                mov.save()

                # Recalcular resumen para este proveedor:
                # Recalcular resumen para este proveedor:
                facturas_totales = Movimiento_Cliente.objects.filter(
                    proveedor=mov.proveedor, detalle='Factura'
                ).aggregate(total=Sum('total'))['total'] or 0

                ahorros_totales = Movimiento_Cliente.objects.filter(
                    proveedor=mov.proveedor, detalle='Ahorro'
                ).aggregate(total=Sum('total'))['total'] or 0

                resumen_obj, _ = Resumen.objects.get_or_create(proveedor=mov.proveedor)
                resumen_obj.facturas = facturas_totales
                resumen_obj.ahorros = ahorros_totales
                resumen_obj.saldo = facturas_totales - ahorros_totales
                resumen_obj.save()

                # Actualizar Excel
                wb = load_workbook(RUTA_EXCEL)
                ws = wb['Proveedores']

                # Buscar fila con ID = index y actualizar
                fila_num = None
                for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
                    if row[0].value == index:
                        fila_num = i
                        break
                if fila_num is None:
                    return redirect('index')  # No encontrado

                ws.cell(row=fila_num, column=2).value = mov.fecha
                ws.cell(row=fila_num, column=2).number_format = 'D/M/YYYY'
                ws.cell(row=fila_num, column=3).value = mov.proveedor
                ws.cell(row=fila_num, column=4).value = mov.detalle
                ws.cell(row=fila_num, column=5).value = mov.obs
                ws.cell(row=fila_num, column=6).value = float(mov.total)

                # Recalcular resumen completo en Excel
                facturas_totales = {}
                ahorros_totales = {}

                for r in ws.iter_rows(min_row=2, values_only=True):
                    prov = r[2]
                    det = r[3]
                    tot = float(r[5]) if r[5] else 0
                    if not prov or not str(prov).strip():
                        continue
                    facturas_totales.setdefault(prov, 0)
                    ahorros_totales.setdefault(prov, 0)
                    if det == 'Factura':
                        facturas_totales[prov] += tot
                    elif det == 'Ahorro':
                        ahorros_totales[prov] += tot

                if 'Resumen' not in wb.sheetnames:
                    resumen = wb.create_sheet('Resumen')
                    resumen.append(['Proveedor', 'Facturas', 'Ahorros', 'Saldo'])
                else:
                    resumen = wb['Resumen']

                # Limpiar contenido resumen (menos encabezado)
                for row in resumen.iter_rows(min_row=2):
                    for cell in row:
                        cell.value = None

                fila_resumen = 2
                for prov in facturas_totales:
                    factura = facturas_totales[prov]
                    ahorro = ahorros_totales.get(prov, 0)
                    saldo = factura - ahorro
                    resumen.cell(row=fila_resumen, column=1).value = prov
                    resumen.cell(row=fila_resumen, column=2).value = factura
                    resumen.cell(row=fila_resumen, column=3).value = ahorro
                    resumen.cell(row=fila_resumen, column=4).value = saldo
                    fila_resumen += 1

                wb.save(RUTA_EXCEL)

            return redirect('index')

    else:
        initial_data = {
            'fecha': mov.fecha.strftime('%Y-%m-%d'),
            'proveedor': mov.proveedor,
            'detalle': mov.detalle,
            'obs': mov.obs,
            'total': mov.total,
        }
        form = MovimientoForm(initial=initial_data)

    return render(request, 'editar.html', {'form': form, 'id': index})


#clientes

def MovimientoCliente(request):
    form = MovimientoForm()
    resumen = []
    resumen_Filtrado = []
    movimientos = []

    proveedor_filtrado = request.GET.get('proveedor', None)
    fecha_filtrada = request.GET.get('fecha', None)

    if os.path.exists(RUTA_EXCEL):
        wb = openpyxl.load_workbook(RUTA_EXCEL)

        # Cargar resumen (igual)
        if 'ResumenCliente' in wb.sheetnames:
            ws_resumen = wb['ResumenCliente']
            for row in ws_resumen.iter_rows(min_row=2, values_only=True):
                resumen.append(row)

        # Filtrar resumen
        if 'ResumenCliente' in wb.sheetnames:
            ws_resumen_Filtrado = wb['ResumenCliente']
            for row in ws_resumen_Filtrado.iter_rows(min_row=2, values_only=True):
                if proveedor_filtrado:
                    if row[0] and row[0].strip().lower() == proveedor_filtrado.strip().lower():
                        resumen_Filtrado.append(row)
                else:
                    resumen_Filtrado.append(row)

        # Cargar movimientos con ID
        if 'ProveedoresCliente' in wb.sheetnames:
            ws_mov = wb['ProveedoresCliente']
            all_movs = []
            for row in ws_mov.iter_rows(min_row=2, values_only=True):
                # row = (ID, Fecha, Proveedor, Detalle, Obs, Total)
                if len(row) >= 6:
                    mov = {
                        'id': row[0],
                        'fecha': row[1],
                        'proveedor': row[2],
                        'detalle': row[3],
                        'obs': row[4],
                        'total': row[5]
                    }
                    all_movs.append(mov)

            # Aplicar filtros
            def filtrar_mov(m):
                cumple_proveedor = True
                cumple_fecha = True

                if proveedor_filtrado:
                    cumple_proveedor = (m['proveedor'].strip().lower() == proveedor_filtrado.strip().lower())

                if fecha_filtrada:
                    fecha_str = ''
                    if isinstance(m['fecha'], datetime):
                        fecha_str = m['fecha'].strftime('%Y-%m-%d')
                    elif isinstance(m['fecha'], str):
                        fecha_str = m['fecha']
                    cumple_fecha = (fecha_str == fecha_filtrada)

                return cumple_proveedor and cumple_fecha

            movimientos = list(filter(filtrar_mov, all_movs))
    else:
        movimientos = []

    paginator = Paginator(movimientos, 10)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    return render(request, 'formularioCliente.html', {
        'form': form,
        'resumen': resumen,
        'movimientos': page_obj,
        'resumen_Filtrado': resumen_Filtrado,
        'proveedor_filtrado': proveedor_filtrado,
        'fecha_filtrada': fecha_filtrada,
        'paginator': paginator,
        'page_obj': page_obj,
    })

def resumenCliente(request):
    proveedor_filtrado = request.GET.get('proveedor', None)
    datos = []
    if os.path.exists(RUTA_EXCEL):
        wb = load_workbook(RUTA_EXCEL)
        if 'ResumenCliente' in wb.sheetnames:
            ws = wb['ResumenCliente']
            for row in ws.iter_rows(min_row=2, values_only=True):
                proveedor = row[0]
                factura = row[1] or 0
                ahorro = row[2] or 0
                saldo = row[3] or 0

                if proveedor_filtrado:
                    if proveedor == proveedor_filtrado:
                        datos.append((proveedor, factura, ahorro, saldo))
                else:
                    datos.append((proveedor, factura, ahorro, saldo))

    return render(request, 'resumenCliente.html', {
        'resumen': datos,
        'proveedor_filtrado': proveedor_filtrado,
    })

def movimientosCliente(request):
    proveedor_filtrado = request.GET.get('proveedor', '').strip()
    fecha_filtrada = request.GET.get('fecha', '').strip()
    movimientos = []
    resumen = []

    if os.path.exists(RUTA_EXCEL):
        wb = openpyxl.load_workbook(RUTA_EXCEL)

        # Cargar resumen para el dropdown de proveedores
        if 'ResumenCliente' in wb.sheetnames:
            ws_resumen = wb['ResumenCliente']
            for row in ws_resumen.iter_rows(min_row=2, values_only=True):
                resumen.append(row)

        # Cargar movimientos y aplicar filtros
        if 'ProveedoresCliente' in wb.sheetnames:
            ws_mov = wb['ProveedoresCliente']
            all_movs = list(ws_mov.iter_rows(min_row=2, values_only=True))

            for mov in all_movs:
                # mov expected: (ID, Fecha, Proveedor, Detalle, Obs, Total)
                mov_id = mov[0]
                fecha_raw = mov[1]
                proveedor = mov[2]
                detalle = mov[3]
                obs = mov[4]
                total = mov[5]

                # Normalizar fecha
                fecha_mov = None
                if isinstance(fecha_raw, datetime):
                    fecha_mov = fecha_raw.date()
                elif isinstance(fecha_raw, str):
                    try:
                        fecha_mov = datetime.strptime(fecha_raw, "%Y-%m-%d").date()
                    except:
                        try:
                            fecha_mov = datetime.strptime(fecha_raw, "%d/%m/%Y").date()
                        except:
                            fecha_mov = None

                # Filtrar por proveedor
                if proveedor_filtrado and (not proveedor or proveedor.strip().lower() != proveedor_filtrado.lower()):
                    continue

                # Filtrar por fecha
                if fecha_filtrada:
                    try:
                        filtro_fecha = datetime.strptime(fecha_filtrada, "%Y-%m-%d").date()
                        if fecha_mov != filtro_fecha:
                            continue
                    except:
                        pass

                movimientos.append({
                    'id': mov_id,
                    'fecha': fecha_mov,
                    'proveedor': proveedor,
                    'detalle': detalle,
                    'obs': obs,
                    'total': total,
                })

    # Paginación 10 por página
    paginator = Paginator(movimientos, 10)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    return render(request, 'movimientosCliente.html', {
        'resumen': resumen,
        'movimientos': page_obj,
        'proveedor_filtrado': proveedor_filtrado,
        'fecha_filtrada': fecha_filtrada,
        'paginator': paginator,
        'page_obj': page_obj,
    })

def agregar_persona_Cliente(request):
    if request.method == 'POST':
        form = ProveedorForm(request.POST)
        if form.is_valid():
            nombre = form.cleaned_data['nombre'].strip()

            # Validar si ya existe en la base de datos
            existe_bd = Resumen_Cliente.objects.filter(proveedor__iexact=nombre).exists()

            if existe_bd:
                form.add_error('nombre', 'El proveedor ya existe en la base de datos.')
            else:
                # Guardar en la base de datos (Resumen)
                with transaction.atomic():
                    resumen_obj = Resumen_Cliente.objects.create(
                        proveedor=nombre,
                        facturas=0,
                        ahorros=0,
                        saldo=0
                    )

                    # Guardar en Excel
                    if os.path.exists(RUTA_EXCEL):
                        wb = load_workbook(RUTA_EXCEL)
                    else:
                        from openpyxl import Workbook
                        wb = Workbook()

                    if 'ResumenCliente' not in wb.sheetnames:
                        ws = wb.create_sheet('ResumenCliente')
                        ws.append(['Proveedor', 'Total Facturas', 'Total Ahorro', 'Saldo'])
                    else:
                        ws = wb['ResumenCliente']

                    proveedores_excel = [row[0].value for row in ws.iter_rows(min_row=2)]

                    if nombre not in proveedores_excel:
                        ws.append([nombre, 0, 0, 0])
                        wb.save(RUTA_EXCEL)

                return redirect('resumenCliente')  # Redirige a donde quieras mostrar resumen

    else:
        form = ProveedorForm()

    return render(request, 'agregar_personaCliente.html', {'form': form})

def dashboardCliente(request):
    proveedor_filtrado = request.GET.get('proveedor', None)
    proveedores = []
    facturas_totales = []
    ahorros_totales = []
    saldos_totales = []

    facturas_por_mes = {}
    ahorros_por_mes = {}

    if os.path.exists(RUTA_EXCEL):
        wb = load_workbook(RUTA_EXCEL)
        if 'ResumenCliente' in wb.sheetnames:
            ws = wb['ResumenCliente']
            for row in ws.iter_rows(min_row=2, values_only=True):
                proveedor = row[0]
                factura = row[1] or 0
                ahorro = row[2] or 0
                saldo = row[3] or 0

                proveedores.append(proveedor)
                facturas_totales.append(factura)
                ahorros_totales.append(ahorro)
                saldos_totales.append(saldo)

        if 'ProveedoresCliente' in wb.sheetnames:
            ws_prov = wb['ProveedoresCliente']
            for row in ws_prov.iter_rows(min_row=2, values_only=True):
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
                        dt = datetime.strptime(fecha, '%Y-%m-%d')
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
    return render(request, 'dashboardCliente.html', context)

def guardar_movimiento_cliente(request):
    if request.method == 'POST':
        form = MovimientoClienteForm(request.POST)
        if form.is_valid():
            data = form.cleaned_data
            proveedor = data['proveedor']
            detalle = data['detalle']
            obs = data['obs']
            fecha = data['fecha']

            # Normalizar total
            total_raw = request.POST.get('total', '')
            total_str = re.sub(r'[^\d]', '', total_raw)
            total = Decimal(total_str) if total_str else Decimal('0')

            # 1. Guardar en base de datos (modelo Django)
            with transaction.atomic():
                mov = Movimiento_Cliente.objects.create(
                    fecha=fecha,
                    proveedor=proveedor,
                    detalle=detalle,
                    obs=obs,
                    total=total
                )

                # Actualizar resumen en base de datos
                resumen_obj, created = Resumen_Cliente.objects.get_or_create(proveedor=proveedor)
                if detalle == 'Factura':
                    resumen_obj.facturas += total
                elif detalle == 'Ahorro':
                    resumen_obj.ahorros += total
                resumen_obj.saldo = resumen_obj.facturas - resumen_obj.ahorros
                resumen_obj.save()

                # 2. Guardar en Excel
                if os.path.exists(RUTA_EXCEL):
                    wb = load_workbook(RUTA_EXCEL)
                else:
                    from openpyxl import Workbook
                    wb = Workbook()

                if 'ProveedoresCliente' not in wb.sheetnames:
                    ws = wb.create_sheet('ProveedoresCliente')
                    ws.append(['ID', 'Fecha', 'Proveedor', 'Detalle', 'Obs', 'Total'])
                else:
                    ws = wb['ProveedoresCliente']

                # El ID que usaremos en Excel es el ID del modelo
                ws.append([mov.id, fecha, proveedor, detalle, obs, float(total)])
                ultima_fila = ws.max_row
                ws.cell(row=ultima_fila, column=2).number_format = 'D/M/YYYY'

                # Actualizar resumen en Excel
                if 'ResumenCliente' not in wb.sheetnames:
                    resumen = wb.create_sheet('ResumenCliente')
                    resumen.append(['Proveedor', 'Facturas', 'Ahorros', 'Saldo'])
                else:
                    resumen = wb['ResumenCliente']

                # Cargar datos actuales del resumen en dict
                datos = {row[0].value: row for row in resumen.iter_rows(min_row=2)}

                if proveedor in datos:
                    row = datos[proveedor]
                    if detalle == 'Factura':
                        row[1].value += float(total)
                    elif detalle == 'Ahorro':
                        row[2].value += float(total)
                    row[3].value = row[1].value - row[2].value
                else:
                    factura = float(total) if detalle == 'Factura' else 0
                    ahorro = float(total) if detalle == 'Ahorro' else 0
                    resumen.append([proveedor, factura, ahorro, factura - ahorro])

                wb.save(RUTA_EXCEL)

            return redirect('MovimientoCliente')

    else:
        form = MovimientoForm()

    return render(request, 'formularioCliente.html', {'form': form})

def editar_movimiento_Cliente(request, index):
    mov = get_object_or_404(Movimiento_Cliente, pk=index)

    if request.method == 'POST':
        form = MovimientoClienteForm(request.POST)
        if form.is_valid():
            data = form.cleaned_data
            with transaction.atomic():
                # Actualizar modelo
                mov.fecha = data['fecha']
                mov.proveedor = data['proveedor']
                mov.detalle = data['detalle']
                mov.obs = data['obs']
                mov.total = data['total']
                mov.save()

                # Recalcular resumen para este proveedor:
                facturas_totales = Movimiento_Cliente.objects.filter(
                    proveedor=mov.proveedor, detalle='Factura'
                ).aggregate(total=Sum('total'))['total'] or 0

                ahorros_totales = Movimiento_Cliente.objects.filter(
                    proveedor=mov.proveedor, detalle='Ahorro'
                ).aggregate(total=Sum('total'))['total'] or 0


                resumen_obj, _ = Resumen_Cliente.objects.get_or_create(proveedor=mov.proveedor)
                resumen_obj.facturas = facturas_totales
                resumen_obj.ahorros = ahorros_totales
                resumen_obj.saldo = facturas_totales - ahorros_totales
                resumen_obj.save()

                # Actualizar Excel
                wb = load_workbook(RUTA_EXCEL)
                ws = wb['ProveedoresCliente']

                # Buscar fila con ID = index y actualizar
                fila_num = None
                for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
                    if row[0].value == index:
                        fila_num = i
                        break
                if fila_num is None:
                    return redirect('index')  # No encontrado

                ws.cell(row=fila_num, column=2).value = mov.fecha
                ws.cell(row=fila_num, column=2).number_format = 'D/M/YYYY'
                ws.cell(row=fila_num, column=3).value = mov.proveedor
                ws.cell(row=fila_num, column=4).value = mov.detalle
                ws.cell(row=fila_num, column=5).value = mov.obs
                ws.cell(row=fila_num, column=6).value = float(mov.total)

                # Recalcular resumen completo en Excel
                facturas_totales = {}
                ahorros_totales = {}

                for r in ws.iter_rows(min_row=2, values_only=True):
                    prov = r[2]
                    det = r[3]
                    tot = float(r[5]) if r[5] else 0
                    if not prov or not str(prov).strip():
                        continue
                    facturas_totales.setdefault(prov, 0)
                    ahorros_totales.setdefault(prov, 0)
                    if det == 'Factura':
                        facturas_totales[prov] += tot
                    elif det == 'Ahorro':
                        ahorros_totales[prov] += tot

                if 'ResumenCliente' not in wb.sheetnames:
                    resumen = wb.create_sheet('ResumenCliente')
                    resumen.append(['Proveedor', 'Facturas', 'Ahorros', 'Saldo'])
                else:
                    resumen = wb['ResumenCliente']

                # Limpiar contenido resumen (menos encabezado)
                for row in resumen.iter_rows(min_row=2):
                    for cell in row:
                        cell.value = None

                fila_resumen = 2
                for prov in facturas_totales:
                    factura = facturas_totales[prov]
                    ahorro = ahorros_totales.get(prov, 0)
                    saldo = factura - ahorro
                    resumen.cell(row=fila_resumen, column=1).value = prov
                    resumen.cell(row=fila_resumen, column=2).value = factura
                    resumen.cell(row=fila_resumen, column=3).value = ahorro
                    resumen.cell(row=fila_resumen, column=4).value = saldo
                    fila_resumen += 1

                wb.save(RUTA_EXCEL)

            return redirect('index')

    else:
        initial_data = {
            'fecha': mov.fecha.strftime('%Y-%m-%d'),
            'proveedor': mov.proveedor,
            'detalle': mov.detalle,
            'obs': mov.obs,
            'total': mov.total,
        }
        form = MovimientoClienteForm(initial=initial_data)

    return render(request, 'editarCliente.html', {'form': form, 'id': index})
