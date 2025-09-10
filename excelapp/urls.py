
from django.urls import path
from . import views

app_name = 'mi_app'

urlpatterns = [
    # Página principal
    path('', views.index, name='index'),

    # Proveedores
    path('proveedores/movimiento/agregar/', views.MovimientoProveedor, name='movimiento_proveedor_agregar'),
    path('proveedores/movimiento/guardar/', views.guardar_movimiento, name='movimiento_proveedor_guardar'),
    path('proveedores/resumen/', views.resumen, name='proveedores_resumen'),
    path('proveedores/movimientos/', views.movimientos, name='proveedores_movimientos'),
    path('proveedores/movimiento/editar/<int:index>/', views.editar_movimiento, name='movimiento_proveedor_editar'),
    path('proveedores/dashboard/', views.dashboardProveedor, name='proveedores_dashboard'),
    path('proveedores/persona/agregar/', views.agregar_persona, name='proveedor_persona_agregar'),
    path('proveedores/persona/editar/<int:id>/', views.editar_proveedor, name='proveedor_persona_editar'),
    path('proveedores/descargar-excel/', views.descargar_excel_proveedor, name='proveedores_descargar_excel'),


    # Clientes
    path('clientes/movimiento/agregar/', views.MovimientoCliente, name='movimiento_cliente_agregar'),
    path('clientes/movimiento/guardar/', views.guardar_movimiento_cliente, name='movimiento_cliente_guardar'),
    path('clientes/resumen/', views.resumenCliente, name='clientes_resumen'),
    path('clientes/movimientos/', views.movimientosCliente, name='clientes_movimientos'),
    path('clientes/movimiento/editar/<int:index>/', views.editar_movimiento_Cliente, name='movimiento_cliente_editar'),
    path('clientes/persona/agregar/', views.agregar_persona_Cliente, name='cliente_persona_agregar'),
    path('clientes/dashboard/', views.dashboardCliente, name='clientes_dashboard'),
    path('clientes/descargar-excel/', views.descargar_excel_cliente, name='clientes_descargar_excel'),
    path('clientes/persona/editar/<int:id>/', views.editar_cliente, name='cliente_persona_editar'),
    
    # Gastos
    path('gastos/', views.gastos, name='gastos_lista'),
    path('gastos/agregar/', views.agregar_gasto, name='gasto_agregar'),
    path('gastos/editar/<int:id>/', views.editar_gasto, name='gasto_editar'),
    path('gastos/eliminar/<int:id>/', views.eliminar_gasto, name='gasto_eliminar'),
    path('gastos/resumen/', views.resumen_gastos, name='gastos_resumen'),
    path('gastos/dashboard/', views.dashboard_gastos, name='gastos_dashboard'),

]