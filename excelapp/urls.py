from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('MovimientoProveedor/', views.MovimientoProveedor, name='MovimientoProveedor'),
    path('guardar/', views.guardar_movimiento, name='guardar_movimiento'),
    path('resumen/', views.resumen, name='resumen'),
    path('editar/<int:index>/', views.editar_movimiento, name='editar_movimiento'),
    path('movimientos/', views.movimientos, name='movimientos'),
    path('agregar_persona/', views.agregar_persona, name='agregar_persona'),
    #Cliente
    path('MovimientoCliente/', views.MovimientoCliente, name='MovimientoCliente'),
    path('guardarCliente/', views.guardar_movimiento_cliente, name='guardar_movimiento_cliente'),
    path('resumenCliente/', views.resumenCliente, name='resumenCliente'),
    path('editarCliente/<int:index>/', views.editar_movimiento_Cliente, name='editar_movimiento_Cliente'),
    path('movimientosCliente/', views.movimientosCliente, name='movimientosCliente'),
    path('agregar_persona_Cliente/', views.agregar_persona_Cliente, name='agregar_persona_Cliente'),
    path('dashboardCliente/', views.dashboardCliente, name='dashboardCliente'),

]