from django.db import models

class Movimiento(models.Model):
    fecha = models.DateField()
    proveedor = models.CharField(max_length=255)
    detalle = models.CharField(max_length=50)
    obs = models.TextField(blank=True, null=True)
    total = models.DecimalField(max_digits=15, decimal_places=0)
    id_factura = models.CharField(max_length=20, blank=True, null=True)  # NUEVO
    estado = models.CharField(max_length=20, blank=True, null=True)      # NUEVO ('Activa', 'Inactiva', 'Abonado')

    def __str__(self):
        return f"{self.fecha} - {self.proveedor} - {self.total}"

class Resumen(models.Model):
    proveedor = models.CharField(max_length=255, unique=True)
    facturas = models.DecimalField(max_digits=15, decimal_places=0, default=0)
    Abonos = models.DecimalField(max_digits=15, decimal_places=0, default=0)
    saldo = models.DecimalField(max_digits=15, decimal_places=0, default=0)

    def __str__(self):
        return f"{self.proveedor} - Saldo: {self.saldo}"


class Movimiento_Cliente(models.Model):
    fecha = models.DateField()
    proveedor = models.CharField(max_length=255)
    detalle = models.CharField(max_length=50)
    obs = models.TextField(blank=True, null=True)
    total = models.DecimalField(max_digits=15, decimal_places=0)
    id_factura = models.CharField(max_length=20, blank=True, null=True)  # NUEVO
    estado = models.CharField(max_length=20, blank=True, null=True)      # NUEVO ('Activa', 'Inactiva', 'Abonado')x

    def __str__(self):
        return f"{self.fecha} - {self.proveedor} - {self.total}"

class Resumen_Cliente(models.Model):
    proveedor = models.CharField(max_length=255, unique=True)
    facturas = models.DecimalField(max_digits=15, decimal_places=0, default=0)
    Abonos = models.DecimalField(max_digits=15, decimal_places=0, default=0)
    saldo = models.DecimalField(max_digits=15, decimal_places=0, default=0)

    def __str__(self):
        return f"{self.proveedor} - Saldo: {self.saldo}"

class Gasto(models.Model):
    CATEGORIAS = [
        ('Parqueadero', 'Parqueadero'),
        ('Flete', 'Flete'),
        ('Varios', 'Varios'),
    ]

    fecha = models.DateField()
    categoria = models.CharField(max_length=20, choices=CATEGORIAS)
    placa = models.CharField(max_length=10, blank=True, null=True)
    conductor = models.CharField(max_length=255, blank=True, null=True)
    precio = models.DecimalField(max_digits=15, decimal_places=0)

    def __str__(self):
        return f"{self.fecha} - {self.categoria} - {self.precio}"