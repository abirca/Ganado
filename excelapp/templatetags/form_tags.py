from django import template

register = template.Library()

@register.filter(name='add_class')
def add_class(field, css_class):
    return field.as_widget(attrs={"class": css_class})

@register.filter
def filter_placa(gastos_list, placa):
    return [g for g in gastos_list if g.get('placa') == placa]

@register.filter
def avg_gasto(gastos_list):
    if not gastos_list:
        return 0
    return sum(g['precio'] for g in gastos_list) / len(gastos_list)

@register.filter
def map_first(list_of_tuples):
    return [item[0] for item in list_of_tuples]

@register.filter
def map_second(list_of_tuples):
    return [item[1] for item in list_of_tuples]
