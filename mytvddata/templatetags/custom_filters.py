# app/templatetags/custom_filters.py
from django import template

register = template.Library()

@register.filter
def get_result_for_year(results, year):
    return next((r for r in results if int(r.time_dim) == int(year)), None)


@register.filter
def get_item(dictionary, key):
    """Permet d'accéder à dictionary[key] dans les templates."""
    return dictionary.get(key)
