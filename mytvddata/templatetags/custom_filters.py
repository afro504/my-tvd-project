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


@register.filter
def get_result_for_allyear(results, year):
    try:
        return next(
            (r for r in results if r.get("time_dim") is not None and int(r["time_dim"]) == int(year)),
            None
        )
    except (ValueError, TypeError):
        return None
