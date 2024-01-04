from django import template
import ast

register = template.Library()

@register.filter(name='extract_categories')
def extract_categories(value):
    try:
        category_list = ast.literal_eval(value)
        cleaned_categories = [category.replace("Centrepoint - ", "").replace(" - PLM", "").strip(' " ') for category in category_list]
        result = " | ".join(cleaned_categories)
        return result
    except (ValueError, SyntaxError):
        return value
