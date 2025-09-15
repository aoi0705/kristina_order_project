from django import template

register = template.Library()

@register.filter
def get_attr(obj, name):
    """obj の属性 name（str）を取り出す。無ければ空文字。"""
    try:
        return getattr(obj, name, "") or ""
    except Exception:
        return ""