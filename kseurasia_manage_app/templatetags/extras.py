from django import template

register = template.Library()

@register.filter
def attr(obj, name):
    """モデルや辞書から動的に属性/キーを取り出す"""
    if obj is None or not name:
        return ""
    # dict もサポート
    if isinstance(obj, dict):
        return obj.get(name, "")
    # モデル/オブジェクト
    return getattr(obj, name, "")
