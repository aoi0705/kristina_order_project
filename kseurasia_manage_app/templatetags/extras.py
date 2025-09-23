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

from django import template
register = template.Library()

@register.filter(name="intcomma")
def intcomma(value):
    """
    settings.py を変更せずに数値にカンマ区切りを付ける簡易フィルタ。
    数値でなければそのまま返す。
    """
    try:
        n = int(float(str(value).replace(",", "").strip()))
    except (TypeError, ValueError):
        return value
    return f"{n:,}"
