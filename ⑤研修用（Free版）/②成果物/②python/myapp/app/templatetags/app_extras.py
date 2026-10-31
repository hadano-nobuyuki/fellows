from django import template

register = template.Library()


@register.filter
def comma(value):
    """3 桁カンマ区切り(例: 1,500)。None は空文字。"""
    if value is None or value == '':
        return ''
    return f'{int(value):,}'


@register.filter
def signed_comma(value):
    """符号付き・カンマ区切り(例: +150 / -2,000 / 0)。None は空文字。"""
    if value is None or value == '':
        return ''
    value = int(value)
    if value > 0:
        return f'+{value:,}'
    return f'{value:,}'
