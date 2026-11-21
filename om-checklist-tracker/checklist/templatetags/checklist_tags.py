from django import template

register = template.Library()


@register.filter
def get_item(mapping, item_id):
    if not hasattr(mapping, "get"):
        return ""
    value = mapping.get(str(item_id), mapping.get(f"item_{item_id}", ""))
    return value.get("value", "") if isinstance(value, dict) else value


@register.filter
def get_comment(mapping, item_id):
    if not hasattr(mapping, "get"):
        return ""
    value = mapping.get(str(item_id), mapping.get(f"comment_{item_id}", ""))
    return value.get("comment", "") if isinstance(value, dict) else value
