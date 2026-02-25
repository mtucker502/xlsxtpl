# xlsxtpl

Jinja2 templating for Excel `.xlsx` files. Like [docxtpl](https://github.com/elapouya/python-docxtpl) but for spreadsheets.

- [Install](#install)
- [Quick start](#quick-start)
- [Template syntax](#template-syntax)
- [Row loops](#row-loops)
- [Conditional rows](#conditional-rows)
- [Custom filters](#custom-filters)

## Install

```bash
uv add git+https://github.com/mtucker502/xlsxtpl.git
```

## Quick start

Create an `.xlsx` template in Excel containing Jinja2 tags in cells. Then render it:

```python
from xlsxtpl import XlsxTemplate

tpl = XlsxTemplate("template.xlsx")
tpl.render({
    "title": "Q4 Review",
    "author": "Jane Smith",
    "revenue": 4200000,
})
tpl.save("output.xlsx")
```

Cell formatting (bold, font, colors, number formats) is preserved through rendering. All worksheets in the workbook are rendered with the same context.

## Template syntax

Standard Jinja2 syntax works inside any cell.

### Variables

```
{{ title }}
{{ metrics.revenue }}
{{ items.0.name }}
```

Pure expressions like `{{ count }}` preserve the Python type (int, float, bool, date, etc.). Mixed content like `Total: {{ count }}` renders as a string.

### Conditionals

```
{% if show_summary %}
{{ summary }}
{% endif %}
```

### For loops

```
{% for item in items %}
{{ item.name }}    {{ item.qty }}
{% endfor %}
```

### Filters

```
{{ name|upper }}
{{ items|length }}
{{ description|default("N/A") }}
```

## Row loops

Use `{% for %}` to duplicate rows for each item in a list. Place the opening tag in its own row, the body rows below, and `{% endfor %}` in a closing row.

**In the template:**

| | A | B | C |
|---|---|---|---|
| 1 | `{% for m in metrics %}` | | |
| 2 | `{{ m.name }}` | `{{ m.value }}` | `{{ m.status }}` |
| 3 | `{% endfor %}` | | |

**Render:**

```python
tpl = XlsxTemplate("template.xlsx")
tpl.render({
    "metrics": [
        {"name": "Revenue", "value": "$4.2M", "status": "On track"},
        {"name": "NPS", "value": "72", "status": "Above target"},
        {"name": "Churn", "value": "3.1%", "status": "At risk"},
    ],
})
tpl.save("output.xlsx")
# → 3 rows, one per metric. Directive rows are removed.
```

Row heights and styles are preserved on duplicated rows. Loops can be nested.

Standard Jinja2 loop variables are available:

| Variable | Description |
|---|---|
| `loop.index` | 1-based iteration count |
| `loop.index0` | 0-based iteration count |
| `loop.first` | `True` on the first iteration |
| `loop.last` | `True` on the last iteration |
| `loop.length` | Total number of items |

## Conditional rows

Use `{% if %}` to conditionally include or exclude rows.

**In the template:**

| | A | B |
|---|---|---|
| 1 | `{% if show_detail %}` | |
| 2 | Detail: | `{{ detail }}` |
| 3 | `{% endif %}` | |

**Render:**

```python
tpl = XlsxTemplate("template.xlsx")

# Rows included — show_detail is truthy
tpl.render({"show_detail": True, "detail": "important"})

# Rows removed — show_detail is falsy
tpl.render({"show_detail": False})
```

## Custom filters

Add custom Jinja2 filters via the `jinja_env`:

```python
tpl = XlsxTemplate("template.xlsx")
tpl.jinja_env.filters["currency"] = lambda v: f"${v:,.2f}"
tpl.render({"price": 4200000})
tpl.save("output.xlsx")
```

Built-in filters: `date` (format dates), `number_format` (thousands separators).
