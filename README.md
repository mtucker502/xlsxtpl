# xlsxtpl

Jinja2 templating for Excel `.xlsx` files. Like [docxtpl](https://github.com/elapouya/python-docxtpl) but for spreadsheets.

- [Install](#install)
- [Quick start](#quick-start)
- [Template syntax](#template-syntax)
- [Row loops](#row-loops)
- [Conditional rows](#conditional-rows)
- [Column loops](#column-loops)
- [Conditional columns](#conditional-columns)
- [Cross-dimensional templates](#cross-dimensional-templates)
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

## Column loops

Use `{%col for %}` to duplicate columns for each item in a list. Place the opening tag in its own column, the body columns to the right, and `{%col endfor %}` in a closing column. The directive columns are removed after processing.

**In the template:**

| | A | B | C | D |
|---|---|---|---|---|
| 1 | `{%col for q in quarters %}` | `{{ q.name }}` | `{{ q.revenue }}` | `{%col endfor %}` |
| 2 | | `{{ q.profit }}` | `{{ q.margin }}` | |

**Render:**

```python
tpl = XlsxTemplate("template.xlsx")
tpl.render({
    "quarters": [
        {"name": "Q1", "revenue": "$1.2M", "profit": "$400K", "margin": "33%"},
        {"name": "Q2", "revenue": "$1.5M", "profit": "$500K", "margin": "33%"},
        {"name": "Q3", "revenue": "$1.8M", "profit": "$700K", "margin": "39%"},
    ],
})
tpl.save("output.xlsx")
# → 6 body columns (2 per quarter). Directive columns are removed.
```

Column widths and styles are preserved on duplicated columns. Standard `loop` variables (`loop.index`, `loop.first`, etc.) are available.

> **Processing order:** Column blocks (`{%col %}`) are expanded before row blocks (`{% %}`). Column expansion is purely structural — columns are duplicated but cell rendering is deferred. This means row loops can appear alongside column loops, and cells can reference variables from both axes (see [Cross-dimensional templates](#cross-dimensional-templates)).

## Conditional columns

Use `{%col if %}` to conditionally include or exclude columns.

**In the template:**

| | A | B | C |
|---|---|---|---|
| 1 | `{%col if show_detail %}` | Detail | `{%col endif %}` |

**Render:**

```python
tpl = XlsxTemplate("template.xlsx")

# Column included — show_detail is truthy
tpl.render({"show_detail": True})

# Column removed — show_detail is falsy
tpl.render({"show_detail": False})
```

## Cross-dimensional templates

Combine `{%col for %}` and `{% for %}` on the same sheet to build pivot tables and other two-dimensional layouts. Column loops expand columns structurally, then row loops expand rows — cells that reference variables from both axes are rendered with the full merged context.

**In the template:**

| | A | B | C | D |
|---|---|---|---|---|
| 1 | Metric | `{%col for q in quarters %}` | `{{ q.name }}` | `{%col endfor %}` |
| 2 | `{% for m in metrics %}` | | | |
| 3 | `{{ m.name }}` | | `{{ data[m.key][q.key] }}` | |
| 4 | `{% endfor %}` | | | |

Place the cross-dimensional expression (`{{ data[m.key][q.key] }}`) in a column-loop body column so it gets duplicated per column iteration. Row-loop directives go in a column outside the column loop (column A here) so they survive directive-column removal.

**Render:**

```python
tpl = XlsxTemplate("template.xlsx")
tpl.render({
    "quarters": [
        {"name": "Q1", "key": "q1"},
        {"name": "Q2", "key": "q2"},
    ],
    "metrics": [
        {"name": "Revenue", "key": "revenue"},
        {"name": "Profit", "key": "profit"},
    ],
    "data": {
        "revenue": {"q1": 100, "q2": 200},
        "profit": {"q1": 10, "q2": 20},
    },
})
tpl.save("output.xlsx")
# → 2 quarter columns × 2 metric rows, with cross-referenced values.
```

### `col_loop` variable

Inside a column loop, both `loop` and `col_loop` refer to the column loop metadata. In cross-dimensional cells (inside both a row loop and a column loop), `loop` refers to the **row** loop and `col_loop` refers to the **column** loop:

| Variable | In col-only cell | In cross-dimensional cell |
|---|---|---|
| `loop` | Column loop | Row loop |
| `col_loop` | Column loop | Column loop |

```
r{{ loop.index }}-c{{ col_loop.index }}
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
