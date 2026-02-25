"""Column loops, conditional columns, and cross-dimensional templates example.

Uses column_loops_template.xlsx which contains:

  Sheet "Quarterly Report"
    - {%col for q in quarters %} to expand quarter columns horizontally
    - Body columns have {{ q.name }}, {{ q.revenue }}, {{ q.expenses }}, {{ q.profit }}

  Sheet "Conditional"
    - {%col if show_discount %} to conditionally show a Discount column
    - {% for p in products %} row loop for product rows

  Sheet "Pivot Table"
    - {%col for q in quarters %} expands quarter columns
    - {% for m in metrics %} expands metric rows
    - {{ data[m.key][q.key] }} references both row and column variables

Run from the repo root:
    python examples/column_loops.py
"""

from pathlib import Path

from xlsxtpl import XlsxTemplate

EXAMPLES_DIR = Path(__file__).parent
TEMPLATE_PATH = EXAMPLES_DIR / "column_loops_template.xlsx"
OUTPUT_PATH = EXAMPLES_DIR / "column_loops_output.xlsx"

context = {
    "title": "2025 Quarterly Report",
    "quarters": [
        {"name": "Q1", "key": "q1", "revenue": 1200000, "expenses": 800000, "profit": 400000},
        {"name": "Q2", "key": "q2", "revenue": 1500000, "expenses": 900000, "profit": 600000},
        {"name": "Q3", "key": "q3", "revenue": 1800000, "expenses": 1000000, "profit": 800000},
        {"name": "Q4", "key": "q4", "revenue": 2100000, "expenses": 1100000, "profit": 1000000},
    ],
    "show_discount": True,
    "products": [
        {"name": "Widget", "price": 29.99, "discount": "10%", "in_stock": True},
        {"name": "Gadget", "price": 49.99, "discount": "15%", "in_stock": True},
        {"name": "Doohickey", "price": 9.99, "discount": "5%", "in_stock": False},
    ],
    # Pivot table data: metrics × quarters
    "metrics": [
        {"name": "Revenue", "key": "revenue"},
        {"name": "Expenses", "key": "expenses"},
        {"name": "Profit", "key": "profit"},
    ],
    "data": {
        "revenue": {"q1": 1200000, "q2": 1500000, "q3": 1800000, "q4": 2100000},
        "expenses": {"q1": 800000, "q2": 900000, "q3": 1000000, "q4": 1100000},
        "profit": {"q1": 400000, "q2": 600000, "q3": 800000, "q4": 1000000},
    },
}

tpl = XlsxTemplate(TEMPLATE_PATH)
tpl.render(context)
tpl.save(OUTPUT_PATH)
print(f"Rendered → {OUTPUT_PATH}")
