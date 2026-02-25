import openpyxl
import pytest

from xlsxtpl.exceptions import TemplateRenderError, TemplateSyntaxError
from xlsxtpl.jinja_env import create_jinja_env
from xlsxtpl.renderer import SheetRenderer


@pytest.fixture
def env():
    return create_jinja_env()


def make_ws(cells: dict[str, object]) -> openpyxl.worksheet.worksheet.Worksheet:
    """Helper to create a worksheet from a dict of cell coords to values."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for coord, value in cells.items():
        ws[coord] = value
    return ws


class TestSimpleVariables:
    def test_pure_expression_preserves_type(self, env):
        ws = make_ws({"A1": "{{ count }}"})
        SheetRenderer(ws, env).render({"count": 42})
        assert ws["A1"].value == 42

    def test_string_variable(self, env):
        ws = make_ws({"A1": "{{ name }}"})
        SheetRenderer(ws, env).render({"name": "Alice"})
        assert ws["A1"].value == "Alice"

    def test_mixed_content(self, env):
        ws = make_ws({"A1": "Hello {{ name }}!"})
        SheetRenderer(ws, env).render({"name": "Bob"})
        assert ws["A1"].value == "Hello Bob!"

    def test_multiple_expressions_in_cell(self, env):
        ws = make_ws({"A1": "{{ first }} {{ last }}"})
        SheetRenderer(ws, env).render({"first": "Jane", "last": "Doe"})
        assert ws["A1"].value == "Jane Doe"

    def test_non_template_cells_untouched(self, env):
        ws = make_ws({"A1": "plain text", "A2": 42, "A3": "{{ x }}"})
        SheetRenderer(ws, env).render({"x": "ok"})
        assert ws["A1"].value == "plain text"
        assert ws["A2"].value == 42
        assert ws["A3"].value == "ok"

    def test_float_preserved(self, env):
        ws = make_ws({"A1": "{{ price }}"})
        SheetRenderer(ws, env).render({"price": 9.99})
        assert ws["A1"].value == 9.99

    def test_bool_preserved(self, env):
        ws = make_ws({"A1": "{{ flag }}"})
        SheetRenderer(ws, env).render({"flag": True})
        assert ws["A1"].value is True

    def test_none_preserved(self, env):
        ws = make_ws({"A1": "{{ val }}"})
        SheetRenderer(ws, env).render({"val": None})
        assert ws["A1"].value is None


class TestForLoops:
    def test_simple_loop(self, env):
        ws = make_ws({
            "A1": "Header",
            "A2": "{% for item in items %}",
            "A3": "{{ item }}",
            "A4": "{% endfor %}",
            "A5": "Footer",
        })
        SheetRenderer(ws, env).render({"items": ["a", "b", "c"]})

        assert ws["A1"].value == "Header"
        assert ws["A2"].value == "a"
        assert ws["A3"].value == "b"
        assert ws["A4"].value == "c"
        assert ws["A5"].value == "Footer"

    def test_loop_with_multiple_columns(self, env):
        ws = make_ws({
            "A1": "{% for item in items %}",
            "A2": "{{ item.name }}",
            "B2": "{{ item.price }}",
            "A3": "{% endfor %}",
        })
        items = [{"name": "Widget", "price": 10}, {"name": "Gadget", "price": 20}]
        SheetRenderer(ws, env).render({"items": items})

        assert ws["A1"].value == "Widget"
        assert ws["B1"].value == 10
        assert ws["A2"].value == "Gadget"
        assert ws["B2"].value == 20

    def test_empty_loop_removes_block(self, env):
        ws = make_ws({
            "A1": "Header",
            "A2": "{% for item in items %}",
            "A3": "{{ item }}",
            "A4": "{% endfor %}",
            "A5": "Footer",
        })
        SheetRenderer(ws, env).render({"items": []})

        assert ws["A1"].value == "Header"
        assert ws["A2"].value == "Footer"

    def test_single_item_loop(self, env):
        ws = make_ws({
            "A1": "{% for x in items %}",
            "A2": "{{ x }}",
            "A3": "{% endfor %}",
        })
        SheetRenderer(ws, env).render({"items": ["only"]})
        assert ws["A1"].value == "only"

    def test_loop_index(self, env):
        ws = make_ws({
            "A1": "{% for x in items %}",
            "A2": "{{ loop.index }}",
            "B2": "{{ x }}",
            "A3": "{% endfor %}",
        })
        SheetRenderer(ws, env).render({"items": ["a", "b"]})
        assert ws["A1"].value == 1
        assert ws["B1"].value == "a"
        assert ws["A2"].value == 2
        assert ws["B2"].value == "b"

    def test_loop_first_last(self, env):
        ws = make_ws({
            "A1": "{% for x in items %}",
            "A2": "{{ loop.first }}",
            "B2": "{{ loop.last }}",
            "A3": "{% endfor %}",
        })
        SheetRenderer(ws, env).render({"items": ["a", "b", "c"]})
        assert ws["A1"].value is True
        assert ws["B1"].value is False
        assert ws["A2"].value is False
        assert ws["B2"].value is False
        assert ws["A3"].value is False
        assert ws["B3"].value is True


class TestIfBlocks:
    def test_if_true(self, env):
        ws = make_ws({
            "A1": "Header",
            "A2": "{% if show %}",
            "A3": "Visible",
            "A4": "{% endif %}",
            "A5": "Footer",
        })
        SheetRenderer(ws, env).render({"show": True})

        assert ws["A1"].value == "Header"
        assert ws["A2"].value == "Visible"
        assert ws["A3"].value == "Footer"

    def test_if_false(self, env):
        ws = make_ws({
            "A1": "Header",
            "A2": "{% if show %}",
            "A3": "Hidden",
            "A4": "{% endif %}",
            "A5": "Footer",
        })
        SheetRenderer(ws, env).render({"show": False})

        assert ws["A1"].value == "Header"
        assert ws["A2"].value == "Footer"

    def test_if_with_expression_condition(self, env):
        ws = make_ws({
            "A1": "{% if total > 100 %}",
            "A2": "Big order!",
            "A3": "{% endif %}",
        })
        SheetRenderer(ws, env).render({"total": 150})
        assert ws["A1"].value == "Big order!"


class TestNestedBlocks:
    def test_for_inside_for(self, env):
        ws = make_ws({
            "A1": "{% for group in groups %}",
            "A2": "{{ group.name }}",
            "A3": "{% for item in group.items %}",
            "A4": "{{ item }}",
            "A5": "{% endfor %}",
            "A6": "{% endfor %}",
        })
        context = {
            "groups": [
                {"name": "G1", "items": ["a", "b"]},
                {"name": "G2", "items": ["c"]},
            ]
        }
        SheetRenderer(ws, env).render(context)

        # G1 header
        assert ws["A1"].value == "G1"
        # G1 items
        assert ws["A2"].value == "a"
        assert ws["A3"].value == "b"
        # G2 header
        assert ws["A4"].value == "G2"
        # G2 items
        assert ws["A5"].value == "c"

    def test_if_inside_for(self, env):
        ws = make_ws({
            "A1": "{% for item in items %}",
            "A2": "{{ item.name }}",
            "A3": "{% if item.special %}",
            "A4": "SPECIAL",
            "A5": "{% endif %}",
            "A6": "{% endfor %}",
        })
        items = [
            {"name": "A", "special": True},
            {"name": "B", "special": False},
        ]
        SheetRenderer(ws, env).render({"items": items})

        # Item A: name + SPECIAL
        assert ws["A1"].value == "A"
        assert ws["A2"].value == "SPECIAL"
        # Item B: name only (SPECIAL removed)
        assert ws["A3"].value == "B"


class TestSyntaxErrors:
    def test_mismatched_tags(self, env):
        ws = make_ws({
            "A1": "{% for x in items %}",
            "A2": "{{ x }}",
            "A3": "{% endif %}",  # wrong closer
        })
        with pytest.raises(TemplateSyntaxError, match="Mismatched"):
            SheetRenderer(ws, env).render({"items": []})

    def test_unclosed_block(self, env):
        ws = make_ws({
            "A1": "{% for x in items %}",
            "A2": "{{ x }}",
        })
        with pytest.raises(TemplateSyntaxError, match="Unclosed"):
            SheetRenderer(ws, env).render({"items": []})

    def test_bad_iterable(self, env):
        ws = make_ws({
            "A1": "{% for x in nonexistent %}",
            "A2": "{{ x }}",
            "A3": "{% endfor %}",
        })
        with pytest.raises(TemplateRenderError, match="Failed to evaluate"):
            SheetRenderer(ws, env).render({})

    def test_bad_expression(self, env):
        ws = make_ws({"A1": "{{ undefined_var }}"})
        with pytest.raises(TemplateRenderError):
            SheetRenderer(ws, env).render({})
