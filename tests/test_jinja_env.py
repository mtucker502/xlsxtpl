from datetime import date, datetime

import pytest

from xlsxtpl.jinja_env import create_jinja_env


class TestCreateJinjaEnv:
    def test_returns_environment(self):
        env = create_jinja_env()
        assert env is not None

    def test_no_autoescape(self):
        env = create_jinja_env()
        assert env.autoescape is False

    def test_strict_undefined(self):
        import jinja2
        env = create_jinja_env()
        assert env.undefined is jinja2.StrictUndefined


class TestDateFilter:
    def test_date_default_format(self):
        env = create_jinja_env()
        tpl = env.from_string("{{ d | date }}")
        result = tpl.render(d=date(2024, 3, 15))
        assert result == "2024-03-15"

    def test_date_custom_format(self):
        env = create_jinja_env()
        tpl = env.from_string("{{ d | date('%m/%d/%Y') }}")
        result = tpl.render(d=date(2024, 3, 15))
        assert result == "03/15/2024"

    def test_datetime_format(self):
        env = create_jinja_env()
        tpl = env.from_string("{{ d | date('%Y-%m-%d %H:%M') }}")
        result = tpl.render(d=datetime(2024, 3, 15, 10, 30))
        assert result == "2024-03-15 10:30"

    def test_non_date_passthrough(self):
        env = create_jinja_env()
        tpl = env.from_string("{{ d | date }}")
        result = tpl.render(d="not a date")
        assert result == "not a date"


class TestNumberFormatFilter:
    def test_default_format(self):
        env = create_jinja_env()
        tpl = env.from_string("{{ n | number_format }}")
        result = tpl.render(n=1234567.891)
        assert result == "1,234,567.89"

    def test_custom_decimals(self):
        env = create_jinja_env()
        tpl = env.from_string("{{ n | number_format(0) }}")
        result = tpl.render(n=1234.7)
        assert result == "1,235"

    def test_non_numeric_passthrough(self):
        env = create_jinja_env()
        tpl = env.from_string("{{ n | number_format }}")
        result = tpl.render(n="text")
        assert result == "text"


class TestBuiltinFilters:
    def test_upper(self):
        env = create_jinja_env()
        tpl = env.from_string("{{ name | upper }}")
        assert tpl.render(name="alice") == "ALICE"

    def test_lower(self):
        env = create_jinja_env()
        tpl = env.from_string("{{ name | lower }}")
        assert tpl.render(name="ALICE") == "alice"

    def test_title(self):
        env = create_jinja_env()
        tpl = env.from_string("{{ name | title }}")
        assert tpl.render(name="hello world") == "Hello World"

    def test_round(self):
        env = create_jinja_env()
        tpl = env.from_string("{{ n | round(2) }}")
        assert tpl.render(n=3.14159) == "3.14"

    def test_int(self):
        env = create_jinja_env()
        tpl = env.from_string("{{ n | int }}")
        assert tpl.render(n=3.7) == "3"


class TestCustomFilterRegistration:
    def test_add_custom_filter(self):
        env = create_jinja_env()
        env.filters["currency"] = lambda v: f"${v:,.2f}"
        tpl = env.from_string("{{ price | currency }}")
        assert tpl.render(price=1234.5) == "$1,234.50"

    def test_override_builtin_filter(self):
        env = create_jinja_env()
        env.filters["upper"] = lambda v: v.upper() + "!!!"
        tpl = env.from_string("{{ name | upper }}")
        assert tpl.render(name="hi") == "HI!!!"
