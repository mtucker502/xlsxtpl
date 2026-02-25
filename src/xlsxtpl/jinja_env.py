from __future__ import annotations

from datetime import date, datetime

import jinja2


def _filter_date(value: date | datetime, fmt: str = "%Y-%m-%d") -> str:
    if isinstance(value, (date, datetime)):
        return value.strftime(fmt)
    return str(value)


def _filter_number_format(value: float | int, decimals: int = 2, thousands: str = ",") -> str:
    if not isinstance(value, (int, float)):
        return str(value)
    formatted = f"{value:,.{decimals}f}" if thousands == "," else f"{value:.{decimals}f}"
    return formatted


def create_jinja_env(**kwargs) -> jinja2.Environment:
    """Create a Jinja2 Environment configured for xlsx templating."""
    env = jinja2.Environment(
        autoescape=False,
        keep_trailing_newline=False,
        undefined=jinja2.StrictUndefined,
        **kwargs,
    )

    env.filters["date"] = _filter_date
    env.filters["number_format"] = _filter_number_format

    return env
