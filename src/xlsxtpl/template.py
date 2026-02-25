from __future__ import annotations

from pathlib import Path
from typing import Any

import jinja2
import openpyxl

from .jinja_env import create_jinja_env
from .renderer import SheetRenderer


class XlsxTemplate:
    """Load an xlsx template, render it with a context dict, and save the output."""

    def __init__(self, path: str | Path) -> None:
        self.workbook = openpyxl.load_workbook(str(path))
        self._jinja_env: jinja2.Environment | None = None

    @property
    def jinja_env(self) -> jinja2.Environment:
        """Jinja2 environment used for rendering. Created on first access."""
        if self._jinja_env is None:
            self._jinja_env = create_jinja_env()
        return self._jinja_env

    @jinja_env.setter
    def jinja_env(self, env: jinja2.Environment) -> None:
        self._jinja_env = env

    def render(self, context: dict[str, Any]) -> None:
        """Render all worksheets in the workbook with the given context."""
        env = self.jinja_env
        for ws in self.workbook.worksheets:
            renderer = SheetRenderer(ws, env)
            renderer.render(context)

    def save(self, path: str | Path) -> None:
        """Save the rendered workbook to the given path."""
        self.workbook.save(str(path))
