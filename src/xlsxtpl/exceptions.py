class XlsxTemplateError(Exception):
    """Base exception for xlsxtpl."""


class TemplateRenderError(XlsxTemplateError):
    """Raised when rendering a template fails."""


class TemplateSyntaxError(XlsxTemplateError):
    """Raised when template block structure is invalid (mismatched tags, etc.)."""
