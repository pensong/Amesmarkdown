from __future__ import annotations


def convert_docx(*args, **kwargs):
    from .docx_converter import convert_docx as _convert_docx

    return _convert_docx(*args, **kwargs)


def convert_markdown(*args, **kwargs):
    from .markdown_export import convert_markdown as _convert_markdown

    return _convert_markdown(*args, **kwargs)


def convert_pdf(*args, **kwargs):
    from .pdf_converter import convert_pdf as _convert_pdf

    return _convert_pdf(*args, **kwargs)


def convert_pptx(*args, **kwargs):
    from .pptx_converter import convert_pptx as _convert_pptx

    return _convert_pptx(*args, **kwargs)


def convert_xlsx(*args, **kwargs):
    from .xlsx_converter import convert_xlsx as _convert_xlsx

    return _convert_xlsx(*args, **kwargs)


__all__ = ["convert_docx", "convert_markdown", "convert_pdf", "convert_pptx", "convert_xlsx"]
