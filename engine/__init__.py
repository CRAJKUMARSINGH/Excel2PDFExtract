"""Excel -> A4 PDF Print Studio - engine package.

The engine is a thin, well-typed layer *around* the proven `apply.py`
pipeline.  `apply.py` is imported, never modified, so the legacy CLI keeps
working while the engine adds: workbook profiling, a layout registry, a
per-sheet configuration schema and a generalised renderer.

Public entry points::

    from engine import reader, profiler, layouts, renderer
    from engine.schema import SheetConfig, WorkbookConfig

See ``MASTER-PLAN.md`` for the full architecture.
"""

__all__ = ["schema", "reader", "profiler", "layouts", "renderer"]
