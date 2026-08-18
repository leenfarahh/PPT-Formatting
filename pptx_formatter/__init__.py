"""
pptx_formatter
==============

Local reference implementation of Phase 1 (MVP) and Phase 2 (Automatic Master
Layout Editing) from the "Automated PowerPoint Formatting Tool" technical plan.

Scope (intentionally limited to match Phases 1-2 of the plan):
    Stage 1 - Style Extraction   -> pptx_formatter.extraction
    Stage 2 - Master Layout Gen  -> pptx_formatter.layout_generator
    Stage 3 - Detailed Formatting (typography, color, grid only;
              charts/tables/icons are Phase 3 and are NOT implemented here)
                                  -> pptx_formatter.formatting

Everything here runs synchronously, in-process, against the local
filesystem. That is deliberate: this package is a local dev/testing harness,
not the production architecture described in the technical plan (no job
queue, no object storage, no Aspose.Slides fallback, no multi-tenant
isolation). See README.md for what's simplified and why.
"""

__version__ = "0.1.0-phase1-2"
