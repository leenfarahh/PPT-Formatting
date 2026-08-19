"""
pptx_formatter
==============

Formats a rough content deck to match a designer's submitted master slide.

    Stage 1  extraction        master .pptx -> Style Spec, archived in the bank
    Stage 2  layout_generator  Style Spec   -> a full set of restyled layouts
    Stage 3  classifier +      rough deck   -> each slide rebuilt on the
             slide_copy +                      layout matching what it is
             formatting

Supporting modules:

    style_spec      the versioned JSON contract every stage reads and writes
    archetypes      the shared layout vocabulary and its similarity scoring
    bank            archive of Style Specs, used to fill layout gaps
    supabase_bank   the same archive backed by Supabase
    jobs            one record per formatting run, with its report
    config          backend selection from the environment
    layout_builder  writes slide-layout parts (python-pptx cannot add layouts)
    part_copy       clones a part and its dependencies, e.g. charts
    rtl             Arabic direction and complex-script font handling
    qa              static validation of the finished deck

Storage is chosen by environment - the local filesystem or Supabase - and
nothing in the pipeline knows which is in use. Jobs run synchronously, in
process; `api/` serves the UI and the JSON API and has no authentication.
See README.md.
"""

__version__ = "1.1.0"
