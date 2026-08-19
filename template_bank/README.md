# Template Bank

This directory is the default store for archived Style Specs. It is
populated at runtime, not checked in.

```
index.json                    one row per entry
entries/<entry_id>/
  style_spec.json             the archived Style Spec
  master.pptx                 the submitted master, archived verbatim
  revisions/rev-<n>.json      prior versions, kept whenever a spec is refined
  assets/                     logo and background images
```

**This is client data.** Entries hold submitted masters and extracted brand
specs, so `entries/` and `index.json` are gitignored. Back the directory up
as data, on storage approved for client material, and point the tool at a
shared location with `--bank` when more than one person needs the same bank.

Nothing here needs to be created by hand. Run `extract` or `pipeline` with a
`--client` and the entry appears.
