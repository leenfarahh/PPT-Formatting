-- PPTX Formatting Tool - Supabase schema.
--
-- Apply with the Supabase SQL editor, or:
--     supabase db execute --file supabase/schema.sql
--
-- This holds client material: submitted master decks, extracted brand
-- specs, and finished decks. Row-level security is therefore enabled on
-- every table and all three storage buckets are private. The API server
-- talks to Supabase with the service-role key, which bypasses RLS by
-- design - so that key must stay server-side and never reach a browser.
--
-- No policies are granted to `anon` or `authenticated` here. That is
-- deliberate: with RLS enabled and no policy, those roles can read
-- nothing. Add policies scoped to your own auth model before exposing
-- this to anything other than the server.

create extension if not exists "pgcrypto";

-- --------------------------------------------------------------------
-- Template Bank
-- --------------------------------------------------------------------

create table if not exists public.bank_entries (
    entry_id      text primary key,
    client        text,
    project       text,
    source_name   text,
    spec_version  text        not null default '1.0',
    revision      integer     not null default 1,
    slide_width   bigint      not null default 0,
    slide_height  bigint      not null default 0,
    -- Canonical archetypes this entry can supply, denormalized from the
    -- spec so gap-filling can skip irrelevant entries without reading
    -- every spec document.
    archetypes    text[]      not null default '{}',
    style_spec    jsonb       not null,
    -- Object path in the `masters` bucket. The master is archived so a
    -- repeat deck can skip extraction and still build on the real file.
    master_object text,
    created_at    timestamptz not null default now(),
    updated_at    timestamptz not null default now()
);

comment on table public.bank_entries is
    'One archived master per row: its Style Spec and the layouts it can supply.';

create index if not exists bank_entries_client_idx
    on public.bank_entries (client);
create index if not exists bank_entries_updated_idx
    on public.bank_entries (updated_at desc);
-- Gap-filling asks "which entries offer this archetype?", which is a
-- containment test over the array.
create index if not exists bank_entries_archetypes_idx
    on public.bank_entries using gin (archetypes);

-- Prior versions, written whenever a designer's correction is folded back
-- in. Keeps the feedback loop auditable and reversible.
create table if not exists public.bank_revisions (
    id         bigint generated always as identity primary key,
    entry_id   text        not null references public.bank_entries (entry_id)
                   on delete cascade,
    revision   integer     not null,
    style_spec jsonb       not null,
    note       text,
    created_at timestamptz not null default now(),
    unique (entry_id, revision)
);

create index if not exists bank_revisions_entry_idx
    on public.bank_revisions (entry_id, revision desc);

-- --------------------------------------------------------------------
-- Jobs
-- --------------------------------------------------------------------

create table if not exists public.jobs (
    job_id           text primary key,
    client           text,
    project          text,
    entry_id         text references public.bank_entries (entry_id) on delete set null,
    master_filename  text,
    content_filename text,
    output_name      text        not null default 'formatted_deck.pptx',
    -- Object path in the `outputs` bucket.
    output_object    text,
    status           text        not null default 'running'
                         check (status in ('running', 'complete', 'failed')),
    error            text,
    slides_processed integer     not null default 0,
    qa_flag_count    integer     not null default 0,
    warning_count    integer     not null default 0,
    stage_1_skipped  boolean     not null default false,
    -- The full per-slide report: routing, the evidence behind it, and QA
    -- findings. Kept so a disputed slide can be reviewed after the fact.
    report           jsonb,
    created_at       timestamptz not null default now(),
    completed_at     timestamptz
);

create index if not exists jobs_created_idx on public.jobs (created_at desc);
create index if not exists jobs_client_idx  on public.jobs (client);

-- Keep bank_entries.updated_at honest even for direct SQL edits.
create or replace function public.touch_updated_at()
returns trigger
language plpgsql
as $$
begin
    new.updated_at = now();
    return new;
end;
$$;

drop trigger if exists bank_entries_touch on public.bank_entries;
create trigger bank_entries_touch
    before update on public.bank_entries
    for each row execute function public.touch_updated_at();

-- --------------------------------------------------------------------
-- Row-level security
-- --------------------------------------------------------------------
-- Enabled with no policies: anon and authenticated can read nothing.
-- Only the service-role key (used by the API server) gets through.

alter table public.bank_entries   enable row level security;
alter table public.bank_revisions enable row level security;
alter table public.jobs           enable row level security;

-- --------------------------------------------------------------------
-- Storage buckets
-- --------------------------------------------------------------------
-- All private. Files are served through the API, which is where access
-- control belongs, rather than by public URL.

insert into storage.buckets (id, name, public)
values ('masters', 'masters', false),
       ('assets',  'assets',  false),
       ('outputs', 'outputs', false)
on conflict (id) do nothing;
