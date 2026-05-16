alter table public.shared_templates
  add column if not exists visibility text not null default 'shared',
  add column if not exists owner_id uuid null;

do $$
begin
  if not exists (
    select 1 from pg_constraint
    where conname = 'shared_templates_visibility_check'
  ) then
    alter table public.shared_templates
      add constraint shared_templates_visibility_check
      check (visibility in ('shared', 'private'));
  end if;
end$$;

create index if not exists idx_shared_templates_visibility on public.shared_templates (visibility);
create index if not exists idx_shared_templates_owner_id on public.shared_templates (owner_id);

