begin;

create extension if not exists pgcrypto;

create table if not exists public.tms_user_profiles (
  id uuid primary key default gen_random_uuid(),
  auth_user_id uuid null references auth.users(id) on delete set null,
  nombre text not null,
  email text not null,
  rol text not null default 'Solo lectura',
  permisos_por_modulo jsonb not null default '{}'::jsonb,
  activo boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (email)
);

create table if not exists public.tms_route_configs (
  id uuid primary key default gen_random_uuid(),
  nombre text not null,
  zona text not null default '',
  dias_operacion smallint[] not null default '{}',
  coste_ruta numeric(14,2) not null default 0,
  activa boolean not null default true,
  observaciones text not null default '',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (nombre)
);

create table if not exists public.tms_client_configs (
  id uuid primary key default gen_random_uuid(),
  codigo_cliente text not null,
  nombre_cliente text not null default '',
  ruta_asignada text not null default '',
  dias_recepcion smallint[] not null default '{}',
  horario_almacen text not null default '',
  contacto text not null default '',
  observaciones text not null default '',
  condiciones_especiales text not null default '',
  camion_permitido text not null default 'CUALQUIERA',
  configuracion_completa boolean not null default false,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (codigo_cliente)
);

create table if not exists public.tms_settings (
  id uuid primary key default gen_random_uuid(),
  clave text not null unique,
  valor jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.tms_import_batches (
  id uuid primary key default gen_random_uuid(),
  tipo text not null,
  archivo text not null default '',
  fecha_operativa date null,
  metadata jsonb not null default '{}'::jsonb,
  created_by uuid null references auth.users(id) on delete set null,
  created_at timestamptz not null default now()
);

create table if not exists public.tms_order_lines (
  id uuid primary key default gen_random_uuid(),
  source_batch_id uuid null references public.tms_import_batches(id) on delete set null,
  clave_unica text not null,
  cliente_id text not null,
  cliente_nombre text not null default '',
  ruta_id uuid null references public.tms_route_configs(id) on delete set null,
  ruta_nombre text not null default '',
  fecha_planificada date null,
  fecha_control date null,
  pedido_cliente text not null default '',
  linea_pedido_cliente text not null default '',
  articulo text not null default '',
  descripcion_articulo text not null default '',
  cantidad_solicitada numeric(14,2) not null default 0,
  cantidad_facturada numeric(14,2) not null default 0,
  cantidad_pendiente numeric(14,2) not null default 0,
  estado_planificacion text not null default 'planificado',
  estado_entrega text not null default 'pendiente',
  origen text not null default 'planificacion semanal',
  camion_asignado text not null default 'CAMION 1',
  manual_programado boolean not null default false,
  metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.tms_route_cost_history (
  id uuid primary key default gen_random_uuid(),
  fecha date not null,
  ruta_id uuid null references public.tms_route_configs(id) on delete set null,
  ruta_nombre text not null,
  coste numeric(14,2) not null default 0,
  fuente_control_diario text not null default '',
  pedidos_asociados jsonb not null default '[]'::jsonb,
  cumplimiento numeric(7,2) not null default 0,
  importe_planificado numeric(14,2) not null default 0,
  importe_real numeric(14,2) not null default 0,
  diferencia numeric(14,2) not null default 0,
  created_at timestamptz not null default now(),
  unique (fecha, ruta_nombre)
);

create table if not exists public.tms_warehouse_plan (
  id uuid primary key default gen_random_uuid(),
  source_batch_id uuid null references public.tms_import_batches(id) on delete set null,
  clave_unica text not null unique,
  fecha date null,
  cliente_id text not null,
  cliente_nombre text not null default '',
  pedido_cliente text not null default '',
  solicitud_id text not null default '',
  articulo text not null default '',
  descripcion_articulo text not null default '',
  cantidad_solicitada numeric(14,2) not null default 0,
  estado text not null default 'Prevista',
  coste_solicitud numeric(14,2) not null default 0,
  observaciones text not null default '',
  metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.tms_warehouse_control (
  id uuid primary key default gen_random_uuid(),
  source_batch_id uuid null references public.tms_import_batches(id) on delete set null,
  clave_unica text not null unique,
  fecha date null,
  cliente_id text not null,
  cliente_nombre text not null default '',
  pedido_cliente text not null default '',
  solicitud_id text not null default '',
  articulo text not null default '',
  descripcion_articulo text not null default '',
  cantidad_solicitada numeric(14,2) not null default 0,
  cantidad_completada numeric(14,2) not null default 0,
  cantidad_pendiente numeric(14,2) not null default 0,
  estado text not null default 'Pendiente',
  coste_solicitud numeric(14,2) not null default 0,
  observaciones text not null default '',
  metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.tms_warehouse_history (
  id uuid primary key default gen_random_uuid(),
  fecha date not null,
  cliente_id text not null,
  cliente_nombre text not null default '',
  pedido_cliente text not null default '',
  solicitud_id text not null default '',
  articulo text not null default '',
  descripcion_articulo text not null default '',
  cantidad_solicitada numeric(14,2) not null default 0,
  cantidad_completada numeric(14,2) not null default 0,
  cantidad_pendiente numeric(14,2) not null default 0,
  estado text not null default 'Pendiente',
  coste_solicitud numeric(14,2) not null default 0,
  observaciones text not null default '',
  fuente text not null default '',
  created_at timestamptz not null default now()
);

create or replace function public.tms_set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists tms_user_profiles_updated_at on public.tms_user_profiles;
create trigger tms_user_profiles_updated_at
before update on public.tms_user_profiles
for each row execute procedure public.tms_set_updated_at();

drop trigger if exists tms_route_configs_updated_at on public.tms_route_configs;
create trigger tms_route_configs_updated_at
before update on public.tms_route_configs
for each row execute procedure public.tms_set_updated_at();

drop trigger if exists tms_client_configs_updated_at on public.tms_client_configs;
create trigger tms_client_configs_updated_at
before update on public.tms_client_configs
for each row execute procedure public.tms_set_updated_at();

drop trigger if exists tms_settings_updated_at on public.tms_settings;
create trigger tms_settings_updated_at
before update on public.tms_settings
for each row execute procedure public.tms_set_updated_at();

drop trigger if exists tms_order_lines_updated_at on public.tms_order_lines;
create trigger tms_order_lines_updated_at
before update on public.tms_order_lines
for each row execute procedure public.tms_set_updated_at();

drop trigger if exists tms_warehouse_plan_updated_at on public.tms_warehouse_plan;
create trigger tms_warehouse_plan_updated_at
before update on public.tms_warehouse_plan
for each row execute procedure public.tms_set_updated_at();

drop trigger if exists tms_warehouse_control_updated_at on public.tms_warehouse_control;
create trigger tms_warehouse_control_updated_at
before update on public.tms_warehouse_control
for each row execute procedure public.tms_set_updated_at();

create unique index if not exists tms_order_lines_unique_key
on public.tms_order_lines (clave_unica, coalesce(fecha_planificada, fecha_control));

alter table public.tms_user_profiles enable row level security;
alter table public.tms_route_configs enable row level security;
alter table public.tms_client_configs enable row level security;
alter table public.tms_settings enable row level security;
alter table public.tms_import_batches enable row level security;
alter table public.tms_order_lines enable row level security;
alter table public.tms_route_cost_history enable row level security;
alter table public.tms_warehouse_plan enable row level security;
alter table public.tms_warehouse_control enable row level security;
alter table public.tms_warehouse_history enable row level security;

drop policy if exists "authenticated users can read tms_user_profiles" on public.tms_user_profiles;
create policy "authenticated users can read tms_user_profiles"
on public.tms_user_profiles for select
to authenticated
using (true);

drop policy if exists "authenticated users can write tms_user_profiles" on public.tms_user_profiles;
create policy "authenticated users can write tms_user_profiles"
on public.tms_user_profiles for all
to authenticated
using (true)
with check (true);

drop policy if exists "authenticated users can read tms_route_configs" on public.tms_route_configs;
create policy "authenticated users can read tms_route_configs"
on public.tms_route_configs for select
to authenticated
using (true);

drop policy if exists "authenticated users can write tms_route_configs" on public.tms_route_configs;
create policy "authenticated users can write tms_route_configs"
on public.tms_route_configs for all
to authenticated
using (true)
with check (true);

drop policy if exists "authenticated users can read tms_client_configs" on public.tms_client_configs;
create policy "authenticated users can read tms_client_configs"
on public.tms_client_configs for select
to authenticated
using (true);

drop policy if exists "authenticated users can write tms_client_configs" on public.tms_client_configs;
create policy "authenticated users can write tms_client_configs"
on public.tms_client_configs for all
to authenticated
using (true)
with check (true);

drop policy if exists "authenticated users can read tms_settings" on public.tms_settings;
create policy "authenticated users can read tms_settings"
on public.tms_settings for select
to authenticated
using (true);

drop policy if exists "authenticated users can write tms_settings" on public.tms_settings;
create policy "authenticated users can write tms_settings"
on public.tms_settings for all
to authenticated
using (true)
with check (true);

drop policy if exists "authenticated users can read tms_import_batches" on public.tms_import_batches;
create policy "authenticated users can read tms_import_batches"
on public.tms_import_batches for select
to authenticated
using (true);

drop policy if exists "authenticated users can write tms_import_batches" on public.tms_import_batches;
create policy "authenticated users can write tms_import_batches"
on public.tms_import_batches for all
to authenticated
using (true)
with check (true);

drop policy if exists "authenticated users can read tms_order_lines" on public.tms_order_lines;
create policy "authenticated users can read tms_order_lines"
on public.tms_order_lines for select
to authenticated
using (true);

drop policy if exists "authenticated users can write tms_order_lines" on public.tms_order_lines;
create policy "authenticated users can write tms_order_lines"
on public.tms_order_lines for all
to authenticated
using (true)
with check (true);

drop policy if exists "authenticated users can read tms_route_cost_history" on public.tms_route_cost_history;
create policy "authenticated users can read tms_route_cost_history"
on public.tms_route_cost_history for select
to authenticated
using (true);

drop policy if exists "authenticated users can write tms_route_cost_history" on public.tms_route_cost_history;
create policy "authenticated users can write tms_route_cost_history"
on public.tms_route_cost_history for all
to authenticated
using (true)
with check (true);

drop policy if exists "authenticated users can read tms_warehouse_plan" on public.tms_warehouse_plan;
create policy "authenticated users can read tms_warehouse_plan"
on public.tms_warehouse_plan for select
to authenticated
using (true);

drop policy if exists "authenticated users can write tms_warehouse_plan" on public.tms_warehouse_plan;
create policy "authenticated users can write tms_warehouse_plan"
on public.tms_warehouse_plan for all
to authenticated
using (true)
with check (true);

drop policy if exists "authenticated users can read tms_warehouse_control" on public.tms_warehouse_control;
create policy "authenticated users can read tms_warehouse_control"
on public.tms_warehouse_control for select
to authenticated
using (true);

drop policy if exists "authenticated users can write tms_warehouse_control" on public.tms_warehouse_control;
create policy "authenticated users can write tms_warehouse_control"
on public.tms_warehouse_control for all
to authenticated
using (true)
with check (true);

drop policy if exists "authenticated users can read tms_warehouse_history" on public.tms_warehouse_history;
create policy "authenticated users can read tms_warehouse_history"
on public.tms_warehouse_history for select
to authenticated
using (true);

drop policy if exists "authenticated users can write tms_warehouse_history" on public.tms_warehouse_history;
create policy "authenticated users can write tms_warehouse_history"
on public.tms_warehouse_history for all
to authenticated
using (true)
with check (true);

insert into public.tms_settings (clave, valor)
values ('warehouse', jsonb_build_object('costoSolicitud', 0))
on conflict (clave) do nothing;

insert into public.tms_user_profiles (nombre, email, rol, permisos_por_modulo)
values (
  'Vinelis Garcia',
  'vinelis.garcia@tms-alvarez.local',
  'Admin',
  jsonb_build_object(
    'importar', jsonb_build_object('ver', true, 'editar', true, 'importar', true),
    'dashboard', jsonb_build_object('ver', true, 'editar', true),
    'calendario', jsonb_build_object('ver', true, 'editar', true),
    'rutas', jsonb_build_object('ver', true, 'editar', true),
    'reportes', jsonb_build_object('ver', true, 'editar', true),
    'comercial', jsonb_build_object('ver', true, 'editar', true),
    'configuracion', jsonb_build_object('ver', true, 'editar', true, 'importar', true),
    'solicitudesAlmacen', jsonb_build_object('ver', true, 'editar', true),
    'almacen', jsonb_build_object('ver', true, 'editar', true)
  )
)
on conflict (email) do update
set rol = excluded.rol,
    permisos_por_modulo = excluded.permisos_por_modulo,
    updated_at = now();

commit;
