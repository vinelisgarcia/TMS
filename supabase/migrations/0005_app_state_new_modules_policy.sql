-- Keep app_state cloud saves aligned with modules added after the original RLS policies.
-- Apply after 0004_atomic_replace_and_route_history.sql.

drop policy if exists "admins and allowed roles can write settings" on public.tms_settings;
create policy "admins and allowed roles can write settings"
on public.tms_settings for all
to authenticated
using (
  public.tms_current_role() = 'Admin'
  or (clave = 'warehouse' and public.tms_can_edit('almacen'))
  or (
    clave = 'app_state'
    and (
      public.tms_can_edit('importar')
      or public.tms_can_edit('solicitudesAlmacen')
      or public.tms_can_edit('almacen')
      or public.tms_can_edit('configuracion')
      or public.tms_can_edit('rutas')
      or public.tms_can_edit('calendario')
      or public.tms_can_edit('materialesTransito')
      or public.tms_can_edit('prioridades')
      or public.tms_can_edit('incidenciasDespacho')
    )
  )
)
with check (
  public.tms_current_role() = 'Admin'
  or (clave = 'warehouse' and public.tms_can_edit('almacen'))
  or (
    clave = 'app_state'
    and (
      public.tms_can_edit('importar')
      or public.tms_can_edit('solicitudesAlmacen')
      or public.tms_can_edit('almacen')
      or public.tms_can_edit('configuracion')
      or public.tms_can_edit('rutas')
      or public.tms_can_edit('calendario')
      or public.tms_can_edit('materialesTransito')
      or public.tms_can_edit('prioridades')
      or public.tms_can_edit('incidenciasDespacho')
    )
  )
);
