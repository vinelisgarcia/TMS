-- Atomic cloud saves and stricter operational persistence helpers.
-- Apply this after 0003_role_write_policies.sql.

create extension if not exists pgcrypto;

alter table public.tms_route_cost_history
  add column if not exists camion text not null default '';

drop index if exists tms_route_cost_history_fecha_ruta_key;
alter table public.tms_route_cost_history
  drop constraint if exists tms_route_cost_history_fecha_ruta_nombre_key;

create unique index if not exists tms_route_cost_history_fecha_ruta_camion_key
on public.tms_route_cost_history (fecha, ruta_nombre, camion);

create or replace function public.tms_replace_table(p_table_name text, p_rows jsonb default '[]'::jsonb)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  v_rows jsonb := coalesce(p_rows, '[]'::jsonb);
begin
  if jsonb_typeof(v_rows) <> 'array' then
    raise exception 'Payload de % debe ser un arreglo JSON', p_table_name;
  end if;

  if p_table_name = 'tms_user_profiles' and public.tms_current_role() <> 'Admin' then
    raise exception 'Solo administradores pueden reemplazar usuarios';
  elsif p_table_name = 'tms_route_configs' and not (public.tms_current_role() = 'Admin' or public.tms_can_edit('configuracion') or public.tms_can_edit('rutas')) then
    raise exception 'Tu rol no puede guardar rutas';
  elsif p_table_name = 'tms_client_configs' and not (public.tms_current_role() = 'Admin' or public.tms_can_edit('configuracion') or public.tms_can_edit('rutas')) then
    raise exception 'Tu rol no puede guardar configuración de clientes';
  elsif p_table_name = 'tms_order_lines' and not (public.tms_current_role() = 'Admin' or public.tms_can_edit('importar') or public.tms_can_edit('calendario') or public.tms_can_edit('rutas')) then
    raise exception 'Tu rol no puede guardar líneas de pedidos';
  elsif p_table_name = 'tms_route_cost_history' and not (public.tms_current_role() = 'Admin' or public.tms_can_edit('rutas')) then
    raise exception 'Tu rol no puede guardar costos de rutas';
  elsif p_table_name = 'tms_warehouse_plan' and not (public.tms_current_role() = 'Admin' or public.tms_can_edit('importar') or public.tms_can_edit('solicitudesAlmacen')) then
    raise exception 'Tu rol no puede guardar solicitudes planificadas de almacén';
  elsif p_table_name = 'tms_warehouse_control' and not (public.tms_current_role() = 'Admin' or public.tms_can_edit('importar') or public.tms_can_edit('almacen')) then
    raise exception 'Tu rol no puede guardar control de almacén';
  elsif p_table_name = 'tms_warehouse_history' and not (public.tms_current_role() = 'Admin' or public.tms_can_edit('solicitudesAlmacen') or public.tms_can_edit('almacen')) then
    raise exception 'Tu rol no puede guardar histórico de almacén';
  elsif p_table_name not in (
    'tms_user_profiles',
    'tms_route_configs',
    'tms_client_configs',
    'tms_order_lines',
    'tms_route_cost_history',
    'tms_warehouse_plan',
    'tms_warehouse_control',
    'tms_warehouse_history'
  ) then
    raise exception 'Tabla no permitida para reemplazo: %', p_table_name;
  end if;

  if p_table_name = 'tms_user_profiles' then
    delete from public.tms_user_profiles;
    insert into public.tms_user_profiles (auth_user_id, nombre, email, rol, permisos_por_modulo, activo)
    select auth_user_id, nombre, lower(email), rol, coalesce(permisos_por_modulo, '{}'::jsonb), coalesce(activo, true)
    from jsonb_to_recordset(v_rows) as r(auth_user_id uuid, nombre text, email text, rol text, permisos_por_modulo jsonb, activo boolean);
  elsif p_table_name = 'tms_route_configs' then
    delete from public.tms_route_configs;
    insert into public.tms_route_configs (id, nombre, zona, dias_operacion, coste_ruta, activa, observaciones)
    select coalesce(id, gen_random_uuid()), nombre, coalesce(zona, ''), coalesce(dias_operacion, '{}'), coalesce(coste_ruta, 0), coalesce(activa, true), coalesce(observaciones, '')
    from jsonb_to_recordset(v_rows) as r(id uuid, nombre text, zona text, dias_operacion smallint[], coste_ruta numeric, activa boolean, observaciones text);
  elsif p_table_name = 'tms_client_configs' then
    delete from public.tms_client_configs;
    insert into public.tms_client_configs (codigo_cliente, nombre_cliente, ruta_asignada, dias_recepcion, horario_almacen, contacto, observaciones, condiciones_especiales, camion_permitido, configuracion_completa, no_miercoles, no_ultima_semana)
    select codigo_cliente, coalesce(nombre_cliente, ''), coalesce(ruta_asignada, ''), coalesce(dias_recepcion, '{}'), coalesce(horario_almacen, ''), coalesce(contacto, ''), coalesce(observaciones, ''), coalesce(condiciones_especiales, ''), coalesce(camion_permitido, 'CUALQUIERA'), coalesce(configuracion_completa, false), coalesce(no_miercoles, false), coalesce(no_ultima_semana, false)
    from jsonb_to_recordset(v_rows) as r(codigo_cliente text, nombre_cliente text, ruta_asignada text, dias_recepcion smallint[], horario_almacen text, contacto text, observaciones text, condiciones_especiales text, camion_permitido text, configuracion_completa boolean, no_miercoles boolean, no_ultima_semana boolean);
  elsif p_table_name = 'tms_order_lines' then
    delete from public.tms_order_lines;
    insert into public.tms_order_lines (source_batch_id, clave_unica, cliente_id, cliente_nombre, ruta_id, ruta_nombre, fecha_planificada, fecha_control, pedido_cliente, linea_pedido_cliente, articulo, descripcion_articulo, cantidad_solicitada, cantidad_facturada, cantidad_pendiente, estado_planificacion, estado_entrega, origen, camion_asignado, manual_programado, metadata)
    select source_batch_id, clave_unica, cliente_id, coalesce(cliente_nombre, ''), ruta_id, coalesce(ruta_nombre, ''), fecha_planificada, fecha_control, coalesce(pedido_cliente, ''), coalesce(linea_pedido_cliente, ''), coalesce(articulo, ''), coalesce(descripcion_articulo, ''), coalesce(cantidad_solicitada, 0), coalesce(cantidad_facturada, 0), coalesce(cantidad_pendiente, 0), coalesce(estado_planificacion, 'planificado'), coalesce(estado_entrega, 'pendiente'), coalesce(origen, 'planificacion semanal'), coalesce(camion_asignado, 'CAMION 1'), coalesce(manual_programado, false), coalesce(metadata, '{}'::jsonb)
    from jsonb_to_recordset(v_rows) as r(source_batch_id uuid, clave_unica text, cliente_id text, cliente_nombre text, ruta_id uuid, ruta_nombre text, fecha_planificada date, fecha_control date, pedido_cliente text, linea_pedido_cliente text, articulo text, descripcion_articulo text, cantidad_solicitada numeric, cantidad_facturada numeric, cantidad_pendiente numeric, estado_planificacion text, estado_entrega text, origen text, camion_asignado text, manual_programado boolean, metadata jsonb);
  elsif p_table_name = 'tms_route_cost_history' then
    delete from public.tms_route_cost_history;
    insert into public.tms_route_cost_history (fecha, ruta_id, ruta_nombre, camion, coste, fuente_control_diario, pedidos_asociados, cumplimiento, importe_planificado, importe_real, diferencia)
    select fecha, ruta_id, ruta_nombre, coalesce(camion, ''), coalesce(coste, 0), coalesce(fuente_control_diario, ''), coalesce(pedidos_asociados, '[]'::jsonb), coalesce(cumplimiento, 0), coalesce(importe_planificado, 0), coalesce(importe_real, 0), coalesce(diferencia, 0)
    from jsonb_to_recordset(v_rows) as r(fecha date, ruta_id uuid, ruta_nombre text, camion text, coste numeric, fuente_control_diario text, pedidos_asociados jsonb, cumplimiento numeric, importe_planificado numeric, importe_real numeric, diferencia numeric);
  elsif p_table_name = 'tms_warehouse_plan' then
    delete from public.tms_warehouse_plan;
    insert into public.tms_warehouse_plan (source_batch_id, clave_unica, fecha, cliente_id, cliente_nombre, pedido_cliente, solicitud_id, articulo, descripcion_articulo, cantidad_solicitada, estado, coste_solicitud, observaciones, metadata)
    select source_batch_id, clave_unica, fecha, cliente_id, coalesce(cliente_nombre, ''), coalesce(pedido_cliente, ''), coalesce(solicitud_id, ''), coalesce(articulo, ''), coalesce(descripcion_articulo, ''), coalesce(cantidad_solicitada, 0), coalesce(estado, 'Prevista'), coalesce(coste_solicitud, 0), coalesce(observaciones, ''), coalesce(metadata, '{}'::jsonb)
    from jsonb_to_recordset(v_rows) as r(source_batch_id uuid, clave_unica text, fecha date, cliente_id text, cliente_nombre text, pedido_cliente text, solicitud_id text, articulo text, descripcion_articulo text, cantidad_solicitada numeric, estado text, coste_solicitud numeric, observaciones text, metadata jsonb);
  elsif p_table_name = 'tms_warehouse_control' then
    delete from public.tms_warehouse_control;
    insert into public.tms_warehouse_control (source_batch_id, clave_unica, fecha, cliente_id, cliente_nombre, pedido_cliente, solicitud_id, articulo, descripcion_articulo, cantidad_solicitada, cantidad_completada, cantidad_pendiente, estado, coste_solicitud, observaciones, metadata)
    select source_batch_id, clave_unica, fecha, cliente_id, coalesce(cliente_nombre, ''), coalesce(pedido_cliente, ''), coalesce(solicitud_id, ''), coalesce(articulo, ''), coalesce(descripcion_articulo, ''), coalesce(cantidad_solicitada, 0), coalesce(cantidad_completada, 0), coalesce(cantidad_pendiente, 0), coalesce(estado, 'Pendiente'), coalesce(coste_solicitud, 0), coalesce(observaciones, ''), coalesce(metadata, '{}'::jsonb)
    from jsonb_to_recordset(v_rows) as r(source_batch_id uuid, clave_unica text, fecha date, cliente_id text, cliente_nombre text, pedido_cliente text, solicitud_id text, articulo text, descripcion_articulo text, cantidad_solicitada numeric, cantidad_completada numeric, cantidad_pendiente numeric, estado text, coste_solicitud numeric, observaciones text, metadata jsonb);
  elsif p_table_name = 'tms_warehouse_history' then
    delete from public.tms_warehouse_history;
    insert into public.tms_warehouse_history (fecha, cliente_id, cliente_nombre, pedido_cliente, solicitud_id, articulo, descripcion_articulo, cantidad_solicitada, cantidad_completada, cantidad_pendiente, estado, coste_solicitud, observaciones, fuente)
    select fecha, cliente_id, coalesce(cliente_nombre, ''), coalesce(pedido_cliente, ''), coalesce(solicitud_id, ''), coalesce(articulo, ''), coalesce(descripcion_articulo, ''), coalesce(cantidad_solicitada, 0), coalesce(cantidad_completada, 0), coalesce(cantidad_pendiente, 0), coalesce(estado, 'Pendiente'), coalesce(coste_solicitud, 0), coalesce(observaciones, ''), coalesce(fuente, '')
    from jsonb_to_recordset(v_rows) as r(fecha date, cliente_id text, cliente_nombre text, pedido_cliente text, solicitud_id text, articulo text, descripcion_articulo text, cantidad_solicitada numeric, cantidad_completada numeric, cantidad_pendiente numeric, estado text, coste_solicitud numeric, observaciones text, fuente text);
  end if;
end;
$$;

grant execute on function public.tms_replace_table(text, jsonb) to authenticated;
