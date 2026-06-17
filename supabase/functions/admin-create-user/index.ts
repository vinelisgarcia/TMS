import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

type AdminCreateUserBody = {
  email?: string;
  password?: string;
  nombre?: string;
  rol?: string;
  permisosPorModulo?: Record<string, unknown>;
};

function jsonResponse(body: Record<string, unknown>, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, 'Content-Type': 'application/json' },
  });
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: corsHeaders });
  if (req.method !== 'POST') return jsonResponse({ error: 'Método no permitido' }, 405);

  const supabaseUrl = Deno.env.get('SUPABASE_URL') ?? '';
  const anonKey = Deno.env.get('SUPABASE_ANON_KEY') ?? '';
  const serviceRoleKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY') ?? '';
  if (!supabaseUrl || !anonKey || !serviceRoleKey) {
    return jsonResponse({ error: 'Faltan variables de Supabase en la función.' }, 500);
  }

  const authHeader = req.headers.get('Authorization') ?? '';
  const callerClient = createClient(supabaseUrl, anonKey, {
    global: { headers: { Authorization: authHeader } },
  });
  const adminClient = createClient(supabaseUrl, serviceRoleKey);

  const { data: callerData, error: callerError } = await callerClient.auth.getUser();
  if (callerError || !callerData.user) return jsonResponse({ error: 'Sesión inválida.' }, 401);

  let { data: callerProfile, error: profileError } = await adminClient
    .from('tms_user_profiles')
    .select('rol,activo')
    .eq('auth_user_id', callerData.user.id)
    .limit(1)
    .maybeSingle();

  if (profileError) return jsonResponse({ error: profileError.message }, 500);
  if (!callerProfile && callerData.user.email) {
    const byEmail = await adminClient
      .from('tms_user_profiles')
      .select('rol,activo')
      .ilike('email', callerData.user.email)
      .limit(1)
      .maybeSingle();
    if (byEmail.error) return jsonResponse({ error: byEmail.error.message }, 500);
    callerProfile = byEmail.data;
  }
  if (!callerProfile || callerProfile.rol !== 'Admin' || callerProfile.activo === false) {
    return jsonResponse({ error: 'Solo administradores pueden crear usuarios.' }, 403);
  }

  const body = (await req.json()) as AdminCreateUserBody;
  const email = String(body.email ?? '').trim().toLowerCase();
  const password = String(body.password ?? '');
  const nombre = String(body.nombre ?? '').trim();
  const rol = String(body.rol ?? 'Solo lectura').trim() || 'Solo lectura';

  if (!email || !nombre || !password) return jsonResponse({ error: 'Nombre, email y contraseña son obligatorios.' }, 400);
  if (password.length < 6) return jsonResponse({ error: 'La contraseña debe tener al menos 6 caracteres.' }, 400);

  const { data: created, error: createError } = await adminClient.auth.admin.createUser({
    email,
    password,
    email_confirm: true,
    user_metadata: { nombre, rol },
  });

  if (createError || !created.user) {
    return jsonResponse({ error: createError?.message ?? 'No se pudo crear el usuario.' }, 400);
  }

  const accountPerms = {
    ...(body.permisosPorModulo ?? {}),
    __account: {
      passwordChangeRequired: true,
      authNeedsConfirmation: false,
      passwordConfigured: true,
      accountStatus: 'Contraseña temporal',
    },
  };

  const { error: upsertError } = await adminClient.from('tms_user_profiles').upsert({
    auth_user_id: created.user.id,
    nombre,
    email,
    rol,
    permisos_por_modulo: accountPerms,
    activo: true,
  }, { onConflict: 'email' });

  if (upsertError) return jsonResponse({ error: upsertError.message }, 500);

  return jsonResponse({ userId: created.user.id, needsConfirmation: false });
});
