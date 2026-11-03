(function(){
 'use strict';
 const state=window.TMSStandaloneAccess={allowed:false,userId:null,ready:null};
 let client,checking=false,generation=0;
 function display(allowed,message){
  const changed=state.allowed!==allowed;state.allowed=allowed;document.documentElement.dataset.moduleAccess=allowed?'allowed':'denied';
  const label=document.getElementById('moduleAccessMessage');if(label)label.textContent=message;
  if(changed)window.dispatchEvent(new CustomEvent('tms:module-access',{detail:{allowed}}));
 }
 async function check(){
  if(checking)return state.allowed;checking=true;const epoch=generation;
  try{
   const {data,error}=await client.auth.getUser();if(error||!data.user){display(false,'Inicia sesión en el TMS para abrir Auditoría Jamaica.');return false;}
   if(state.userId&&state.userId!==data.user.id)display(false,'Comprobando la nueva cuenta…');state.userId=data.user.id;
   const profile=await client.from('tms_user_profiles').select('rol,activo').eq('auth_user_id',data.user.id).maybeSingle();
   if(profile.error||!profile.data||profile.data.activo!==true){display(false,'La cuenta no tiene un perfil activo en el TMS.');return false;}
   let allowed=profile.data.rol==='Admin';
   if(!allowed){const settings=await client.from('tms_settings').select('valor').eq('clave','role_permissions').maybeSingle();if(settings.error)throw settings.error;allowed=settings.data?.valor?.[profile.data.rol]?.auditoriaJamaica?.ver===true;}
   if(epoch!==generation)return false;display(allowed,allowed?'':'Tu rol no tiene acceso a Auditoría Jamaica.');return allowed;
  }catch(error){display(false,'No se pudo verificar el acceso. Revisa la conexión y vuelve a intentarlo.');return false;}finally{checking=false;}
 }
 state.ready=new Promise(resolve=>document.addEventListener('DOMContentLoaded',async()=>{
  if(!window.supabase){display(false,'No se pudo cargar el servicio de acceso. Vuelve a cargar esta página.');resolve(false);return;}
  client=window.supabase.createClient(window.TMSStandaloneConfig.url,window.TMSStandaloneConfig.key);
  client.auth.onAuthStateChange(event=>{if(event==='SIGNED_OUT'){generation++;display(false,'La sesión terminó. Inicia sesión en el TMS.');}else setTimeout(check,0);});
  document.getElementById('moduleAccessRetry').onclick=check;
  document.addEventListener('visibilitychange',()=>{if(!document.hidden)check();});
  window.addEventListener('focus',check);setInterval(check,15000);resolve(await check());
 }));
})();
