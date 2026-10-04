/** Graph photos require real user identity mapping and an authenticated host. */
export function createGraphPhotoLoader({getAccessToken,resolveEntraUserId,fetcher=fetch}) {
 return async (dataverseUserId,signal)=>{
  const entraUserId=await resolveEntraUserId(dataverseUserId,signal);
  if(!entraUserId||signal?.aborted)return null;
  const token=await getAccessToken(signal);
  if(!token||signal?.aborted)return null;
  const response=await fetcher(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(entraUserId)}/photo/$value`,{
   headers:{Authorization:`Bearer ${token}`},credentials:'omit',signal,
  });
  return readPhoto(response);
 };
}
/** Recommended for PCF: authenticated same-origin server maps IDs and calls Graph. */
export function createProxyPhotoLoader({getPhotoUrl,origin,fetcher=fetch}) {
 return async (dataverseUserId,signal)=>{
  const value=await getPhotoUrl(dataverseUserId,signal);
  if(!value||signal?.aborted)return null;
  const url=new URL(value,origin);
  if(url.origin!==origin)throw new Error('Profile photo proxy must be same-origin');
  return readPhoto(await fetcher(url.href,{credentials:'same-origin',signal}));
 };
}
async function readPhoto(response){
 if([401,403,404].includes(response.status))return null;
 if(!response.ok)throw new Error('Profile photo request failed');
 const blob=await response.blob();
 return blob.type.startsWith('image/')&&blob.size>0?blob:null;
}
/** One request per user per host lifetime; object URLs are never persisted. */
export function createPhotoStore(loader,{createObjectURL=URL.createObjectURL,revokeObjectURL=URL.revokeObjectURL}={}) {
 const cache=new Map(),urls=new Set(),controller=new AbortController();
 let disposed=false;
 return {
  get(userId){
   if(disposed||!userId||!loader)return Promise.resolve(null);
   if(!cache.has(userId))cache.set(userId,Promise.resolve().then(()=>loader(userId,controller.signal)).then(blob=>{
    if(disposed||!blob)return null;
    const url=createObjectURL(blob);urls.add(url);return url;
   }).catch(()=>null));
   return cache.get(userId);
  },
  dispose(){disposed=true;controller.abort();urls.forEach(url=>revokeObjectURL(url));urls.clear();cache.clear();},
 };
}
