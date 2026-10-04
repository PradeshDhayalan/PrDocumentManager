import assert from 'node:assert/strict';
import {createGraphPhotoLoader,createProxyPhotoLoader,createPhotoStore} from '../src/ProfilePhotoLoaders.mjs';
const photo=new Blob(['photo-fixture'],{type:'image/jpeg'});
const controller=new AbortController();
let requests=[];
const graph=createGraphPhotoLoader({resolveEntraUserId:async id=>{assert.equal(id,'dataverse-user');return 'entra-object-id';},getAccessToken:async()=> 'test-token',fetcher:async(url,options)=>{requests.push({url,options});return {ok:true,status:200,blob:async()=>photo};}});
assert.equal(await graph('dataverse-user',controller.signal),photo);
assert.equal(requests[0].url,'https://graph.microsoft.com/v1.0/users/entra-object-id/photo/$value');
assert.equal(requests[0].options.headers.Authorization,'Bearer test-token');
assert.equal(requests[0].options.signal,controller.signal);
for(const status of [401,403,404]){
 const denied=createGraphPhotoLoader({resolveEntraUserId:async()=> 'user',getAccessToken:async()=> 'test-token',fetcher:async()=>({status,ok:false})});
 assert.equal(await denied('dataverse-user'),null);
}
const noAuth=createGraphPhotoLoader({resolveEntraUserId:async()=> 'user',getAccessToken:async()=> null,fetcher:async()=>{throw new Error('Must not fetch without authentication');}});
assert.equal(await noAuth('dataverse-user'),null);
const proxy=createProxyPhotoLoader({origin:'http://localhost:5173',getPhotoUrl:async()=> '/photos/user',fetcher:async(url,options)=>{assert.equal(url,'http://localhost:5173/photos/user');assert.equal(options.credentials,'same-origin');return {status:200,ok:true,blob:async()=>photo};}});
assert.equal(await proxy('user'),photo);
await assert.rejects(createProxyPhotoLoader({origin:'http://localhost:5173',getPhotoUrl:async()=> 'https://elsewhere.example/photo'})('user'));
let calls=0,revoked=[],signal;
const store=createPhotoStore(async(_,s)=>{calls++;signal=s;return photo;},{createObjectURL:()=> 'blob:profile-photo',revokeObjectURL:url=>revoked.push(url)});
assert.deepEqual(await Promise.all([store.get('user'),store.get('user')]),['blob:profile-photo','blob:profile-photo']);
assert.equal(calls,1);
store.dispose();assert.equal(signal.aborted,true);assert.deepEqual(revoked,['blob:profile-photo']);assert.equal(await store.get('user'),null);
let finish;
const pending=createPhotoStore(()=>new Promise(resolve=>finish=resolve),{createObjectURL:()=>{throw new Error('Must not create a URL after unmount');},revokeObjectURL:()=>{}});
const request=pending.get('user');await Promise.resolve();pending.dispose();finish(photo);assert.equal(await request,null);
const failed=createPhotoStore(async()=>{throw new Error('Network unavailable');});assert.equal(await failed.get('user'),null);failed.dispose();
console.log('PASS: Graph identity mapping, bearer request, auth/photo fallback, same-origin proxy, shared caching, cancellation and object URL cleanup.');
