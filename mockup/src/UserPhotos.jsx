import React,{createContext,useContext,useEffect,useMemo,useState} from 'react';
import {createPhotoStore} from './ProfilePhotoLoaders.mjs';
const PhotoContext=createContext(null);
export function UserPhotoProvider({loader,children}){
 const store=useMemo(()=>createPhotoStore(loader),[loader]);
 useEffect(()=>()=>store.dispose(),[store]);
 return <PhotoContext.Provider value={store}>{children}</PhotoContext.Provider>;
}
export function useUserPhoto(userId){
 const store=useContext(PhotoContext),[photo,setPhoto]=useState(null);
 useEffect(()=>{
  let active=true;setPhoto(null);
  if(store&&userId)store.get(userId).then(url=>{if(active)setPhoto(url);});
  return ()=>{active=false;};
 },[store,userId]);
 return photo;
}
/** Sample IDs are harness-only; production rows must carry Dataverse user IDs. */
export const sampleUserIds={
 'Michael Bose':'sample-michael','Alex Wilber':'sample-alex','Irvin Sayers':'sample-irvin',
 'MOD Administrator':'sample-admin','Adele Vance':'sample-adele','Megan Bowen':'sample-megan',
};
