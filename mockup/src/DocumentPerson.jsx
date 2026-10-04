import React from 'react';
import {useUserPhoto,sampleUserIds} from './UserPhotos';
import {Persona,makeStyles,tokens} from '@fluentui/react-components';
const useStyles=makeStyles({person:{minWidth:0,'& .fui-Persona__primaryText':{fontSize:'13px',fontWeight:tokens.fontWeightRegular,color:tokens.colorNeutralForeground2,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}},activity:{marginBottom:'14px'}});
export default function DocumentPerson({name,userId,secondaryText,activity=false}){
 const s=useStyles();
 const photo=useUserPhoto(userId||sampleUserIds[name]);
 return <Persona className={activity?s.activity:s.person} name={name} size={activity?'medium':'extra-small'} avatar={{color:'colorful',image:photo?{src:photo}:undefined}} secondaryText={secondaryText}/>;
}
