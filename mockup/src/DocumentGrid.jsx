import React,{useMemo} from 'react';
import DocumentThumbnail from './DocumentThumbnail';
import DocumentPerson from './DocumentPerson';
import {DataGrid,DataGridHeader,DataGridHeaderCell,DataGridBody,DataGridRow,DataGridCell,createTableColumn,makeStyles,mergeClasses,tokens,Button,Tooltip,Menu,MenuTrigger,MenuPopover,MenuList,MenuItem,MenuDivider,Badge} from '@fluentui/react-components';
import {DocumentRegular,AddRegular,MoreHorizontalRegular,LinkRegular,WarningRegular} from '@fluentui/react-icons';
const useStyles=makeStyles({
 grid:{minWidth:'1180px',padding:'0 12px',boxSizing:'border-box'},compactGrid:{minWidth:'730px'},
 header:{height:'48px',borderBottom:`1px solid ${tokens.colorNeutralStroke2}`,backgroundColor:tokens.colorNeutralBackground1},
 headerCell:{fontSize:'14px',fontWeight:tokens.fontWeightSemibold,padding:'0 12px','&:hover':{backgroundColor:tokens.colorNeutralBackground1}},
 row:{minHeight:'56px',borderBottom:`1px solid ${tokens.colorNeutralStroke2}`,'&:hover':{backgroundColor:tokens.colorSubtleBackgroundHover},'&:hover .row-actions':{opacity:1},'&:focus-within .row-actions':{opacity:1}},
 selected:{backgroundColor:tokens.colorBrandBackground2,'&:hover':{backgroundColor:tokens.colorBrandBackground2Hover}},
 cell:{padding:'0 12px',fontSize:'14px',overflow:'hidden'},
 selection:{flexBasis:'36px',width:'36px',minWidth:'36px',maxWidth:'36px',padding:0},
 icon:{flexBasis:'36px',width:'36px',minWidth:'36px',maxWidth:'36px',padding:'0 6px'},
 name:{flexBasis:'270px',flexGrow:1,minWidth:'250px'},modified:{flexBasis:'170px',minWidth:'170px',maxWidth:'170px'},by:{flexBasis:'176px',minWidth:'176px',maxWidth:'176px'},type:{flexBasis:'126px',minWidth:'126px',maxWidth:'126px'},status:{flexBasis:'100px',minWidth:'100px',maxWidth:'100px'},expiry:{flexBasis:'160px',minWidth:'160px',maxWidth:'160px'},add:{flexBasis:'136px',minWidth:'136px',maxWidth:'136px'},
 nameCell:{position:'relative',width:'100%',display:'flex',alignItems:'center',justifyContent:'space-between',gap:'12px',minWidth:0},
 fileName:{flex:1,fontWeight:tokens.fontWeightRegular,fontSize:'14px',overflow:'hidden',whiteSpace:'normal',lineHeight:'18px',display:'-webkit-box',WebkitBoxOrient:'vertical',WebkitLineClamp:2,textOverflow:'ellipsis',minWidth:0},
 actions:{position:'absolute',right:0,backgroundColor:tokens.colorSubtleBackgroundHover,display:'flex',opacity:0,flexShrink:0,gap:'2px','& button':{minWidth:'28px',padding:'4px',color:tokens.colorNeutralForeground2}},
 secondary:{fontSize:'13px',color:tokens.colorNeutralForeground2,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'},
 badge:{fontSize:'12px',height:'22px',padding:'0 8px',fontWeight:tokens.fontWeightRegular},
 expiryCell:{display:'flex',flexDirection:'column',gap:'3px',alignItems:'flex-start',padding:'5px 0'},
 failed:{display:'flex',alignItems:'center',gap:'4px',fontSize:'11px',color:tokens.colorPaletteRedForeground1},
 addButton:{paddingLeft:0,fontWeight:tokens.fontWeightRegular,fontSize:'14px',whiteSpace:'nowrap',color:tokens.colorNeutralForeground2},
});
export default function DocumentGrid({rows,selected,setSelected,select,details,FileIcon,head,onPreview,onDetails,feedback}){
 const s=useStyles();
 const columns=useMemo(()=>[
 createTableColumn({columnId:'icon',renderHeaderCell:()=> <DocumentRegular aria-label="File type"/>,renderCell:d=><FileIcon doc={d}/>}),
 createTableColumn({columnId:'name',renderHeaderCell:()=>head('Name'),renderCell:d=><div className={s.nameCell}><DocumentThumbnail doc={d} miniature icon={<FileIcon doc={d}/>}/><div className={s.fileName}>{d.name}{d.status==='Failed'&&<div className={s.failed}><WarningRegular/> Upload failed · Remove</div>}</div><div className={mergeClasses('row-actions',s.actions)}><Tooltip content="Copy link" relationship="label"><Button appearance="subtle" icon={<LinkRegular/>} aria-label={`Copy link to ${d.name}`} onClick={e=>{e.stopPropagation();feedback('Copy link · Visual demonstration only');}}/></Tooltip><Tooltip content="More actions" relationship="label"><Button appearance="subtle" icon={<MoreHorizontalRegular/>} aria-label={`Details for ${d.name}`} onClick={e=>{e.stopPropagation();onDetails(d.id);}}/></Tooltip></div></div>}),
 createTableColumn({columnId:'modified',renderHeaderCell:()=>head('Modified'),renderCell:d=><span className={s.secondary}>{d.modified}</span>}),
 createTableColumn({columnId:'by',renderHeaderCell:()=>head('Modified By'),renderCell:d=><DocumentPerson name={d.by}/>}),
 ...(!details?[
 createTableColumn({columnId:'type',renderHeaderCell:()=>head('Document type'),renderCell:d=><span>{d.type}</span>}),
 createTableColumn({columnId:'status',renderHeaderCell:()=>head('Status'),renderCell:d=><Badge className={s.badge} appearance="tint" color={d.status==='Failed'?'danger':d.status==='Draft'?'subtle':'success'}>{d.status}</Badge>}),
 createTableColumn({columnId:'expiry',renderHeaderCell:()=>head('Expiry date'),renderCell:d=><div className={s.expiryCell}><span className={s.secondary}>{d.expiry==='Expired'?'September 24, 2026':d.expiry}</span>{d.expiry==='Expired'?<Badge className={s.badge} appearance="tint" color="danger">Expired</Badge>:d.expiry==='October 24, 2026'?<Badge className={s.badge} appearance="tint" color="warning">Expires in 20 days</Badge>:null}</div>}),
 createTableColumn({columnId:'add',renderHeaderCell:()=><Menu><MenuTrigger disableButtonEnhancement><Button appearance="transparent" icon={<AddRegular/>} className={s.addButton}>Add column</Button></MenuTrigger><MenuPopover><MenuList><MenuItem onClick={()=>feedback('Source is available in the details pane.')}>Source</MenuItem><MenuItem onClick={()=>feedback('File size is available in the details pane.')}>File size</MenuItem><MenuDivider/><MenuItem onClick={()=>feedback('Default columns restored.')}>Reset to default</MenuItem></MenuList></MenuPopover></Menu>,renderCell:()=>null}),
 ]:[])
 ],[details,head,FileIcon,s,feedback,onDetails]);
 return <DataGrid items={rows} columns={columns} getRowId={d=>d.id} selectionMode="multiselect" selectedItems={new Set(selected)} onSelectionChange={(_,d)=>setSelected([...d.selectedItems])} subtleSelection={!selected.length} selectionAppearance="subtle" focusMode="composite" aria-label="Documents" className={mergeClasses(s.grid,details&&s.compactGrid)}>
 <DataGridHeader><DataGridRow className={s.header} selectionCell={{subtle:false,className:s.selection,checkboxIndicator:{shape:'circular','aria-label':'Select all documents'}}}>{({renderHeaderCell,columnId})=><DataGridHeaderCell sortable={false} className={mergeClasses(s.headerCell,s[columnId])}>{renderHeaderCell()}</DataGridHeaderCell>}</DataGridRow></DataGridHeader>
 <DataGridBody>{({item,rowId})=><DataGridRow key={rowId} data-document-id={rowId} className={mergeClasses(s.row,selected.includes(rowId)&&s.selected)} selectionCell={{className:mergeClasses(s.selection,'select-cell'),checkboxIndicator:{shape:'circular','aria-label':`Select ${item.name}`}}} onClick={e=>select(rowId,e,!!e.target.closest('.select-cell'))} onDoubleClick={()=>onPreview(rowId)} onKeyDown={e=>{if(e.key==='Enter')onPreview(rowId);}}>{({renderCell,columnId})=><DataGridCell focusMode="group" className={mergeClasses(s.cell,s[columnId])}>{renderCell(item)}</DataGridCell>}</DataGridRow>}</DataGridBody>
 </DataGrid>;
}
