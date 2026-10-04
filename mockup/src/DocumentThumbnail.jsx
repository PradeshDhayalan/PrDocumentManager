import React from 'react';
import {makeStyles,mergeClasses,tokens,Text} from '@fluentui/react-components';
import {GlobeRegular,WarningRegular} from '@fluentui/react-icons';
const useStyles=makeStyles({
 root:{height:'100%',width:'100%',position:'relative',overflow:'hidden',display:'flex',alignItems:'center',justifyContent:'center',backgroundColor:tokens.colorNeutralBackground3},
 image:{height:'100%',width:'100%',objectFit:'cover'},
 paper:{width:'106px',height:'142px',padding:'12px',boxSizing:'border-box',backgroundColor:tokens.colorNeutralBackground1,color:tokens.colorNeutralForeground1,boxShadow:tokens.shadow4,marginTop:'20px',flexShrink:0,overflow:'hidden'},
 wordRule:{height:'3px',width:'20px',backgroundColor:tokens.colorPaletteBlueForeground2,marginBottom:'10px'},
 title:{fontSize:'9px',lineHeight:'12px',fontWeight:tokens.fontWeightSemibold,marginBottom:'8px'},
 small:{fontSize:'6px',lineHeight:'9px',color:tokens.colorNeutralForeground2,marginBottom:'7px'},
 lines:{display:'flex',flexDirection:'column',gap:'4px',marginBottom:'10px'},line:{height:'2px',backgroundColor:tokens.colorNeutralStroke2},shortLine:{width:'64%'},signature:{fontSize:'8px',fontStyle:'italic',borderBottom:`1px solid ${tokens.colorNeutralStroke2}`,paddingBottom:'3px',marginTop:'6px'},
 sheet:{width:'180px',height:'104px',overflow:'hidden',backgroundColor:tokens.colorNeutralBackground1,boxShadow:tokens.shadow4,borderRadius:tokens.borderRadiusSmall},
 sheetTitle:{height:'22px',backgroundColor:tokens.colorPaletteGreenBackground2,color:tokens.colorPaletteGreenForeground2,fontSize:'9px',fontWeight:tokens.fontWeightSemibold,padding:'0 8px',display:'flex',alignItems:'center'},
 table:{width:'100%',borderCollapse:'collapse',fontSize:'6px',lineHeight:'12px',color:tokens.colorNeutralForeground2},td:{border:`1px solid ${tokens.colorNeutralStroke2}`,padding:'0 4px',whiteSpace:'nowrap'},th:{backgroundColor:tokens.colorNeutralBackground2,fontWeight:tokens.fontWeightSemibold},
 slide:{width:'180px',height:'102px',padding:'14px',boxSizing:'border-box',backgroundColor:tokens.colorNeutralBackground1,boxShadow:tokens.shadow4,position:'relative',overflow:'hidden'},
 slideHeading:{fontSize:'14px',lineHeight:'16px',fontWeight:tokens.fontWeightSemibold,maxWidth:'120px',color:tokens.colorBrandForeground1},slideEyebrow:{fontSize:'6px',letterSpacing:'1px',color:tokens.colorNeutralForeground2,marginBottom:'5px'},
 chart:{position:'absolute',bottom:'14px',right:'12px',width:'65px',height:'44px',display:'flex',alignItems:'flex-end',gap:'5px',borderBottom:`1px solid ${tokens.colorNeutralStroke2}`},bar:{width:'12px',backgroundColor:tokens.colorBrandBackground},bar2:{height:'30px',backgroundColor:tokens.colorBrandBackground2},bar3:{height:'40px',backgroundColor:tokens.colorPaletteDarkOrangeBackground2},bar1:{height:'20px'},
 textPage:{width:'170px',height:'106px',padding:'12px',boxSizing:'border-box',backgroundColor:tokens.colorNeutralBackground1,boxShadow:tokens.shadow4},code:{fontSize:'7px',lineHeight:'12px',fontFamily:tokens.fontFamilyMonospace,color:tokens.colorNeutralForeground2,margin:0,whiteSpace:'pre-wrap'},
 notebook:{width:'168px',height:'102px',padding:'14px',boxSizing:'border-box',backgroundColor:tokens.colorNeutralBackground1,borderLeft:`12px solid ${tokens.colorPalettePurpleBackground2}`,boxShadow:tokens.shadow4},notebookHeading:{display:'flex',alignItems:'center',gap:'6px',fontSize:'11px',fontWeight:tokens.fontWeightSemibold,marginBottom:'12px'},
 fallback:{display:'flex',alignItems:'center',justifyContent:'center',flexDirection:'column',gap:'6px',fontSize:'11px',color:tokens.colorNeutralForeground3},
 miniature:{width:'46px',height:'34px',border:`1px solid ${tokens.colorNeutralStroke2}`,borderRadius:tokens.borderRadiusSmall,flexShrink:0},
 miniScene:{display:'flex',alignItems:'center',justifyContent:'center',position:'absolute',width:'192px',height:'144px',transform:'scale(.24)',transformOrigin:'top left',top:0,left:0},
});
function Lines({s}){return <div className={s.lines}>{[0,1,2,3,4].map(n=><div key={n} className={mergeClasses(s.line,n===4&&s.shortLine)}/>)}</div>}
export default function DocumentThumbnail({doc,miniature=false,icon}){
 const s=useStyles();
 let preview;
 if(doc.status==='Failed')preview=<div className={s.fallback}><WarningRegular fontSize={28}/><span>Preview unavailable</span></div>;
 else if(doc.fileType==='Photo')return <div className={mergeClasses(s.root,miniature&&s.miniature)}><img className={s.image} src="/previews/site-photo.jpg" alt={`Thumbnail of ${doc.name}`} loading="lazy"/></div>;
 else if(doc.fileType==='Word'||doc.fileType==='PDF')preview=<div className={s.paper}><div className={s.wordRule}/><div className={s.title}>{doc.name.includes('Insurance')?'Certificate of Insurance':doc.name.includes('NDA')?'Non-Disclosure Agreement':doc.name.includes('Statement')?'Statement of Work':'Master Services Agreement'}</div><div className={s.small}>CONTOSO LTD · 2026</div><Lines s={s}/><div className={s.small}>1. Scope and commercial terms</div><Lines s={s}/>{doc.fileType==='PDF'&&<div className={s.signature}>Michael Bose</div>}</div>;
 else if(doc.fileType==='Excel')preview=<div className={s.sheet}><div className={s.sheetTitle}>{doc.name.includes('Budget')?'Budget 2026':'Pricing model'}</div><table className={s.table} aria-hidden="true"><tbody><tr>{['Item','Qty','Rate','Total'].map(c=><td key={c} className={mergeClasses(s.td,s.th)}>{c}</td>)}</tr>{[['Discovery','1','$4,500','$4,500'],['Design','4','$2,200','$8,800'],['Delivery','8','$1,800','$14,400'],['Support','12','$450','$5,400'],['Total','','','$33,100']].map((r,i)=><tr key={i}>{r.map((c,j)=><td key={j} className={s.td}>{c}</td>)}</tr>)}</tbody></table></div>;
 else if(doc.fileType==='PowerPoint')preview=<div className={s.slide}><div className={s.slideEyebrow}>CONTOSO / 2026</div><div className={s.slideHeading}>{doc.name.includes('Board')?'Business review':'Moving forward, together.'}</div><div className={s.small} style={{marginTop:10}}>Strategy · Delivery · Growth</div><div className={s.chart}><div className={mergeClasses(s.bar,s.bar1)}/><div className={mergeClasses(s.bar,s.bar2)}/><div className={mergeClasses(s.bar,s.bar3)}/></div></div>;
 else if(doc.fileType==='Text')preview=<div className={s.textPage}><div className={s.title}>Project kickoff notes</div><pre className={s.code}>{'September 12, 2026\n\nAgenda\n1. Project scope\n2. Delivery milestones\n3. Actions and owners\n\nNext review: September 19'}</pre></div>;
 else if(doc.fileType==='Link')preview=<div className={s.notebook}><div className={s.notebookHeading}><GlobeRegular/> Team notebook</div><div className={s.small}>CONTOSO SALES</div><Lines s={s}/></div>;
 else preview=<div className={s.fallback}>{icon}<span>No preview available</span></div>;
 return <div className={mergeClasses(s.root,miniature&&s.miniature)} role="img" aria-label={`${doc.status==='Failed'?'Unavailable':'Sample'} thumbnail for ${doc.name}`}>{miniature?<div className={s.miniScene}>{preview}</div>:preview}</div>;
}
