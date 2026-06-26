/* eslint-disable */
// ============================================================================
// DLS Property & Deed Map — SPFx client-side web part (framework: none)
// Port of the standalone 2026-06-18 map. Leaflet is BUNDLED (imported below),
// NOT from a CDN (this tenant blocks external scripts); the Leaflet CSS is HARDCODED in
// the LEAFLET_CSS const below (a build-time inline step proved unreliable). The tenant ALLOWS external
// fetch (TN/KY parcel services, ArcGIS geocoder, the deed Worker) and external
// <img> (Esri tiles), so those work as-is.
//
// Features: live parcels per viewport (TN statewide 86 co. + Davidson/Hamilton/
// Rutherford/Montgomery/Williamson/Shelby + KY Simpson/Pulaski/Warren); owner /
// address / parcel search; click a parcel for owner/address/parcel + a book/page-
// first deed lookup (Cloudflare Worker → latest warranty-deed book/page → auto-run
// TitleSearcher; Sumner/Trousdale → US Title Search handoff; owner-name fallback).
//
// SPFx note: popup buttons use data-act/data-id + ONE delegated listener (module
// scope means inline onclick can't see our functions).
// ============================================================================
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import * as LeafletNS from 'leaflet';

const L: any = LeafletNS as any;

// Leaflet 1.9.4 stylesheet, hardcoded inline (the tenant blocks external CSS links, and a
// build-time inline step proved unreliable — so the CSS lives here so it can never be dropped).
const LEAFLET_CSS = `
.leaflet-pane,.leaflet-tile,.leaflet-marker-icon,.leaflet-marker-shadow,.leaflet-tile-container,.leaflet-pane>svg,.leaflet-pane>canvas,.leaflet-zoom-box,.leaflet-image-layer,.leaflet-layer{position:absolute;left:0;top:0;}
.leaflet-container{overflow:hidden;-webkit-tap-highlight-color:transparent;background:#ddd;outline-offset:1px;font-family:"Helvetica Neue",Arial,Helvetica,sans-serif;font-size:12px;line-height:1.5;}
.leaflet-tile,.leaflet-marker-icon,.leaflet-marker-shadow{-webkit-user-select:none;-moz-user-select:none;user-select:none;-webkit-user-drag:none;}
.leaflet-tile{filter:inherit;visibility:hidden;}
.leaflet-tile-loaded{visibility:inherit;}
.leaflet-zoom-box{width:0;height:0;box-sizing:border-box;z-index:800;}
.leaflet-pane{z-index:400;}
.leaflet-tile-pane{z-index:200;}
.leaflet-overlay-pane{z-index:400;}
.leaflet-shadow-pane{z-index:500;}
.leaflet-marker-pane{z-index:600;}
.leaflet-tooltip-pane{z-index:650;}
.leaflet-popup-pane{z-index:700;}
.leaflet-map-pane canvas{z-index:100;}
.leaflet-map-pane svg{z-index:200;}
.leaflet-zoom-animated{transform-origin:0 0;}
.leaflet-control{position:relative;z-index:800;pointer-events:auto;float:left;clear:both;}
.leaflet-top,.leaflet-bottom{position:absolute;z-index:1000;pointer-events:none;}
.leaflet-top{top:0;}.leaflet-right{right:0;}.leaflet-bottom{bottom:0;}.leaflet-left{left:0;}
.leaflet-right .leaflet-control{float:right;margin-right:10px;}
.leaflet-top .leaflet-control{margin-top:10px;}.leaflet-bottom .leaflet-control{margin-bottom:10px;}.leaflet-left .leaflet-control{margin-left:10px;}
.leaflet-bar{box-shadow:0 1px 5px rgba(0,0,0,0.65);border-radius:4px;}
.leaflet-bar a{background-color:#fff;border-bottom:1px solid #ccc;width:26px;height:26px;line-height:26px;display:block;text-align:center;text-decoration:none;color:#000;font:bold 18px/26px "Lucida Console",Monaco,monospace;}
.leaflet-bar a:hover{background-color:#f4f4f4;}
.leaflet-bar a:first-child{border-top-left-radius:4px;border-top-right-radius:4px;}
.leaflet-bar a:last-child{border-bottom-left-radius:4px;border-bottom-right-radius:4px;border-bottom:none;}
.leaflet-control-zoom{border:2px solid rgba(0,0,0,0.2);background-clip:padding-box;border-radius:4px;}
.leaflet-control-layers{box-shadow:0 1px 5px rgba(0,0,0,0.4);background:#fff;border-radius:5px;}
.leaflet-control-layers-toggle{background-image:url(https://unpkg.com/leaflet@1.9.4/dist/images/layers.png);width:36px;height:36px;}
.leaflet-control-layers .leaflet-control-layers-list,.leaflet-control-layers-expanded .leaflet-control-layers-toggle{display:none;}
.leaflet-control-layers-expanded .leaflet-control-layers-list{display:block;position:relative;}
.leaflet-control-layers-expanded{padding:6px 10px 6px 6px;color:#333;background:#fff;}
.leaflet-control-layers label{display:block;font-size:13px;}
.leaflet-control-attribution{background:rgba(255,255,255,0.7);margin:0;padding:0 5px;color:#333;font-size:11px;}
.leaflet-control-attribution a{text-decoration:none;color:#0078A8;}
.leaflet-container .leaflet-overlay-pane svg{max-width:none!important;max-height:none!important;}
.leaflet-container .leaflet-marker-pane img,.leaflet-container .leaflet-shadow-pane img,.leaflet-container .leaflet-tile-pane img,.leaflet-container img.leaflet-image-layer,.leaflet-container .leaflet-tile{max-width:none!important;max-height:none!important;width:auto;padding:0;}
.leaflet-marker-icon,.leaflet-marker-shadow{display:block;}
.leaflet-pane > svg path,.leaflet-tile-container{pointer-events:none;}
.leaflet-pane > svg path.leaflet-interactive,svg.leaflet-image-layer.leaflet-interactive path{pointer-events:visiblePainted;pointer-events:auto;}
.leaflet-zoom-anim .leaflet-zoom-animated{transition:transform 0.25s cubic-bezier(0,0,0.25,1);}
.leaflet-zoom-anim .leaflet-tile,.leaflet-pan-anim .leaflet-tile{transition:none;}
.leaflet-zoom-anim .leaflet-zoom-hide{visibility:hidden;}
.leaflet-container{cursor:grab;}
.leaflet-interactive{cursor:pointer;}
.leaflet-grab{cursor:grab;}
.leaflet-dragging .leaflet-grab,.leaflet-dragging .leaflet-interactive{cursor:grabbing;}
.leaflet-popup{position:absolute;text-align:center;margin-bottom:20px;}
.leaflet-popup-content-wrapper{padding:1px;text-align:left;border-radius:12px;background:#fff;box-shadow:0 3px 14px rgba(0,0,0,0.4);}
.leaflet-popup-content{margin:13px 19px;line-height:1.4;min-height:1px;}
.leaflet-popup-tip-container{width:40px;height:20px;position:absolute;left:50%;margin-top:-1px;margin-left:-20px;overflow:hidden;pointer-events:none;}
.leaflet-popup-tip{width:17px;height:17px;padding:1px;margin:-10px auto 0;pointer-events:auto;background:#fff;box-shadow:0 3px 14px rgba(0,0,0,0.4);transform:rotate(45deg);}
.leaflet-popup-close-button{position:absolute;top:0;right:0;border:none;text-align:center;width:24px;height:24px;font:16px/24px Tahoma,Verdana,sans-serif;color:#757575;text-decoration:none;background:transparent;cursor:pointer;}
.leaflet-popup-close-button:hover{color:#585858;}
.leaflet-fade-anim .leaflet-popup{opacity:0;transition:opacity 0.2s linear;}
.leaflet-fade-anim .leaflet-map-pane .leaflet-popup{opacity:1;}
`;

// ---- TitleSearcher county map (sub:true = flat-rate; else Pay-As-You-Go) ----
const TS_TN: any = {
  CLAY:{c:'T91',sub:true}, JACKSON:{c:'T44',sub:true}, MACON:{c:'T99',sub:true}, SMITH:{c:'63',sub:true},
  ANDERSON:{c:'2'}, BEDFORD:{c:'34'}, BLEDSOE:{c:'T92'}, BRADLEY:{c:'13'}, CAMPBELL:{c:'T12'},
  CARTER:{c:'T7'}, CLAIBORNE:{c:'33'}, COCKE:{c:'T19'}, COFFEE:{c:'40'}, CUMBERLAND:{c:'5'},
  DECATUR:{c:'52'}, FAYETTE:{c:'57'}, FENTRESS:{c:'T25'}, FRANKLIN:{c:'T42'}, GILES:{c:'28'},
  GRAINGER:{c:'T90'}, GREENE:{c:'T65'}, HAMBLEN:{c:'29'}, HAWKINS:{c:'T46'}, HICKMAN:{c:'59'},
  HUMPHREYS:{c:'23'}, JEFFERSON:{c:'20'}, JOHNSON:{c:'T39'}, LAWRENCE:{c:'65'}, LINCOLN:{c:'T48'},
  LOUDON:{c:'T95'}, MADISON:{c:'18'}, MARION:{c:'21'}, MAURY:{c:'53'}, MONROE:{c:'62'},
  MOORE:{c:'M1'}, PERRY:{c:'T68'}, PICKETT:{c:'14'}, POLK:{c:'T70'}, RHEA:{c:'T43'},
  ROANE:{c:'T69'}, SCOTT:{c:'76'}, SEQUATCHIE:{c:'T77'}, SEVIER:{c:'T16'}, SHELBY:{c:'T79'},
  SULLIVAN:{c:'T94'}, UNICOI:{c:'56'}, UNION:{c:'T89'}, VANBUREN:{c:'T88'}, WASHINGTON:{c:'3'},
  WEAKLEY:{c:'32'}, WHITE:{c:'36'}, WILLIAMSON:{c:'T4'}, WILSON:{c:'24'}
};
const US_TN: any = { SUMNER:1, TROUSDALE:1 };   // these go to US Title Search (handoff)
const TS_BASE = 'https://www.titlesearcher.com/';
const US_BASE = 'https://www.ustitlesearch.net/default.asp';

// ---- per-county SOURCE REGISTRY (candidate field names => schema-resilient) ----
const SOURCES: any[] = [
  { id:'tn', label:'TN — Statewide (86 counties)', state:'TN',
    url:'https://services1.arcgis.com/YuVBSS7Y1of2Qud1/arcgis/rest/services/Tennessee_Property_Boundaries_Public_Use/FeatureServer/0/query',
    bbox:[-90.45,34.94,-81.60,36.72], countyField:'COUNTY_NAME', where:'1=1',
    f:{pin:['PARCELID'],owner:['OWNER'],owner2:['OWNER2'],address:['ADDRESS'],subdiv:['SUBDIV'],lot:['LOT'],acres:['DEEDAC'],assr:['LINK_TPV','LINK_TPAD'],tpad:['LINK_TPAD'],gislinkf:['GISLINK']},
    search:{owner:'OWNER',address:'ADDRESS',parcel:'PARCELID'} },
  { id:'davidson', label:'TN — Davidson / Nashville', state:'TN', county:'DAVIDSON',
    url:'https://maps.nashville.gov/arcgis/rest/services/Cadastral/Parcels/MapServer/0/query',
    bbox:[-87.06,35.96,-86.51,36.41], where:"FeatureType IS NULL OR FeatureType<>'Unit'",
    f:{pin:['APN','STANPAR'],owner:['Owner'],address:['PropAddr'],mail:['OwnAddr1'],acres:['Acres','DeededAcreage'],zoning:['Zoning'],deedref:['OwnInstr']},
    search:{owner:'Owner',address:'PropAddr',parcel:'APN'} },
  { id:'hamilton', label:'TN — Hamilton / Chattanooga', state:'TN', county:'HAMILTON',
    url:'https://mapsdev.hamiltontn.gov/hcwa03/rest/services/Live_Parcels/MapServer/0/query',
    bbox:[-85.55,34.98,-84.96,35.46], where:"OWNERNAME1<>'Update in Progress'",
    f:{pin:['PARCEL','TAX_MAP_NO','GISLINK'],owner:['OWNERNAME1'],owner2:['OWNERNAME2'],address:['ADDRESS']},
    search:{owner:'OWNERNAME1',address:'ADDRESS',parcel:'PARCEL'} },
  { id:'rutherford', label:'TN — Rutherford / Murfreesboro', state:'TN', county:'RUTHERFORD',
    url:'https://services.arcgis.com/36I6IHIdr660pAyH/ArcGIS/rest/services/ParcelsCAMA1/FeatureServer/0/query',
    bbox:[-86.62,35.64,-86.03,36.05], where:'GISLINK IS NOT NULL',
    f:{pin:['ParcelID','GISLINK'],owner:['Owner1'],owner2:['Owner2'],address:['FormattedLocation','STREETADDRESS'],mail:['MailingAddress'],subdiv:['SUBDIVISION'],lot:['LOT'],acres:['CALCACRES','DEEDACRES'],zoning:['ZONING'],legalref:['LegalReference']},
    search:{owner:'Owner1',address:'FormattedLocation',parcel:'ParcelID'} },
  { id:'montgomery', label:'TN — Montgomery / Clarksville', state:'TN', county:'MONTGOMERY',
    url:'https://apnsgis4.apsu.edu/arcgis/rest/services/CMCGIS/MontViewer/FeatureServer/2/query',
    bbox:[-87.50,36.39,-87.00,36.71], where:'1=1',
    f:{pin:['parcelid','gislink'],owner:['owner'],owner2:['owner2'],address:['propertyaddress']},
    search:{owner:'owner',address:'propertyaddress',parcel:'parcelid'} },
  { id:'williamson', label:'TN — Williamson / Franklin', state:'TN', county:'WILLIAMSON',
    url:'http://arcgis2.williamson-tn.org/arcgis/rest/services/IDT/DataPull/MapServer/4/query',
    bbox:[-87.18,35.68,-86.68,36.08], where:'1=1', note:'HTTP-only host — blocked from an HTTPS page (mixed content)',
    f:{pin:['parcel_id','GISLINK'],owner:['owner1'],owner2:['owner2'],address:['ADDRESS']},
    search:{owner:'owner1',address:'ADDRESS',parcel:'parcel_id'} },
  { id:'shelby', label:'TN — Shelby / Memphis', state:'TN', county:'SHELBY',
    url:'https://gis.shelbycountytn.gov/public/rest/services/Parcel/CERT_Parcel/MapServer/0/query',
    bbox:[-90.31,34.94,-89.64,35.42], where:'1=1', note:'their DB connection was intermittent',
    f:{pin:['PARCELID','PARID'],owner:['OWNER'],owner2:['OWNER_EXT'],address:['PAR_ADDR1'],mail:['OWN_ADDR1']},
    search:{owner:'OWNER',address:'PAR_ADDR1',parcel:'PARCELID'} },
  { id:'ky_simpson', label:'KY — Simpson / Franklin', state:'KY', county:'SIMPSON',
    url:'https://services8.arcgis.com/D3RgmiBYTvYcNK2j/arcgis/rest/services/Parcel2026view/FeatureServer/0/query',
    bbox:[-86.78,36.62,-86.42,36.87], where:"PIDN<>' '",
    f:{pin:['PIDN'],owner:['NAME'],address:['Property_L'],mail:['Address_Li'],acres:['ACRES'],deedref:['DEED']},
    search:{owner:'NAME',address:'Property_L',parcel:'PIDN'} },
  { id:'ky_pulaski', label:'KY — Pulaski / Somerset', state:'KY', county:'PULASKI',
    url:'https://services5.arcgis.com/cnJiyVVCFyUslPPa/arcgis/rest/services/ParcelUpdate_2026/FeatureServer/2/query',
    bbox:[-84.82,36.91,-84.29,37.29], where:"parcel_id<>' '",
    f:{pin:['parcel_id','Parc_lbl','Account'],owner:['owner1'],owner2:['owner2'],address:['prop_stree'],mail:['own_street'],acres:['legal_acre'],deedBook:['deed_book'],deedPage:['deed_page']},
    search:{owner:'owner1',address:'prop_stree',parcel:'parcel_id'} },
  { id:'ky_warren', label:'KY — Warren / Bowling Green', state:'KY', county:'WARREN',
    url:'https://webgis.bgky.org/server/rest/services/CCPC/CCPC_Parcels/MapServer/0/query',
    bbox:[-86.61,36.77,-86.26,37.11], where:'1=1', ownerWithheld:true,
    f:{pin:['PVA_PARCEL'],address:['ADDRESS'],subdiv:['SUBNAME'],lot:['LOT_NUMBER'],acres:['ACRES'],zoning:['ZONING']},
    search:{address:'ADDRESS',parcel:'PVA_PARCEL'} }
];

const MINZOOM = 14;
const PARCEL_DETAIL_ZOOM = 15;  // in Zoning:View, untagged parcel outlines show only at/above this zoom (keeps the town overview clean)

// ---- Self-tagged zoning layer (reference only), jurisdiction-aware ----
// Each entry = one adopted map: bounds [[S,W],[N,E]] + its OWN districts/colors/names.
// RBS georef is EXACT; Lafayette & Macon are APPROXIMATE (opacity slider + 'approx' badge).
// RBS/Lafayette/Macon all taggable; "Tag lots as" picks the jurisdiction (auto = most-specific at the click).
const ZJURS: any[] = [
  { id:'RBS', name:'Red Boiling Springs', file:'rbs_zoning_overlay.webp', accuracy:'exact', taggable:true,
    bounds:[[36.51467,-85.87444],[36.54963,-85.831]], opacity:0.62,
    zones:['R-1','R-2','C-1','C-2','C-3','I-1'],
    colors:{'R-1':'#FBE10A','R-2':'#F2A23B','C-1':'#F4B0A0','C-2':'#F07F86','C-3':'#E8332E','I-1':'#B7B7B7'},
    names:{'R-1':'Low Density Residential','R-2':'High Density Residential','C-1':'Central Business','C-2':'General Commercial','C-3':'Highway Commercial','I-1':'General Industrial'} },
  { id:'Lafayette', name:'Lafayette', file:'laf_zoning_overlay.webp', accuracy:'approx', taggable:true,
    bounds:[[36.50841,-86.06043],[36.54108,-86.00161]], opacity:0.66,
    zones:['R-1','R-2','C-1','C-2','M','I-1'],
    colors:{'R-1':'#FBE10A','R-2':'#F2A23B','C-1':'#E8332E','C-2':'#F3A0C0','M':'#B59BC9','I-1':'#5BB8E8'},
    names:{'R-1':'Low Density Residential','R-2':'High Density Residential','C-1':'Central Business','C-2':'General Business','M':'Mixed Commercial / Industrial','I-1':'Light Industrial'} },
  { id:'Macon', name:'Macon County', file:'macon_zoning_overlay.webp', accuracy:'approx', taggable:true,
    bounds:[[36.42805,-86.22837],[36.66132,-85.76304]], opacity:0.55,
    zones:['A-1','R-1','R-2','C-1','I-1','I-2'],
    colors:{'A-1':'#CFE0A8','R-1':'#FBE10A','R-2':'#F2A23B','C-1':'#E8332E','I-1':'#C840C8','I-2':'#F58FD0'},
    names:{'A-1':'Agricultural','R-1':'Residential','R-2':'Residential (high density)','C-1':'Commercial','I-1':'Industrial','I-2':'Industrial (heavy)'} }
];
function pinKey(s:any){ return (s==null?'':String(s)).toUpperCase().replace(/\s+/g,' ').trim(); }
function jurById(id:any){ for(let i=0;i<ZJURS.length;i++){ if(ZJURS[i].id===id) return ZJURS[i]; } return null; }
function jurAt(ll:any){ let best:any=null, ba=Infinity; ZJURS.forEach((j:any)=>{ if(!j.taggable) return; const b=j.bounds; if(ll.lat>=b[0][0]&&ll.lat<=b[1][0]&&ll.lng>=b[0][1]&&ll.lng<=b[1][1]){ const area=(b[1][0]-b[0][0])*(b[1][1]-b[0][1]); if(area<ba){ ba=area; best=j; } } }); return best; }
function nearestJur(ll:any){ let best:any=null, bd=Infinity; ZJURS.forEach((j:any)=>{ if(!j.taggable) return; const b=j.bounds; const cy=(b[0][0]+b[1][0])/2, cx=(b[0][1]+b[1][1])/2; const d=(ll.lat-cy)*(ll.lat-cy)+(ll.lng-cx)*(ll.lng-cx); if(d<bd){ bd=d; best=j; } }); return best; }

// ---- straight-line parcel splitting (half-plane clipping; no library). Rings are [[lng,lat],...]. ----
function ringOpen(r:any){ if(!r||r.length<2) return r||[]; const a=r.slice(); if(a.length>1 && a[0][0]===a[a.length-1][0] && a[0][1]===a[a.length-1][1]) a.pop(); return a; }
function ringBounds(r:any){ let mnx=Infinity,mny=Infinity,mxx=-Infinity,mxy=-Infinity; for(let i=0;i<r.length;i++){ const p=r[i]; if(p[0]<mnx)mnx=p[0]; if(p[0]>mxx)mxx=p[0]; if(p[1]<mny)mny=p[1]; if(p[1]>mxy)mxy=p[1]; } return [mnx,mny,mxx,mxy]; }
function sideOf(P:any,A:any,B:any){ return (B[0]-A[0])*(P[1]-A[1])-(B[1]-A[1])*(P[0]-A[0]); }
// Robust split of a simple polygon ring by an INFINITE line (handles concave / irregular lots — no
// half-plane "seam" bleed across a parcel boundary). Returns 1+ simple sub-rings that exactly tile the lot.
function splitRingByLine(ringIn:any,A:any,B:any){
  const ring=ringOpen(ringIn); const n=ring.length; if(n<3) return [ring];
  let scale=1; for(let i=0;i<n;i++){ const ax=Math.abs(ring[i][0]), ay=Math.abs(ring[i][1]); if(ax>scale)scale=ax; if(ay>scale)scale=ay; }
  const EPS=1e-9*scale;
  const P:any[]=[]; const isX:any[]=[];
  for(let i=0;i<n;i++){ const cur=ring[i], nxt=ring[(i+1)%n]; const sc=sideOf(cur,A,B), sn=sideOf(nxt,A,B); P.push(cur); isX.push(Math.abs(sc)<=EPS); if((sc>EPS&&sn<-EPS)||(sc<-EPS&&sn>EPS)){ const t=sc/(sc-sn); P.push([cur[0]+t*(nxt[0]-cur[0]),cur[1]+t*(nxt[1]-cur[1])]); isX.push(true); } }
  const m=P.length;
  const dirx=B[0]-A[0], diry=B[1]-A[1];
  const par=(p:any)=>((p[0]-A[0])*dirx+(p[1]-A[1])*diry);
  const xs:any[]=[]; for(let i=0;i<m;i++){ if(isX[i]) xs.push(i); }
  if(xs.length<2) return [ring];
  xs.sort((a:any,b:any)=>par(P[a])-par(P[b]));
  const pair:any={}; for(let k=0;k+1<xs.length;k+=2){ pair[xs[k]]=xs[k+1]; pair[xs[k+1]]=xs[k]; }
  const usedB:any[]=[]; for(let z=0;z<m;z++){ usedB[z]=false; }
  const pieces:any[]=[];
  for(let s=0;s<m;s++){
    if(usedB[s]) continue;
    const piece:any[]=[]; let i=s; let avc=true; let guard=0;
    while(guard++ < 4*m+8){ piece.push(P[i]); if(isX[i] && (i in pair) && !avc){ i=pair[i]; avc=true; } else { usedB[i]=true; i=(i+1)%m; avc=false; } if(i===s) break; }
    const cl:any[]=[]; for(let k=0;k<piece.length;k++){ const p=piece[k]; if(cl.length===0 || Math.abs(cl[cl.length-1][0]-p[0])>EPS || Math.abs(cl[cl.length-1][1]-p[1])>EPS) cl.push(p); }
    let ar=0; for(let k=0;k<cl.length;k++){ const a=cl[k], b=cl[(k+1)%cl.length]; ar+=a[0]*b[1]-b[0]*a[1]; }
    if(cl.length>=3 && Math.abs(ar)/2>EPS*EPS) pieces.push(cl);
  }
  return pieces.length?pieces:[ring];
}
function splitByLine(pieces:any,A:any,B:any){ let res:any[]=[]; for(let i=0;i<pieces.length;i++){ res=res.concat(splitRingByLine(pieces[i],A,B)); } return res; }
function outerRing(geom:any){ if(!geom) return null; if(geom.type==='Polygon') return ringOpen(geom.coordinates[0]); if(geom.type==='MultiPolygon'){ let best:any=null, ba=-1; for(let i=0;i<geom.coordinates.length;i++){ const r=geom.coordinates[i][0]; const b=ringBounds(r); const area=(b[2]-b[0])*(b[3]-b[1]); if(area>ba){ ba=area; best=r; } } return best?ringOpen(best):null; } return null; }

// ---- module-scope pure helpers ----
function esc(s: any): string { return (s==null?'':String(s)).replace(/[&<>"]/g, (c:string)=>(({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'} as any)[c])); }
function qs(o: any): string { return Object.keys(o).map((k)=>encodeURIComponent(k)+'='+encodeURIComponent(o[k])).join('&'); }
function pick(attrs: any, cands: any): string { if(!cands) return ''; for(let i=0;i<cands.length;i++){ const v=attrs[cands[i]]; if(v!==undefined&&v!==null&&String(v).trim()!=='') return String(v).trim(); } return ''; }
function signedArea(r:any){ let s=0; for(let i=0;i<r.length-1;i++){ s+=(r[i][0]*r[i+1][1]-r[i+1][0]*r[i][1]); } return s; }
function centroid(r:any){ let x=0,y=0; const n=r.length; for(let i=0;i<n;i++){ x+=r[i][0]; y+=r[i][1]; } return [x/n,y/n]; }
function pointInRing(p:any,r:any){ const x=p[0],y=p[1]; let inside=false; for(let i=0,j=r.length-1;i<r.length;j=i++){ const xi=r[i][0],yi=r[i][1],xj=r[j][0],yj=r[j][1]; const hit=((yi>y)!=(yj>y))&&(x<(xj-xi)*(y-yi)/(yj-yi)+xi); if(hit) inside=!inside; } return inside; }
function ringsToGeoJSON(rings:any){ let outers:any[]=[], holes:any[]=[]; rings.forEach((r:any)=>{ (signedArea(r)>=0?outers:holes).push(r); }); if(outers.length===0){ outers=rings; holes=[]; } const polys=outers.map((o:any)=>[o]); holes.forEach((h:any)=>{ const c=centroid(h); let idx=0; for(let i=0;i<polys.length;i++){ if(pointInRing(c,polys[i][0])){ idx=i; break; } } polys[idx].push(h); }); return polys.length===1?{type:'Polygon',coordinates:polys[0]}:{type:'MultiPolygon',coordinates:polys}; }
function esriToFeatures(data:any){ if(!data||!data.features) return []; return data.features.map((ft:any)=>{ let geom=null; if(ft.geometry&&ft.geometry.rings){ geom=ringsToGeoJSON(ft.geometry.rings); } return {type:'Feature',properties:ft.attributes||{},geometry:geom}; }).filter((f:any)=>f.geometry); }
function normalize(attrs:any, src:any){ const n:any={src:src}; n.pin=pick(attrs,src.f.pin); n.owner=pick(attrs,src.f.owner); n.owner2=pick(attrs,src.f.owner2); n.address=pick(attrs,src.f.address); n.mail=pick(attrs,src.f.mail); n.subdiv=pick(attrs,src.f.subdiv); n.lot=pick(attrs,src.f.lot); n.acres=pick(attrs,src.f.acres); n.zoning=pick(attrs,src.f.zoning); n.assr=pick(attrs,src.f.assr); n.tpad=pick(attrs,src.f.tpad); const gm=(n.tpad.match(/gislink=([^&]+)/)||[])[1]; n.gislink=gm?decodeURIComponent(gm):pick(attrs,src.f.gislinkf); n.deedBook=pick(attrs,src.f.deedBook); n.deedPage=pick(attrs,src.f.deedPage); n.legalref=pick(attrs,src.f.legalref); n.deedref=pick(attrs,src.f.deedref); n.state=src.state; n.county=src.county||pick(attrs,[src.countyField]); if(n.county) n.county=n.county.toUpperCase().replace(/ COUNTY$/,'').trim(); return n; }
function parseBookPage(n:any){ if(n.deedBook&&n.deedPage&&/\d/.test(n.deedBook)&&/\d/.test(n.deedPage)) return {book:n.deedBook.replace(/[^0-9A-Za-z]/g,''),page:n.deedPage.replace(/[^0-9A-Za-z]/g,'')}; const ref=n.legalref||''; const m=ref.match(/^\s*([0-9A-Za-z]+)\s*[-\/]\s*([0-9A-Za-z]+)\s*$/); if(m) return {book:m[1],page:m[2]}; return null; }
function tsNameUrl(owner:string){ const name=(owner||'').split(',')[0].trim(); return TS_BASE+'nameSearch.php?'+qs({nameType:'2',searchType:'PA',indexType:'BOTH',p1:name,p2:'',expandAll:'on',startDate:'',endDate:'',itype:'0',executeSearch:'Execute Search'}); }
function tsBookPageUrl(bp:any){ return TS_BASE+'bookPageSearch.php?'+qs({book:bp.book,page:bp.page,fileNumber:'',executeSearch:'Execute Search'}); }
function outFieldsFor(s:any){ const set:any={}; ['pin','owner','owner2','address','mail','subdiv','lot','acres','zoning','assr','tpad','gislinkf','deedBook','deedPage','legalref','deedref'].forEach((k)=>{ (s.f[k]||[]).forEach((fn:string)=>{ set[fn]=1; }); }); if(s.countyField) set[s.countyField]=1; return Object.keys(set).join(',')||'*'; }
function bboxIntersect(a:any,b:any){ return !(b[0]>a[2]||b[2]<a[0]||b[1]>a[3]||b[3]<a[1]); }

export interface IPropertyDeedMapWebPartProps { title: string; workerUrl: string; zoneListTitle: string; zoningAssetBase: string; }

export default class PropertyDeedMapWebPart extends BaseClientSideWebPart<IPropertyDeedMapWebPartProps> {
  private map:any; private parcelLayer:any; private hiLayer:any; private labels:any; private bases:any;
  private POP:any = {}; private pseq=0;
  private inflight:any[] = []; private loadTimer:any = null; private rzTimer:any = null;
  private loadedBounds:any = null; private loadedZoom:number = -1;
  private zoneByPin:any = {}; private zoningView=true; private zoningEdit=false;
  private zTarget:any = null; private loadSeq=0; private tagJur:string='auto';
  private splitState:any=null; private splitLayer:any=null; private splitTmp:any[]=[]; private splitMarkers:any[]=[]; private _splitClick:any=null; private _splitDrawPopup:any=null;
  private femaLayer:any=null; private _femaOn=false; private areasLayer:any=null; private _areasRenderer:any=null; private areas:any[]=[]; private _areasOn=false;
  private areaState:any=null; private areaMarkers:any[]=[]; private areaLine:any=null; private _areaClick:any=null;

  protected onInit(): Promise<void> {
    if (!document.getElementById('dls-leaflet-css')) {
      const st = document.createElement('style'); st.id='dls-leaflet-css'; st.textContent = LEAFLET_CSS; document.head.appendChild(st);
    }
    if (!document.getElementById('dls-hatch-svg')) {
      const hd = document.createElement('div'); hd.id='dls-hatch-svg'; hd.style.cssText='position:absolute;width:0;height:0;overflow:hidden';
      hd.innerHTML='<svg width="0" height="0"><defs><pattern id="dls-hatch" patternUnits="userSpaceOnUse" width="8" height="8"><path d="M0,8 L8,0" stroke="#444" stroke-width="1"/><path d="M0,0 L8,8" stroke="#444" stroke-width="1"/></pattern></defs></svg>';
      document.body.appendChild(hd);
    }
    // ONE delegated listener for popup buttons (data-act). domElement persists across renders.
    this.domElement.addEventListener('click', (e:any) => {
      const t = e.target && e.target.closest ? e.target.closest('[data-act]') : null;
      if (!t) return;
      e.preventDefault();
      this.onAct(t.getAttribute('data-act'), t.getAttribute('data-id'), t.getAttribute('data-arg'));
    });
    document.addEventListener('keydown', (e:any) => { if(e.key==='Escape'){ const fsEl=this.domElement.querySelector('.dls-pm.fs'); if(fsEl) this.toggleFs(); } });
    return Promise.resolve();
  }

  private get workerUrl(): string { return this.properties.workerUrl || 'https://dls-deed.alex-564.workers.dev/'; }
  private get zoneListTitle(): string { return this.properties.zoneListTitle || 'DLS Zoning Assignments'; }
  private get zoningAssetBase(): string { let b=this.properties.zoningAssetBase || (this.context.pageContext.web.absoluteUrl + '/SiteAssets/zoning/'); return b.charAt(b.length-1)==='/'?b:b+'/'; }

  public render(): void {
    this.domElement.innerHTML = `
      <style>
        .dls-pm{font-family:'Segoe UI',Arial,sans-serif;color:#0f172a;width:100%;box-sizing:border-box;}
        @media (min-width:1300px) and (orientation:landscape){ .dls-pm{width:98vw;position:relative;left:50%;margin-left:-49vw;} }
        .dls-pm .bar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;background:#1f2a37;color:#fff;padding:7px 10px;border-radius:8px 8px 0 0;border-bottom:3px solid #f59e0b;}
        .dls-pm .bar strong{font-size:14px;}
        .dls-pm select,.dls-pm input{font-size:12.5px;padding:5px 7px;border:1px solid #3b4a5e;border-radius:5px;background:#fff;color:#0f172a;}
        .dls-pm #q{width:210px;}
        .dls-pm button{cursor:pointer;border:1px solid #3b4a5e;border-radius:5px;background:#f59e0b;color:#1a1205;font-weight:600;padding:5px 10px;font-size:12.5px;}
        .dls-pm button.ghost{background:#33445a;color:#fff;}
        .dls-pm .sp{flex:1;}
        .dls-pm #status{font-size:11px;color:#9fb0c3;white-space:nowrap;}
        .dls-pm .stage{position:relative;}
        .dls-pm #map{height:calc(88vh - 220px);min-height:460px;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 8px 8px;}
        .dls-pm.fs{position:fixed;inset:0;left:0;margin-left:0;width:auto;z-index:100000;background:#1f2a37;display:flex;flex-direction:column;border-radius:0;}
        .dls-pm.fs .bar{border-radius:0;}
        .dls-pm.fs .stage{flex:1;min-height:0;}
        .dls-pm.fs #map{height:100% !important;min-height:0;border-radius:0;border:none;}
        .dls-pm #results{position:absolute;z-index:1000;top:8px;left:8px;width:290px;max-height:75%;overflow:auto;background:#fff;border:1px solid #cbd5e1;border-radius:8px;box-shadow:0 6px 20px rgba(0,0,0,.18);display:none;}
        .dls-pm #results h4{margin:0;padding:8px 10px;background:#1f2a37;color:#fff;font-size:12px;border-radius:8px 8px 0 0;display:flex;justify-content:space-between;}
        .dls-pm #results h4 .x{cursor:pointer;color:#cbd5e1;font-weight:700;}
        .dls-pm .rrow{padding:7px 10px;border-bottom:1px solid #eef2f7;cursor:pointer;font-size:12px;}
        .dls-pm .rrow:hover{background:#fff7ec;} .dls-pm .rrow b{display:block;} .dls-pm .rrow span{color:#64748b;}
        .dls-pm #legend{position:absolute;z-index:900;bottom:14px;left:8px;background:rgba(255,255,255,.94);border:1px solid #cbd5e1;border-radius:8px;padding:8px 10px;font-size:11px;max-width:260px;}
        .dls-pm .src{color:#64748b;} .dls-pm .disc{font-size:10px;color:#64748b;margin-top:6px;border-top:1px dashed #cbd5e1;padding-top:4px;}
        .lp h3{margin:0 0 4px;font-size:14px;} .lp .co{color:#64748b;font-size:11px;margin-bottom:6px;}
        .lp table{border-collapse:collapse;font-size:12px;margin-bottom:6px;} .lp td{padding:1px 6px 1px 0;vertical-align:top;} .lp td.k{color:#64748b;white-space:nowrap;}
        .lp .deed{border-top:1px solid #e2e8f0;padding-top:6px;margin-top:2px;} .lp .deed .lbl{font-size:11px;color:#64748b;margin-bottom:3px;}
        .lp a.btn,.lp button.cp{display:inline-block;font-size:11px;font-weight:600;text-decoration:none;border-radius:5px;padding:4px 8px;margin:2px 3px 2px 0;border:1px solid #cbd5e1;cursor:pointer;}
        .lp a.ts{background:#16a34a;color:#fff;border-color:transparent;} .lp a.ts.payg{background:#d97706;}
        .lp a.us{background:#0369a1;color:#fff;border-color:transparent;} .lp a.assr{background:#475569;color:#fff;border-color:transparent;}
        .lp button.cp{background:#f1f5f9;} .lp .note{font-size:10px;color:#64748b;margin-top:4px;}
        .badge{display:inline-block;padding:1px 6px;border-radius:9px;font-size:10px;font-weight:700;color:#fff;}
        .b-ok{background:#16a34a;} .b-warn{background:#d97706;}
        .leaflet-popup-content{margin:10px 12px;max-width:280px;}
        .dls-pm #zmode{background:#33445a;color:#fff;}
        .dls-pm #zlegend{position:absolute;z-index:900;bottom:14px;right:8px;background:rgba(255,255,255,.96);border:1px solid #cbd5e1;border-radius:8px;padding:8px 10px;font-size:11px;width:236px;max-height:78%;overflow:auto;display:none;}
        .dls-pm #zlegend b{font-size:11.5px;} .dls-pm #zlegend .zi{display:flex;align-items:center;gap:6px;margin:3px 0;}
        .dls-pm #zlegend .zsw{width:13px;height:13px;border-radius:3px;border:1px solid rgba(0,0,0,.3);flex:none;}
        .dls-pm #zlegend .zdisc{margin-top:5px;border-top:1px dashed #cbd5e1;padding-top:4px;color:#64748b;}
        .dls-pm #zlegend .zrow{display:flex;align-items:center;gap:5px;margin:4px 0;}
        .dls-pm #zlegend .zrow label{flex:1;display:flex;align-items:center;gap:5px;cursor:pointer;}
        .dls-pm #zlegend .zrow input[type=range]{width:58px;}
        .dls-pm #zlegend .zacc{font-size:9px;padding:0 5px;border-radius:7px;}
        .dls-pm #zlegend .zacc.exact{background:#dcfce7;color:#166534;} .dls-pm #zlegend .zacc.approx{background:#fef3c7;color:#92400e;}
        .dls-pm #zlegend .zdiv{border-top:1px solid #cbd5e1;margin:6px 0;}
        .dls-pm #zlegend .zjh{font-weight:700;font-size:10.5px;margin:5px 0 2px;color:#334155;}
        .dls-pm #zlegend .ztag{margin:5px 0 7px;font-size:11px;color:#334155;}
        .dls-pm #zlegend .ztag select{font-size:11px;padding:2px 4px;margin-top:2px;max-width:100%;}
        .zp{font-family:'Segoe UI',Arial,sans-serif;min-width:212px;} .zp .zp-h{font-weight:700;font-size:13px;margin-bottom:2px;}
        .zp .zp-pin,.zp .zp-cur{font-size:11px;color:#475569;} .zp .zp-cur{margin-bottom:6px;}
        .zp .zp-grid{display:grid;grid-template-columns:1fr 1fr;gap:4px;margin-bottom:6px;}
        .zp .zbtn{border:1px solid #94a3b8;border-radius:5px;padding:5px 6px;font-size:11px;font-weight:700;cursor:pointer;color:#1a1205;text-align:left;}
        .zp .zbtn small{display:block;font-weight:400;font-size:9px;color:#334155;line-height:1.15;margin-top:1px;}
        .zp .zp-fl{display:flex;align-items:center;gap:6px;font-size:11.5px;margin:4px 0;cursor:pointer;}
        .zp .zp-clear{background:#f1f5f9;border:1px solid #cbd5e1;border-radius:5px;padding:4px 8px;font-size:11px;cursor:pointer;}
        .zp .zp-note{font-size:10px;color:#64748b;margin-top:5px;}
        .zp .zbtn2{display:block;width:100%;text-align:left;border:1px solid #cbd5e1;border-radius:5px;padding:5px 7px;margin:3px 0;font-size:11.5px;background:#f8fafc;cursor:pointer;color:#0f172a;}
        .zp .zbtn2:hover{background:#eef2f7;}
        .zp .sp-opts{margin:4px 0;}
        .zp .spc{border-top:1px solid #e2e8f0;padding-top:4px;margin-top:4px;}
        .zp .spc-h{font-size:11px;margin-bottom:3px;color:#475569;} .zp .spc-h span{color:#0f172a;font-weight:700;margin-left:6px;}
        .zp .spc-g{display:flex;flex-wrap:wrap;gap:3px;}
        .zp .zbtns{border:1px solid #94a3b8;border-radius:4px;padding:3px 7px;font-size:10.5px;font-weight:700;color:#1a1205;cursor:pointer;}
        .zp .zbnone{background:#f1f5f9 !important;color:#475569;font-weight:600;}
        .zp .zp-save{background:#16a34a;color:#fff;border:1px solid transparent;border-radius:5px;padding:4px 10px;font-size:11px;font-weight:600;cursor:pointer;}
        .znum{background:#1d4ed8 !important;color:#fff;border:none !important;border-radius:50%;width:18px !important;height:18px !important;line-height:18px;text-align:center;font-size:11px;font-weight:700;box-shadow:0 1px 3px rgba(0,0,0,.45);}
      </style>
      <div class="dls-pm">
        <div class="bar">
          <strong>${esc(this.properties.title) || 'DLS Property &amp; Deed Map'}</strong>
          <select id="area" title="Dataset to search / jump to"></select>
          <select id="mode"><option value="owner">Owner</option><option value="address">Address</option><option value="parcel">Parcel ID</option></select>
          <input id="q" placeholder="Search owner name&hellip;" />
          <button id="go">Search</button>
          <button id="clear" class="ghost">Clear</button>
          <span class="sp"></span>
          <select id="base"><option value="aerial">Aerial</option><option value="streets" selected>Streets</option><option value="topo">Topo</option></select>
          <select id="zmode" title="Zoning layer (View / Edit)"><option value="off">Zoning: Off</option><option value="view" selected>Zoning: View</option><option value="edit">Zoning: Edit (tag lots)</option></select>
          <button id="fs" class="ghost" title="Full screen (Esc to exit)">Full screen</button>
          <span id="status">Loading&hellip;</span>
        </div>
        <div class="stage">
          <div id="map"></div>
          <div id="results"><h4><span id="rtitle">Results</span><span class="x" id="rclose">&times;</span></h4><div id="rlist"></div></div>
          <div id="legend"><b>Parcels load at zoom ${MINZOOM}+</b> &mdash; pan/zoom to your area.<br/><span class="src" id="legsrc">Active data: &mdash;</span><div class="disc">Reference only &mdash; not a boundary survey, title opinion, or zoning determination. Parcel &amp; owner data are pulled live from each assessor and may lag.</div></div>
          <div id="zlegend"></div>
        </div>
      </div>`;

    const $ = (s:string)=>this.domElement.querySelector(s) as any;
    const areaSel = $('#area');
    SOURCES.forEach((s)=>{ const o=document.createElement('option'); o.value=s.id; o.textContent=s.label; areaSel.appendChild(o); });
    $('#mode').onchange = (e:any)=>{ const p:any={owner:'Search owner name…',address:'Search street address…',parcel:'Search parcel ID…'}; $('#q').placeholder=p[e.target.value]; };
    $('#go').onclick = ()=>this.runSearch();
    $('#q').addEventListener('keydown',(e:any)=>{ if(e.key==='Enter') this.runSearch(); });
    $('#clear').onclick = ()=>{ this.hiLayer.clearLayers(); $('#results').style.display='none'; };
    $('#rclose').onclick = ()=>{ $('#results').style.display='none'; };
    $('#base').onchange = (e:any)=>this.setBase(e.target.value);
    $('#zmode').onchange = (e:any)=>this.setZoningMode(e.target.value);
    $('#fs').onclick = ()=>this.toggleFs();
    this.buildMap();
    this.buildZPanel();
    this.setZoningMode('view');   // default to Zoning: View with the side panel open
  }

  private buildMap(): void {
    const mapEl = this.domElement.querySelector('#map');
    this.map = L.map(mapEl,{minZoom:6,maxZoom:20}).setView([36.521,-86.029],14);
    this.bases = {
      aerial: L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',{maxZoom:20,attribution:'Imagery © Esri'}),
      streets: L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}',{maxZoom:20,attribution:'© Esri'}),
      topo: L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer/tile/{z}/{y}/{x}',{maxZoom:20,attribution:'© Esri'})
    };
    this.bases.streets.addTo(this.map);
    this.labels = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}',{maxZoom:20,opacity:.9});
    this.map.createPane('zoning'); this.map.getPane('zoning').style.zIndex='350'; this.map.getPane('zoning').style.pointerEvents='none';
    this.map.createPane('zsplit'); this.map.getPane('zsplit').style.zIndex='420'; this.map.getPane('zsplit').style.pointerEvents='none';
    this.splitLayer = L.layerGroup().addTo(this.map);
    this.map.createPane('areas'); this.map.getPane('areas').style.zIndex='430'; this.map.getPane('areas').style.pointerEvents='none';
    this._areasRenderer = L.svg({pane:'areas'}); this._areasRenderer.addTo(this.map);
    this.areasLayer = L.layerGroup().addTo(this.map);
    const FemaTiles:any = L.TileLayer.extend({ getTileUrl:function(coords:any){ const map=this._map; const ts=this.getTileSize(); const nw=map.unproject(L.point(coords.x*ts.x,coords.y*ts.y),coords.z); const se=map.unproject(L.point((coords.x+1)*ts.x,(coords.y+1)*ts.y),coords.z); const a=L.CRS.EPSG3857.project(nw), b=L.CRS.EPSG3857.project(se); const bbox=Math.min(a.x,b.x)+','+Math.min(a.y,b.y)+','+Math.max(a.x,b.x)+','+Math.max(a.y,b.y); return 'https://hazards.fema.gov/arcgis/rest/services/public/NFHL/MapServer/export?bbox='+bbox+'&bboxSR=3857&imageSR=3857&size='+ts.x+','+ts.y+'&dpi=96&format=png32&transparent=true&f=image'; } });
    this.femaLayer = new FemaTiles('', {tileSize:256, opacity:0.55, pane:'zoning', maxZoom:20, attribution:'Flood data © FEMA NFHL'});
    ZJURS.forEach((j:any)=>{ j._layer = L.imageOverlay(this.zoningAssetBase+j.file, j.bounds, {opacity:j.opacity, interactive:false, pane:'zoning'}); j._on = false; });
    this.parcelLayer = L.geoJSON(null,{ style:(ft:any)=>this.parcelStyle(ft), onEachFeature:(ft:any,layer:any)=>this.onFeat(ft,layer) }).addTo(this.map);
    this.hiLayer = L.geoJSON(null,{ style:{color:'#ff2d55',weight:3,fill:false} }).addTo(this.map);
    this.map.on('moveend',()=>{ clearTimeout(this.loadTimer); this.loadTimer=setTimeout(()=>this.maybeLoad(),250); });
    this.map.on('zoomend',()=>{ try{ this.restyleParcels(); }catch(e){} });   // toggle untagged-parcel outlines at PARCEL_DETAIL_ZOOM
    this.setStatus('Pan/zoom to your area — parcels load at zoom '+MINZOOM+'+');
    setTimeout(()=>{ try{ this.map.invalidateSize(); }catch(e){} this.loadParcels(); },400);
    window.addEventListener('resize',()=>{ clearTimeout(this.rzTimer); this.rzTimer=setTimeout(()=>{ try{ if(this.map) this.map.invalidateSize(); }catch(e){} },200); });
    this.loadZoning();
    this.loadAreas();
  }

  private setBase(v:string): void {
    Object.keys(this.bases).forEach((k)=>this.map.removeLayer(this.bases[k]));
    this.bases[v].addTo(this.map);
    if(v==='aerial') this.labels.addTo(this.map); else this.map.removeLayer(this.labels);
  }

  private activeSources(): any[] {
    const b=this.map.getBounds(); const vb=[b.getWest(),b.getSouth(),b.getEast(),b.getNorth()];
    return SOURCES.filter((s)=>bboxIntersect(vb,s.bbox));
  }

  // Fetch ArcGIS JSON; if a server rejects resultRecordCount (no pagination support), retry once without it.
  private arcgisFetch(url:string, signal?:any): Promise<any> {
    const go=(u:string)=>fetch(u, signal?{signal:signal}:{}).then((r)=>r.json());
    return go(url).then((d:any)=>{
      if(d && d.error && /pagination|resultRecordCount|exceed/i.test(String(d.error.message||''))){
        const u2=url.replace(/&resultRecordCount=\d+/,'');
        if(u2!==url) return go(u2);
      }
      return d;
    });
  }

  // Only reload if the view left the already-loaded (padded) area — keeps popups open and avoids flicker on small pans / popup auto-pan.
  private maybeLoad(): void {
    const z=this.map.getZoom();
    if(z>=MINZOOM && this.loadedBounds && this.loadedZoom===z && this.loadedBounds.contains(this.map.getBounds())) return;
    this.loadParcels();
  }

  private loadParcels(): void {
    this.inflight.forEach((c)=>{ try{c.abort();}catch(e){} }); this.inflight=[];
    const mySeq = ++this.loadSeq;   // anti-flicker: a stale load must not clobber a newer one
    const legsrc = this.domElement.querySelector('#legsrc') as any;
    if(this.map.getZoom()<MINZOOM){ this.parcelLayer.clearLayers(); this.loadedBounds=null; this.loadedZoom=-1; this.setStatus('Zoom in to load parcels (zoom ≥ '+MINZOOM+')'); if(legsrc) legsrc.textContent='Active data: —'; return; }
    const srcs=this.activeSources();
    if(srcs.length===0){ this.parcelLayer.clearLayers(); this.loadedBounds=null; this.setStatus('No parcel source covers this view'); return; }
    if(legsrc) legsrc.textContent='Active data: '+srcs.map((s)=>s.label.replace(/^..? — /,'')).join(', ');
    const pb=this.map.getBounds().pad(0.6); this.loadedBounds=pb; this.loadedZoom=this.map.getZoom();
    const env=[pb.getWest(),pb.getSouth(),pb.getEast(),pb.getNorth()].join(',');
    this.setStatus('Loading parcels…');   // keep the OLD parcels on screen until the new set is ready (no flash)
    let got=0, done=0; const errs:string[]=[]; const acc:any[]=[];
    const short=(s:any)=>s.label.replace(/^..? — /,'');
    const finish=()=>{ if(mySeq!==this.loadSeq) return; if(done===srcs.length){ this.parcelLayer.clearLayers(); if(acc.length) this.parcelLayer.addData(acc); this.setStatus(got+' parcels'+(errs.length?'  · unavailable: '+errs.join('; '):'')); } };
    const httpsPage = (typeof location!=='undefined' && location.protocol==='https:');
    srcs.forEach((s)=>{
      if(httpsPage && /^http:\/\//i.test(s.url)){ errs.push(short(s)+' (HTTP-only — needs HTTPS proxy)'); done++; finish(); return; }
      const ctrl=new AbortController(); this.inflight.push(ctrl);
      const url=s.url+'?'+qs({where:s.where||'1=1',geometry:env,geometryType:'esriGeometryEnvelope',inSR:4326,spatialRel:'esriSpatialRelIntersects',outFields:outFieldsFor(s),returnGeometry:true,outSR:4326,resultRecordCount:2000,f:'json'});
      this.arcgisFetch(url,ctrl.signal).then((d:any)=>{
        if(d.error) throw new Error(d.error.message||'service error');
        const feats=esriToFeatures(d); feats.forEach((f:any)=>{ f.properties.__src=s.id; acc.push(f); }); got+=feats.length;
      }).catch((e:any)=>{ if(e.name!=='AbortError') errs.push(short(s)+' ('+e.message+')'); })
        .then(()=>{ done++; finish(); });
    });
  }

  private onFeat(feat:any, layer:any): void {
    // Standalone MAP popup (NOT bound to the parcel layer) so a parcel reload can't close it → no re-clicking.
    layer.on('click',(ev:any)=>{
      if(this.splitState||this.areaState) return;
      const src=SOURCES.filter((s)=>s.id===feat.properties.__src)[0]||SOURCES[0];
      const n=normalize(feat.properties,src);
      const ll=(ev&&ev.latlng)||(layer.getBounds&&layer.getBounds().getCenter());
      if(this.zoningEdit){ this.openZonePicker(n, ll, feat); return; }
      L.popup({maxWidth:300,autoPanPadding:[24,24]}).setLatLng(ll).setContent(this.popupHtml(n)).openOn(this.map);
    });
  }

  private popupHtml(n:any): string {
    const id='_p'+(this.pseq++); this.POP[id]=n;
    let rows='';
    const row=(k:string,v:any)=>{ if(v) rows+='<tr><td class="k">'+k+'</td><td>'+esc(v)+'</td></tr>'; };
    let owner=n.owner+(n.owner2?'; '+n.owner2:'');
    if(n.src.ownerWithheld && !owner) owner='<i>(owner not published by county)</i>';
    rows+='<tr><td class="k">Owner</td><td>'+(owner||'—')+'</td></tr>';
    row('Address',n.address); row('Parcel',n.pin);
    if(n.acres) row('Acres', (+n.acres? (+n.acres).toFixed(2):n.acres));
    if(n.subdiv) row('Subdiv', n.subdiv+(n.lot?'  Lot '+n.lot:''));
    if(n.zoning) row('Zoning', n.zoning);
    const zt=this.zoneByPin[pinKey(n.pin)]; if(zt){ const zj=jurById(zt.jur)||ZJURS[0]; if(zt.split&&zt.pieces){ row('Zone ('+(zt.jur||'')+')', zt.pieces.map((p:any)=>p.z||'blank').join(' / ')+' (split lot)'); } else { row('Zone ('+(zt.jur||'')+')', zt.zone+' — '+((zj.names&&zj.names[zt.zone])||'')+(zt.flood?' · Floodplain':'')); } }
    if(n.deedBook||n.deedPage) row('Deed','Bk '+n.deedBook+' Pg '+n.deedPage);
    else if(n.legalref) row('Deed ref', n.legalref);
    else if(n.deedref) row('Deed ref', n.deedref);
    return '<div class="lp"><h3>'+esc(n.county||'')+(n.state?', '+n.state:'')+'</h3><div class="co">'+esc(n.src.label)+'</div><table>'+rows+'</table>'+this.deedSection(n,id)+'</div>';
  }

  private deedSection(n:any, id:string): string {
    let out='<div class="deed"><div class="lbl">Deed records</div>';
    const e=this.POP[id]; e.owner=n.owner; e.gislink=n.gislink; e.county=n.county;
    const tsInfo = n.state==='TN' ? TS_TN[n.county] : null;
    const isUS = n.state==='TN' && US_TN[n.county];
    e.tsCnum = tsInfo ? tsInfo.c : null;
    e.localBP = parseBookPage(n);
    e.site = isUS ? 'US' : (tsInfo ? 'TS' : null);
    if(e.site==='TS'){
      const sub = tsInfo.sub===true;
      const badge = sub ? '<span class="badge b-ok">included</span>' : '<span class="badge b-warn">pay-per-use</span>';
      out+='<a class="btn ts'+(sub?'':' payg')+'" href="#" data-act="deedGo" data-id="'+id+'">Deed search &rarr; TitleSearcher: '+esc(n.county)+'</a> '+badge;
      out+='<div class="note">'+(((e.gislink&&this.workerUrl)||e.localBP)?'Pulls the latest warranty-deed book/page automatically, then searches; ':'')+'falls back to owner-name search. In results, pick the WD row matching the owner.</div>';
    } else if(e.site==='US'){
      out+='<a class="btn us" href="#" data-act="deedGoUS" data-id="'+id+'">Deed search &rarr; US Title Search: '+esc(n.county)+'</a>';
      out+='<div class="note">Opens US Title Search (your session) and surfaces the latest warranty-deed book/page to enter — that site has no direct deep-link.</div>';
    } else if(n.state==='TN'){
      out+='<span class="note">No deed site mapped for '+esc(n.county)+' — use the assessor link.</span><br/>';
    } else if(n.state==='KY'){
      out+='<a class="btn ts payg" href="'+TS_BASE+'countySelect.php" target="_blank" rel="noopener">TitleSearcher (KY) · pick county</a> <span class="note">KY not yet mapped</span><br/>';
    }
    out+='<div style="margin-top:5px">';
    if(n.owner) out+='<button class="cp" data-act="cpf" data-id="'+id+'" data-arg="owner">Copy owner</button>';
    if(n.pin) out+='<button class="cp" data-act="cpf" data-id="'+id+'" data-arg="pin">Copy parcel</button>';
    if(e.site==='TS' && n.owner) out+='<button class="cp" data-act="deedName" data-id="'+id+'">Name search</button>';
    out+='</div>';
    if(n.assr) out+='<a class="btn assr" href="'+esc(n.assr)+'" target="_blank" rel="noopener">Assessor record</a> ';
    else if(n.state==='TN') out+='<a class="btn assr" href="https://assessment.cot.tn.gov/RE_Assessment/" target="_blank" rel="noopener">TN assessment</a> ';
    return out+'</div>';
  }

  private onAct(act:string, id:string, arg:string): void {
    if(act==='deedGo') this.deedGo(id);
    else if(act==='deedName') this.deedName(id);
    else if(act==='deedGoUS') this.deedGoUS(id);
    else if(act==='cpf'){ const e=this.POP[id]; if(e) this.copyText(e[arg]||''); }
    else if(act==='zset') this.saveZone(arg);
    else if(act==='zclear') this.clearZone();
    else if(act==='zsplitopen') this.startSplitMenu();
    else if(act==='zsplit') this.beginSplit(arg);
    else if(act==='zpz'){ const p=(arg||'').split('|'); this.setPieceZone(+p[0], p[1]); }
    else if(act==='zsplitsave') this.saveSplit();
    else if(act==='zsplitcancel') this.cancelSplit();
    else if(act==='zareafinish') this.finishAreaDraw();
    else if(act==='zareacancel') this.cancelAreaDraw();
    else if(act==='zareasave') this.saveArea(arg);
  }

  private openDeferred(): any { const w=window.open('','_blank'); try{ if(w) w.document.write('<p style="font:14px/1.4 sans-serif;padding:18px;color:#333">Looking up the latest deed…</p>'); }catch(e){} return w; }
  private tsCountyThen(w:any,cnum:string,url:string): void { if(!w) return; w.location=TS_BASE+'countySearchPage.php?cnum='+cnum; setTimeout(()=>{ try{w.location=url;}catch(e){} },1600); }

  private deedGo(id:string): void {
    const e=this.POP[id]; if(!e||!e.tsCnum) return;
    const w=this.openDeferred();
    const nameFallback=()=>this.tsCountyThen(w,e.tsCnum,tsNameUrl(e.owner||''));
    if(e.localBP){ this.tsCountyThen(w,e.tsCnum,tsBookPageUrl(e.localBP)); return; }
    if(e.gislink && this.workerUrl){
      fetch(this.workerUrl+'?gislink='+encodeURIComponent(e.gislink)).then((r)=>r.json())
        .then((d:any)=>{ if(d&&d.ok&&d.best&&d.best.book){ this.setStatus('Latest deed: '+(d.best.type||'')+' Bk '+d.best.book+' Pg '+d.best.page); this.tsCountyThen(w,e.tsCnum,tsBookPageUrl(d.best)); } else { this.setStatus('No book/page found — using owner-name search'); nameFallback(); } })
        .catch(()=>nameFallback());
    } else nameFallback();
  }
  private deedName(id:string): void { const e=this.POP[id]; if(!e||!e.tsCnum) return; this.tsCountyThen(this.openDeferred(),e.tsCnum,tsNameUrl(e.owner||'')); }
  private deedGoUS(id:string): void {
    const e=this.POP[id]; window.open(US_BASE,'_blank');
    const show=(bp:any)=>{ this.setStatus('US Title Search · '+(e.county||'')+': Begin Search → Book/Page → Book '+bp.book+'  Page '+bp.page+(bp.type?'  ('+bp.type+')':'')); this.copyText(bp.book+' '+bp.page); };
    if(e.localBP){ show(e.localBP); return; }
    if(e.gislink && this.workerUrl){
      fetch(this.workerUrl+'?gislink='+encodeURIComponent(e.gislink)).then((r)=>r.json())
        .then((d:any)=>{ if(d&&d.ok&&d.best&&d.best.book) show(d.best); else this.setStatus('No book/page found — use name search in US Title Search'); })
        .catch(()=>this.setStatus('Deed lookup unavailable — use name search in US Title Search'));
    } else this.setStatus('Opened US Title Search → Begin Search.');
  }
  private copyText(txt:string): void { try{ if((navigator as any).clipboard) (navigator as any).clipboard.writeText(txt); }catch(e){} this.setStatus('Copied: '+txt); }

  private runSearch(): void {
    const $ = (s:string)=>this.domElement.querySelector(s) as any;
    const src=SOURCES.filter((s)=>s.id===$('#area').value)[0]; const mode=$('#mode').value; const term=($('#q').value||'').trim();
    if(term.length<2){ this.setStatus('Type at least 2 characters'); return; }
    const field=src.search[mode];
    if(!field){ this.setStatus(src.label+' has no '+mode+' field'); return; }
    let where;
    const sql=(x:string)=>x.replace(/'/g,"''");
    if(mode==='parcel') where="UPPER("+field+") LIKE '"+sql(term.toUpperCase())+"%'";
    else where=term.toUpperCase().split(/\s+/).map((t:string)=>"UPPER("+field+") LIKE '%"+sql(t)+"%'").join(' AND ');
    if(src.where && src.where!=='1=1') where='('+src.where+') AND ('+where+')';
    this.setStatus('Searching '+src.label.replace(/^..? — /,'')+'…');
    const url=src.url+'?'+qs({where:where,outFields:outFieldsFor(src),returnGeometry:true,outSR:4326,resultRecordCount:60,f:'json'});
    this.arcgisFetch(url).then((d:any)=>{ if(d.error) throw new Error(d.error.message||'error'); const feats=esriToFeatures(d); feats.forEach((f:any)=>{ f.properties.__src=src.id; }); this.showResults(feats,src); })
      .catch((e:any)=>this.setStatus('Search failed: '+e.message));
  }

  private showResults(feats:any[], src:any): void {
    const box=this.domElement.querySelector('#results') as any; const list=this.domElement.querySelector('#rlist') as any;
    (this.domElement.querySelector('#rtitle') as any).textContent=feats.length+' result'+(feats.length===1?'':'s');
    list.innerHTML='';
    if(feats.length===0){ list.innerHTML='<div class="rrow"><span>No matches.</span></div>'; box.style.display='block'; this.setStatus('No matches'); return; }
    feats.forEach((f:any)=>{ const n=normalize(f.properties,src); const r=document.createElement('div'); r.className='rrow'; r.innerHTML='<b>'+esc(n.owner||n.address||n.pin||'(parcel)')+'</b><span>'+esc([n.address,n.pin].filter(Boolean).join(' · '))+'</span>'; r.onclick=()=>this.gotoFeature(f,n); list.appendChild(r); });
    box.style.display='block'; this.setStatus(feats.length+' result(s)');
  }

  private gotoFeature(f:any, n:any): void {
    this.hiLayer.clearLayers(); this.hiLayer.addData(f);
    try{ this.map.fitBounds(this.hiLayer.getBounds(),{maxZoom:18,padding:[40,40]}); }catch(e){}
    L.popup({maxWidth:300}).setLatLng(this.hiLayer.getBounds().getCenter()).setContent(this.popupHtml(n)).openOn(this.map);
  }

  // ======================= RBS zoning layer =======================
  private cfg(): any { return (SPHttpClient as any).configurations.v1; }
  private listApi(): string { return this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + this.zoneListTitle.replace(/'/g,"''") + "')"; }
  private spGet(url:string): Promise<any> { return this.context.spHttpClient.get(url, this.cfg(), {headers:{Accept:'application/json;odata=nometadata'}}).then((r:any)=>r.json()); }
  private spPost(url:string, body:any, extra?:any): Promise<any> {
    const headers:any = {Accept:'application/json;odata=nometadata','Content-Type':'application/json;odata=nometadata','odata-version':''};
    if(extra){ for(const k in extra){ headers[k]=extra[k]; } }
    return this.context.spHttpClient.post(url, this.cfg(), {headers:headers, body: body?JSON.stringify(body):'{}'});
  }

  private loadZoning(): void {
    this.spGet(this.listApi()+'/items?$select=Id,ParcelID,Zone,Floodplain,Jurisdiction,SplitGeoJSON&$top=5000').then((d:any)=>{
      const items=(d&&d.value)||[]; const m:any={};
      items.forEach((it:any)=>{ if(it.ParcelID&&it.Zone){ const e:any={zone:it.Zone,flood:!!it.Floodplain,id:it.Id,jur:it.Jurisdiction||'RBS'}; if(it.SplitGeoJSON){ try{ const arr=JSON.parse(it.SplitGeoJSON); if(arr&&arr.length){ e.split=true; e.pieces=arr; } }catch(x){} } m[pinKey(it.ParcelID)]=e; } });
      this.zoneByPin=m; this.restyleParcels(); this.buildSplitLayer();
    }).catch(()=>{ /* list missing / no access — zoning just stays empty */ });
  }

  private featPin(ft:any): string { const src=SOURCES.filter((s)=>s.id===ft.properties.__src)[0]||SOURCES[0]; return pinKey(pick(ft.properties, src.f.pin)); }
  private parcelStyle(ft:any): any {
    if(this.zoningView){
      const z=this.zoneByPin[this.featPin(ft)];
      if(z){ if(z.split) return {color:'#6b5300',weight:1,fillColor:'#000',fillOpacity:0.001}; const j=jurById(z.jur)||ZJURS[0]; const c=(j.colors&&j.colors[z.zone])||'#888'; return {color:'#6b5300',weight:1,fillColor:c,fillOpacity:0.55}; }
      // untagged lot in Zoning:View — keep the town overview clean: hide the outline until you zoom in for detail (still clickable for owner/deed)
      if(this.map && this.map.getZoom() < PARCEL_DETAIL_ZOOM) return {stroke:false,fill:true,fillColor:'#000',fillOpacity:0.001};
    }
    return {color:'#ffd24d',weight:1,fillColor:'#000',fillOpacity:0.001};
  }
  private restyleParcels(): void { try{ if(this.parcelLayer) this.parcelLayer.setStyle((ft:any)=>this.parcelStyle(ft)); }catch(e){} }

  private buildZPanel(): void {
    const el=this.domElement.querySelector('#zlegend') as any; if(!el) return;
    let h='<b>Zoning</b>';
    h+='<div class="ztag">Tag lots as: <select id="ztagjur"><option value="auto">Auto-detect</option>';
    ZJURS.forEach((j:any)=>{ if(j.taggable) h+='<option value="'+j.id+'">'+esc(j.name)+'</option>'; });
    h+='</select></div>';
    ZJURS.forEach((j:any)=>{ if(!j.taggable) return; h+='<div class="zjh">'+j.name+'</div>'; j.zones.forEach((z:string)=>{ h+='<div class="zi"><span class="zsw" style="background:'+j.colors[z]+'"></span>'+z+' &middot; '+j.names[z]+'</div>'; }); });
    h+='<div class="zdiv"></div><div class="zjh">Other layers</div>';
    h+='<div class="zrow"><label><input type="checkbox" id="zfema"'+(this._femaOn?' checked':'')+'> FEMA flood (NFHL)</label><span class="zacc exact">live</span><input type="range" min="20" max="100" value="'+(this.femaLayer&&this.femaLayer.options?Math.round(this.femaLayer.options.opacity*100):55)+'" id="zfemaop"></div>';
    h+='<div class="zrow"><label><input type="checkbox" id="zareas"'+(this._areasOn?' checked':'')+'> Drawn areas (historic dist.)</label></div>';
    h+='<button class="zbtn2" id="zdrawarea" style="margin-top:4px">Draw an area&hellip;</button>';
    h+='<div class="zdisc">Tagged lots are colored by their district. Use "Tag lots as" to lock a jurisdiction for edge lots. FEMA flood is the official 1% layer. Confirm zoning with the city/county.</div>';
    el.innerHTML=h;
    const self=this;
    const ts=el.querySelector('#ztagjur') as any; if(ts){ ts.value=this.tagJur; ts.addEventListener('change',function(e:any){ self.tagJur=e.target.value; }); }
    const fm=el.querySelector('#zfema') as any; if(fm) fm.addEventListener('change',function(e:any){ self._femaOn=!!e.target.checked; self.applyFema(); });
    const fo=el.querySelector('#zfemaop') as any; if(fo) fo.addEventListener('input',function(e:any){ if(self.femaLayer) self.femaLayer.setOpacity((+e.target.value)/100); });
    const ar=el.querySelector('#zareas') as any; if(ar) ar.addEventListener('change',function(e:any){ self._areasOn=!!e.target.checked; self.buildAreasLayer(); });
    const da=el.querySelector('#zdrawarea') as any; if(da) da.addEventListener('click',function(){ self.startAreaDraw(); });
  }

  private setZoningMode(v:string): void {
    this.zoningView = (v==='view'||v==='edit');
    this.zoningEdit = (v==='edit');
    const zl=this.domElement.querySelector('#zlegend') as any;
    if(zl) zl.style.display = this.zoningView ? 'block' : 'none';
    this.applyOverlays();
    this.restyleParcels();
    this.buildSplitLayer();
    this.applyFema();
    this.buildAreasLayer();
    if(v==='edit') this.setStatus('Zoning EDIT — set "Tag lots as" if needed, click a lot, choose its zone. Reference only.');
    else if(v==='view') this.setStatus('Zoning VIEW — tagged lots are colored by their district.');
    else this.setStatus('Zoning off.');
    this.map.closePopup();
  }

  private applyOverlays(): void {
    ZJURS.forEach((j:any)=>{ if(!j._layer) return; const show=this.zoningView && j._on; if(show){ if(!this.map.hasLayer(j._layer)){ j._layer.addTo(this.map); j._layer.bringToFront(); } } else if(this.map.hasLayer(j._layer)){ this.map.removeLayer(j._layer); } });
  }

  private openZonePicker(n:any, ll:any, feat?:any): void {
    const pin=pinKey(n.pin); if(!pin){ this.setStatus('This parcel has no ID — cannot tag it.'); return; }
    const j = (this.tagJur && this.tagJur!=='auto') ? jurById(this.tagJur) : (jurAt(ll) || nearestJur(ll));
    if(!j){ this.setStatus('No zoning jurisdiction available to tag.'); return; }
    this.zTarget={pin:pin, raw:String(n.pin).trim(), jur:j.id, ll:ll, ring:outerRing(feat&&feat.geometry)};
    const cur=this.zoneByPin[pin];
    const curTxt = cur? (cur.split&&cur.pieces? esc(cur.pieces.map((p:any)=>p.z||'blank').join(' / '))+' (split, '+esc(cur.jur)+')' : esc(cur.zone)+' ('+esc(cur.jur)+')'+(cur.flood?' + Floodplain':'')) : '—';
    let g=''; j.zones.forEach((z:string)=>{ g+='<button class="zbtn" data-act="zset" data-arg="'+z+'" style="background:'+j.colors[z]+'">'+z+'<small>'+j.names[z]+'</small></button>'; });
    const html='<div class="zp"><div class="zp-h">Set zoning &middot; '+esc(j.name)+'</div>'
      +'<div class="zp-pin">Parcel: <b>'+esc(n.pin)+'</b></div>'
      +'<div class="zp-cur">Current: <b>'+curTxt+'</b></div>'
      +'<div class="zp-grid">'+g+'</div>'
      +'<label class="zp-fl"><input type="checkbox" id="zpFlood"'+(cur&&cur.flood?' checked':'')+'> In 1% floodplain</label>'
      +'<div style="margin-top:4px"><button class="zp-clear" data-act="zsplitopen">Split lot&hellip;</button> <button class="zp-clear" data-act="zclear">Clear</button></div>'
      +'<div class="zp-note">'+esc(j.name)+(j.accuracy==='approx'?' overlay is approximate — verify boundary lots. ':' ')+'Reference only — not an official determination.</div></div>';
    L.popup({maxWidth:280,autoPanPadding:[24,24]}).setLatLng(ll).setContent(html).openOn(this.map);
  }

  private saveZone(zone:string): void {
    const t=this.zTarget; if(!t||!t.pin||!t.jur) return;
    const j=jurById(t.jur); if(!j||j.zones.indexOf(zone)<0) return;
    const fl=this.domElement.querySelector('#zpFlood') as any; const flood=!!(fl&&fl.checked);
    const cur=this.zoneByPin[t.pin];
    const done=(id:number)=>{ this.zoneByPin[t.pin]={zone:zone,flood:flood,id:id,jur:t.jur}; this.restyleParcels(); this.setStatus('Saved '+t.raw+' → '+zone+' ('+t.jur+')'+(flood?' + floodplain':'')); this.map.closePopup(); };
    if(cur && cur.id){
      this.spPost(this.listApi()+'/items('+cur.id+')', {Zone:zone,Floodplain:flood,Jurisdiction:t.jur}, {'X-HTTP-Method':'MERGE','IF-MATCH':'*'})
        .then((r:any)=>{ if(r.status>=200&&r.status<300) done(cur.id); else this.setStatus('Save failed ('+r.status+') — check list permissions'); })
        .catch((e:any)=>this.setStatus('Save failed: '+e));
    } else {
      this.spPost(this.listApi()+'/items', {Title:t.raw,ParcelID:t.raw,Jurisdiction:t.jur,Zone:zone,Floodplain:flood})
        .then((r:any)=>{ if(r.status>=200&&r.status<300) return r.json(); throw new Error('HTTP '+r.status); })
        .then((d:any)=>done(d&&d.Id))
        .catch((e:any)=>this.setStatus('Save failed: '+e));
    }
  }

  private clearZone(): void {
    const t=this.zTarget; if(!t||!t.pin) return; const cur=this.zoneByPin[t.pin];
    if(!cur||!cur.id){ this.map.closePopup(); return; }
    this.spPost(this.listApi()+'/items('+cur.id+')', null, {'X-HTTP-Method':'DELETE','IF-MATCH':'*'})
      .then((r:any)=>{ if(r.status>=200&&r.status<300){ delete this.zoneByPin[t.pin]; this.restyleParcels(); this.setStatus('Cleared zoning for '+t.raw); this.map.closePopup(); } else this.setStatus('Clear failed ('+r.status+')'); })
      .catch((e:any)=>this.setStatus('Clear failed: '+e));
  }

  // ======================= split-zoned lots =======================
  private jurColor(jur:string, z:string): string { const j=jurById(jur)||ZJURS[0]; return (j.colors&&j.colors[z])||'#888'; }

  private startSplitMenu(): void {
    const t=this.zTarget; if(!t||!t.ring||t.ring.length<3){ this.setStatus('This lot has no usable shape to split.'); return; }
    const jn=jurById(t.jur)?jurById(t.jur).name:t.jur;
    const html='<div class="zp"><div class="zp-h">Split lot &middot; '+esc(jn)+'</div>'
      +'<div class="zp-pin">Parcel: <b>'+esc(t.raw)+'</b></div>'
      +'<div class="sp-opts">'
      +'<button class="zbtn2" data-act="zsplit" data-arg="h2">Halve &mdash; top / bottom</button>'
      +'<button class="zbtn2" data-act="zsplit" data-arg="v2">Halve &mdash; left / right</button>'
      +'<button class="zbtn2" data-act="zsplit" data-arg="h3">Thirds &mdash; rows</button>'
      +'<button class="zbtn2" data-act="zsplit" data-arg="v3">Thirds &mdash; columns</button>'
      +'<button class="zbtn2" data-act="zsplit" data-arg="draw1">Draw 1 line &rarr; 2 parts</button>'
      +'<button class="zbtn2" data-act="zsplit" data-arg="draw2">Draw 2 lines &rarr; 3 parts</button>'
      +'</div><button class="zp-clear" data-act="zsplitcancel">Cancel</button>'
      +'<div class="zp-note">Then pick a zone for each numbered part. Approximate &mdash; reference only.</div></div>';
    L.popup({maxWidth:240,autoPanPadding:[24,24]}).setLatLng(t.ll||this.map.getCenter()).setContent(html).openOn(this.map);
  }

  private beginSplit(kind:string): void {
    const t=this.zTarget; if(!t||!t.ring) return;
    const id=this.zoneByPin[t.pin]?this.zoneByPin[t.pin].id:null;
    this.splitState={pin:t.pin, raw:t.raw, jur:t.jur, id:id, rings:[ringOpen(t.ring)], pieces:[], pts:[], lines:0, ll:t.ll};
    this.clearSplitPreview();
    if(kind==='draw1'||kind==='draw2'){
      this.splitState.lines=(kind==='draw2'?2:1);
      this.setStatus('Cut 1: click two points on the map for the line.');
      this._splitClick=(e:any)=>this.onSplitClick(e); this.map.on('click',this._splitClick);
      this._splitDrawPopup = L.popup({maxWidth:210,closeOnClick:false,autoClose:false}).setLatLng(t.ll||this.map.getCenter())
        .setContent('<div class="zp"><div class="zp-h">Drawing cut line'+(this.splitState.lines>1?'s':'')+'</div><div class="zp-note">Click <b>two points</b> on the map for each line ('+this.splitState.lines+' line'+(this.splitState.lines>1?'s':'')+'). Each extends across the lot.</div><button class="zp-clear" data-act="zsplitcancel">Cancel</button></div>')
        .openOn(this.map);
      return;
    }
    const b=ringBounds(this.splitState.rings[0]); const eps=1e-4; let rings=this.splitState.rings;
    if(kind==='h2'){ const cy=(b[1]+b[3])/2; rings=splitByLine(rings,[b[0]-eps,cy],[b[2]+eps,cy]); }
    else if(kind==='v2'){ const cx=(b[0]+b[2])/2; rings=splitByLine(rings,[cx,b[1]-eps],[cx,b[3]+eps]); }
    else if(kind==='h3'){ const y1=b[1]+(b[3]-b[1])/3,y2=b[1]+2*(b[3]-b[1])/3; rings=splitByLine(rings,[b[0]-eps,y1],[b[2]+eps,y1]); rings=splitByLine(rings,[b[0]-eps,y2],[b[2]+eps,y2]); }
    else if(kind==='v3'){ const x1=b[0]+(b[2]-b[0])/3,x2=b[0]+2*(b[2]-b[0])/3; rings=splitByLine(rings,[x1,b[1]-eps],[x1,b[3]+eps]); rings=splitByLine(rings,[x2,b[1]-eps],[x2,b[3]+eps]); }
    this.splitState.rings=rings; this.finishSplitGeometry();
  }

  private onSplitClick(e:any): void {
    const st=this.splitState; if(!st) return;
    st.pts.push([e.latlng.lng, e.latlng.lat]);
    this.splitMarkers.push(L.circleMarker(e.latlng,{radius:4,color:'#1d4ed8',weight:2,fill:true,fillColor:'#1d4ed8',fillOpacity:1,pane:'zsplit',interactive:false}).addTo(this.map));
    if(st.pts.length%2===0){
      const A=st.pts[st.pts.length-2], B=st.pts[st.pts.length-1];
      st.rings=splitByLine(st.rings, A, B);
      this.splitMarkers.push(L.polyline([[A[1],A[0]],[B[1],B[0]]],{color:'#1d4ed8',weight:2,dashArray:'5,4',pane:'zsplit',interactive:false}).addTo(this.map));
      const done=st.pts.length/2;
      if(done>=st.lines){ if(this._splitClick){ this.map.off('click',this._splitClick); this._splitClick=null; } this.finishSplitGeometry(); }
      else this.setStatus('Cut '+(done+1)+': click two points for the next line.');
    }
  }

  private finishSplitGeometry(): void {
    const st=this.splitState; if(!st) return;
    if(this._splitDrawPopup){ try{ this.map.closePopup(this._splitDrawPopup); }catch(e){} this._splitDrawPopup=null; }
    if(!st.rings || st.rings.length<2){ this.setStatus('That line did not divide the lot — try again.'); this.cancelSplit(); return; }
    st.pieces=st.rings.map((r:any)=>({r:r, z:null}));
    this.renderSplitPreview(); this.openLabelPopup();
  }

  private renderSplitPreview(): void {
    this.clearSplitPreview();
    const st=this.splitState; if(!st) return;
    st.pieces.forEach((p:any,i:number)=>{
      const latlngs=p.r.map((c:any)=>[c[1],c[0]]);
      const fill = p.z? this.jurColor(st.jur,p.z) : '#94a3b8';
      this.splitTmp.push(L.polygon(latlngs,{color:'#1d4ed8',weight:2,dashArray:'4,3',fillColor:fill,fillOpacity:p.z?0.55:0.2,pane:'zsplit',interactive:false}).addTo(this.map));
      const c=centroid(ringOpen(p.r));
      this.splitMarkers.push(L.marker([c[1],c[0]],{pane:'zsplit',interactive:false,icon:L.divIcon({className:'znum',html:String(i+1),iconSize:[18,18]})}).addTo(this.map));
    });
  }

  private openLabelPopup(): void {
    const st=this.splitState; if(!st) return; const j=jurById(st.jur)||ZJURS[0];
    let body='';
    st.pieces.forEach((p:any,i:number)=>{
      let g=''; j.zones.forEach((z:string)=>{ g+='<button class="zbtns" data-act="zpz" data-arg="'+i+'|'+z+'" style="background:'+j.colors[z]+'">'+z+'</button>'; });
      g+='<button class="zbtns zbnone" data-act="zpz" data-arg="'+i+'|-">None</button>';
      body+='<div class="spc"><div class="spc-h"><b>Piece '+(i+1)+'</b><span id="spc'+i+'">'+(p.z?esc(p.z):'blank')+'</span></div><div class="spc-g">'+g+'</div></div>';
    });
    const html='<div class="zp"><div class="zp-h">Label split &middot; '+esc(j.name)+'</div>'+body
      +'<div style="margin-top:7px"><button class="zp-save" data-act="zsplitsave">Save split</button> <button class="zp-clear" data-act="zsplitcancel">Cancel</button></div>'
      +'<div class="zp-note">Pick a zone for each numbered piece. Approximate &mdash; reference only.</div></div>';
    L.popup({maxWidth:250,autoPanPadding:[24,24]}).setLatLng(st.ll||this.map.getCenter()).setContent(html).openOn(this.map);
  }

  private setPieceZone(i:number, zone:string): void {
    const st=this.splitState; if(!st||!st.pieces[i]) return;
    if(zone==='-'){ st.pieces[i].z=null; }
    else { const j=jurById(st.jur); if(!j||j.zones.indexOf(zone)<0) return; st.pieces[i].z=zone; }
    const sp=this.domElement.querySelector('#spc'+i) as any; if(sp) sp.textContent=st.pieces[i].z||'blank';
    this.renderSplitPreview();
  }

  private saveSplit(): void {
    const st=this.splitState; if(!st) return;
    let firstZone:any=null; for(let i=0;i<st.pieces.length;i++){ if(st.pieces[i].z){ firstZone=st.pieces[i].z; break; } }
    if(!firstZone){ this.setStatus('Give at least one piece a zone (or use Clear to blank the whole lot).'); return; }
    const arr=st.pieces.map((p:any)=>({z:p.z||null, r:p.r})); const json=JSON.stringify(arr);
    const done=(id:number)=>{ this.zoneByPin[st.pin]={split:true,pieces:arr,zone:firstZone,flood:false,id:id,jur:st.jur}; this.splitState=null; this.clearSplitPreview(); this.buildSplitLayer(); this.restyleParcels(); this.setStatus('Saved split for '+st.raw+' ('+arr.map((p:any)=>p.z||'blank').join(' / ')+')'); this.map.closePopup(); };
    const body:any={Zone:firstZone, Jurisdiction:st.jur, SplitGeoJSON:json, Floodplain:false};
    if(st.id){ this.spPost(this.listApi()+'/items('+st.id+')', body, {'X-HTTP-Method':'MERGE','IF-MATCH':'*'}).then((r:any)=>{ if(r.status>=200&&r.status<300) done(st.id); else this.setStatus('Save failed ('+r.status+')'); }).catch((e:any)=>this.setStatus('Save failed: '+e)); }
    else { body.Title=st.raw; body.ParcelID=st.raw; this.spPost(this.listApi()+'/items', body).then((r:any)=>{ if(r.status>=200&&r.status<300) return r.json(); throw new Error('HTTP '+r.status); }).then((d:any)=>done(d&&d.Id)).catch((e:any)=>this.setStatus('Save failed: '+e)); }
  }

  private cancelSplit(): void { if(this._splitClick){ this.map.off('click',this._splitClick); this._splitClick=null; } if(this._splitDrawPopup){ try{ this.map.closePopup(this._splitDrawPopup); }catch(e){} this._splitDrawPopup=null; } this.clearSplitPreview(); this.splitState=null; this.map.closePopup(); this.setStatus('Split cancelled.'); }

  private clearSplitPreview(): void { const a=this.splitTmp||[]; for(let i=0;i<a.length;i++){ try{this.map.removeLayer(a[i]);}catch(e){} } this.splitTmp=[]; const m=this.splitMarkers||[]; for(let i=0;i<m.length;i++){ try{this.map.removeLayer(m[i]);}catch(e){} } this.splitMarkers=[]; }

  private buildSplitLayer(): void {
    if(!this.splitLayer) return; this.splitLayer.clearLayers();
    if(!this.zoningView) return; const m=this.zoneByPin; const self=this;
    Object.keys(m).forEach((pin)=>{ const z=m[pin]; if(z&&z.split&&z.pieces){ z.pieces.forEach((p:any)=>{ if(!p.z||!p.r||p.r.length<3) return; const latlngs=p.r.map((c:any)=>[c[1],c[0]]); self.splitLayer.addLayer(L.polygon(latlngs,{color:'#6b5300',weight:1,fillColor:self.jurColor(z.jur,p.z),fillOpacity:0.55,interactive:false,pane:'zsplit'})); }); } });
  }

  // ======================= FEMA flood + drawn areas =======================
  private apiList(title:string): string { return this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + title.replace(/'/g,"''") + "')"; }

  private applyFema(): void { if(!this.femaLayer) return; if(this._femaOn && this.zoningView){ if(!this.map.hasLayer(this.femaLayer)){ this.femaLayer.addTo(this.map); } } else if(this.map.hasLayer(this.femaLayer)){ this.map.removeLayer(this.femaLayer); } }

  private loadAreas(): void {
    this.spGet(this.apiList('DLS Map Areas')+'/items?$select=Id,AreaType,Jurisdiction,AreaGeoJSON&$top=2000').then((d:any)=>{
      const items=(d&&d.value)||[]; const arr:any[]=[];
      items.forEach((it:any)=>{ let ring=null; if(it.AreaGeoJSON){ try{ ring=JSON.parse(it.AreaGeoJSON); }catch(e){} } if(ring&&ring.length>=3) arr.push({id:it.Id,type:it.AreaType,jur:it.Jurisdiction,ring:ring}); });
      this.areas=arr; this.buildAreasLayer();
    }).catch(()=>{ /* list missing / no access */ });
  }

  private buildAreasLayer(): void {
    if(!this.areasLayer) return; this.areasLayer.clearLayers();
    if(!this.zoningView || !this._areasOn) return; const self=this;
    (this.areas||[]).forEach((a:any)=>{
      const latlngs=a.ring.map((c:any)=>[c[1],c[0]]); const isHist=(a.type==='Historic District'); const isFlood=(a.type==='Floodplain');
      const poly=L.polygon(latlngs,{renderer:self._areasRenderer,pane:'areas',color:isFlood?'#1d4ed8':'#333',weight:isFlood?3:1.5,fill:isHist,fillColor:'#000',fillOpacity:0,interactive:false});
      poly.addTo(self.areasLayer);
      if(isHist){ const patch=function(){ try{ if(poly._path){ poly._path.setAttribute('fill','url(#dls-hatch)'); poly._path.setAttribute('fill-opacity','1'); } }catch(e){} }; patch(); poly.on('add',patch); }
    });
  }

  private startAreaDraw(): void {
    this.clearAreaDraw(); this.areaState={pts:[]};
    this._areaClick=(e:any)=>this.onAreaClick(e); this.map.on('click',this._areaClick);
    this.setStatus('Draw area: click points around the boundary, then Finish.');
    L.popup({closeOnClick:false,autoClose:false,maxWidth:215}).setLatLng(this.map.getCenter()).setContent('<div class="zp"><div class="zp-h">Draw an area</div><div class="zp-note">Click points around the boundary on the map (3+). Then Finish.</div><button class="zp-save" data-act="zareafinish">Finish</button> <button class="zp-clear" data-act="zareacancel">Cancel</button></div>').openOn(this.map);
  }

  private onAreaClick(e:any): void {
    const st=this.areaState; if(!st) return; st.pts.push([e.latlng.lng,e.latlng.lat]);
    this.areaMarkers.push(L.circleMarker(e.latlng,{renderer:this._areasRenderer,pane:'areas',radius:3,color:'#7c3aed',weight:2,fill:true,fillColor:'#7c3aed',fillOpacity:1,interactive:false}).addTo(this.map));
    const latlngs=st.pts.map((c:any)=>[c[1],c[0]]);
    if(this.areaLine){ try{this.map.removeLayer(this.areaLine);}catch(e2){} }
    this.areaLine=L.polyline(latlngs,{renderer:this._areasRenderer,pane:'areas',color:'#7c3aed',weight:2,dashArray:'4,3',interactive:false}).addTo(this.map);
    this.setStatus('Area: '+st.pts.length+' point(s). Click more, then Finish.');
  }

  private finishAreaDraw(): void {
    const st=this.areaState; if(!st) return; if(st.pts.length<3){ this.setStatus('Add at least 3 points first.'); return; }
    this.map.closePopup();
    const html='<div class="zp"><div class="zp-h">Save area</div><div class="zp-note">What is this area?</div>'
      +'<div class="sp-opts"><button class="zbtn2" data-act="zareasave" data-arg="Historic District">Historic District (hatch)</button>'
      +'<button class="zbtn2" data-act="zareasave" data-arg="Floodplain">Floodplain (blue outline)</button>'
      +'<button class="zbtn2" data-act="zareasave" data-arg="Other">Other</button></div>'
      +'<button class="zp-clear" data-act="zareacancel">Cancel</button></div>';
    L.popup({closeOnClick:false,autoClose:false,maxWidth:230}).setLatLng(this.map.getCenter()).setContent(html).openOn(this.map);
  }

  private saveArea(type:string): void {
    const st=this.areaState; if(!st||st.pts.length<3) return;
    const ring=st.pts.slice(); const c=centroid(ring); const jj=jurAt({lat:c[1],lng:c[0]})||nearestJur({lat:c[1],lng:c[0]}); const jur=jj?jj.id:'RBS';
    const body:any={Title:type+' ('+jur+')', AreaType:type, Jurisdiction:jur, AreaGeoJSON:JSON.stringify(ring)};
    this.spPost(this.apiList('DLS Map Areas')+'/items', body).then((r:any)=>{ if(r.status>=200&&r.status<300) return r.json(); throw new Error('HTTP '+r.status); })
      .then((d:any)=>{ if(!this.areas) this.areas=[]; this.areas.push({id:d&&d.Id,type:type,jur:jur,ring:ring}); this.clearAreaDraw(); this.map.closePopup(); this._areasOn=true; const ar=this.domElement.querySelector('#zareas') as any; if(ar) ar.checked=true; this.buildAreasLayer(); this.setStatus('Saved '+type+'.'); })
      .catch((e:any)=>this.setStatus('Save area failed: '+e));
  }

  private cancelAreaDraw(): void { this.clearAreaDraw(); this.map.closePopup(); this.setStatus('Area drawing cancelled.'); }

  private clearAreaDraw(): void {
    if(this._areaClick){ this.map.off('click',this._areaClick); this._areaClick=null; }
    const m=this.areaMarkers||[]; for(let i=0;i<m.length;i++){ try{this.map.removeLayer(m[i]);}catch(e){} } this.areaMarkers=[];
    if(this.areaLine){ try{this.map.removeLayer(this.areaLine);}catch(e){} this.areaLine=null; }
    this.areaState=null;
  }

  private toggleFs(): void {
    const el=this.domElement.querySelector('.dls-pm') as any; if(!el) return;
    const on=!el.classList.contains('fs'); if(on){ el.classList.add('fs'); } else { el.classList.remove('fs'); }
    const b=this.domElement.querySelector('#fs') as any; if(b) b.textContent = on?'Exit full screen':'Full screen';
    setTimeout(()=>{ try{ this.map.invalidateSize(); }catch(e){} }, 60);
  }

  private setStatus(t:string): void { const el=this.domElement.querySelector('#status'); if(el) el.textContent=t; }

  protected get dataVersion(): Version { return Version.parse('1.0'); }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages:[{ header:{description:'Property & Deed Map settings'}, groups:[{ groupName:'Settings', groupFields:[
      PropertyPaneTextField('title',{label:'Title'}),
      PropertyPaneTextField('workerUrl',{label:'Deed Worker URL (Cloudflare)'}),
      PropertyPaneTextField('zoneListTitle',{label:'Zoning list title'}),
      PropertyPaneTextField('zoningAssetBase',{label:'Zoning overlays folder URL'})
    ]}]}]};
  }
}
