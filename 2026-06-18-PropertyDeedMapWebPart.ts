/* eslint-disable */
// ============================================================================
// DLS Property & Deed Map â SPFx client-side web part (framework: none)
// Port of the standalone 2026-06-18 map. Leaflet is BUNDLED (imported below),
// NOT from a CDN (this tenant blocks external scripts); the Leaflet CSS is HARDCODED in
// the LEAFLET_CSS const below (a build-time inline step proved unreliable). The tenant ALLOWS external
// fetch (TN/KY parcel services, ArcGIS geocoder, the deed Worker) and external
// <img> (Esri tiles), so those work as-is.
//
// Features: live parcels per viewport (TN statewide 86 co. + Davidson/Hamilton/
// Rutherford/Montgomery/Williamson/Shelby + KY Simpson/Pulaski/Warren); owner /
// address / parcel search; click a parcel for owner/address/parcel + a book/page-
// first deed lookup (Cloudflare Worker â latest warranty-deed book/page â auto-run
// TitleSearcher; Sumner/Trousdale â US Title Search handoff; owner-name fallback).
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
// build-time inline step proved unreliable â so the CSS lives here so it can never be dropped).
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
  { id:'tn', label:'TN â Statewide (86 counties)', state:'TN',
    url:'https://services1.arcgis.com/YuVBSS7Y1of2Qud1/arcgis/rest/services/Tennessee_Property_Boundaries_Public_Use/FeatureServer/0/query',
    bbox:[-90.45,34.94,-81.60,36.72], countyField:'COUNTY_NAME', where:'1=1',
    f:{pin:['PARCELID'],owner:['OWNER'],owner2:['OWNER2'],address:['ADDRESS'],subdiv:['SUBDIV'],lot:['LOT'],acres:['DEEDAC'],assr:['LINK_TPV','LINK_TPAD'],tpad:['LINK_TPAD'],gislinkf:['GISLINK'],parcelno:['PARCEL']},
    search:{owner:'OWNER',address:'ADDRESS',parcel:'PARCELID'} },
  { id:'davidson', label:'TN â Davidson / Nashville', state:'TN', county:'DAVIDSON',
    url:'https://maps.nashville.gov/arcgis/rest/services/Cadastral/Parcels/MapServer/0/query',
    bbox:[-87.06,35.96,-86.51,36.41], where:"FeatureType IS NULL OR FeatureType<>'Unit'",
    f:{pin:['APN','STANPAR'],owner:['Owner'],address:['PropAddr'],mail:['OwnAddr1'],acres:['Acres','DeededAcreage'],zoning:['Zoning'],deedref:['OwnInstr']},
    search:{owner:'Owner',address:'PropAddr',parcel:'APN'} },
  { id:'hamilton', label:'TN â Hamilton / Chattanooga', state:'TN', county:'HAMILTON',
    url:'https://mapsdev.hamiltontn.gov/hcwa03/rest/services/Live_Parcels/MapServer/0/query',
    bbox:[-85.55,34.98,-84.96,35.46], where:"OWNERNAME1<>'Update in Progress'",
    f:{pin:['PARCEL','TAX_MAP_NO','GISLINK'],owner:['OWNERNAME1'],owner2:['OWNERNAME2'],address:['ADDRESS']},
    search:{owner:'OWNERNAME1',address:'ADDRESS',parcel:'PARCEL'} },
  { id:'rutherford', label:'TN â Rutherford / Murfreesboro', state:'TN', county:'RUTHERFORD',
    url:'https://services.arcgis.com/36I6IHIdr660pAyH/ArcGIS/rest/services/ParcelsCAMA1/FeatureServer/0/query',
    bbox:[-86.62,35.64,-86.03,36.05], where:'GISLINK IS NOT NULL',
    f:{pin:['ParcelID','GISLINK'],owner:['Owner1'],owner2:['Owner2'],address:['FormattedLocation','STREETADDRESS'],mail:['MailingAddress'],subdiv:['SUBDIVISION'],lot:['LOT'],acres:['CALCACRES','DEEDACRES'],zoning:['ZONING'],legalref:['LegalReference']},
    search:{owner:'Owner1',address:'FormattedLocation',parcel:'ParcelID'} },
  { id:'montgomery', label:'TN â Montgomery / Clarksville', state:'TN', county:'MONTGOMERY',
    url:'https://apnsgis4.apsu.edu/arcgis/rest/services/CMCGIS/MontViewer/FeatureServer/2/query',
    bbox:[-87.50,36.39,-87.00,36.71], where:'1=1',
    f:{pin:['parcelid','gislink'],owner:['owner'],owner2:['owner2'],address:['propertyaddress']},
    search:{owner:'owner',address:'propertyaddress',parcel:'parcelid'} },
  { id:'williamson', label:'TN â Williamson / Franklin', state:'TN', county:'WILLIAMSON',
    url:'http://arcgis2.williamson-tn.org/arcgis/rest/services/IDT/DataPull/MapServer/4/query',
    bbox:[-87.18,35.68,-86.68,36.08], where:'1=1', note:'HTTP-only host â blocked from an HTTPS page (mixed content)',
    f:{pin:['parcel_id','GISLINK'],owner:['owner1'],owner2:['owner2'],address:['ADDRESS']},
    search:{owner:'owner1',address:'ADDRESS',parcel:'parcel_id'} },
  { id:'shelby', label:'TN â Shelby / Memphis', state:'TN', county:'SHELBY',
    url:'https://gis.shelbycountytn.gov/public/rest/services/Parcel/CERT_Parcel/MapServer/0/query',
    bbox:[-90.31,34.94,-89.64,35.42], where:'1=1', note:'their DB connection was intermittent',
    f:{pin:['PARCELID','PARID'],owner:['OWNER'],owner2:['OWNER_EXT'],address:['PAR_ADDR1'],mail:['OWN_ADDR1']},
    search:{owner:'OWNER',address:'PAR_ADDR1',parcel:'PARCELID'} },
  { id:'ky_simpson', label:'KY â Simpson / Franklin', state:'KY', county:'SIMPSON',
    url:'https://services8.arcgis.com/D3RgmiBYTvYcNK2j/arcgis/rest/services/Parcel2026view/FeatureServer/0/query',
    bbox:[-86.78,36.62,-86.42,36.87], where:"PIDN<>' '",
    f:{pin:['PIDN'],owner:['NAME'],address:['Property_L'],mail:['Address_Li'],acres:['ACRES'],deedref:['DEED']},
    search:{owner:'NAME',address:'Property_L',parcel:'PIDN'} },
  { id:'ky_pulaski', label:'KY â Pulaski / Somerset', state:'KY', county:'PULASKI',
    url:'https://services5.arcgis.com/cnJiyVVCFyUslPPa/arcgis/rest/services/ParcelUpdate_2026/FeatureServer/2/query',
    bbox:[-84.82,36.91,-84.29,37.29], where:"parcel_id<>' '",
    f:{pin:['parcel_id','Parc_lbl','Account'],owner:['owner1'],owner2:['owner2'],address:['prop_stree'],mail:['own_street'],acres:['legal_acre'],deedBook:['deed_book'],deedPage:['deed_page']},
    search:{owner:'owner1',address:'prop_stree',parcel:'parcel_id'} },
  { id:'ky_warren', label:'KY â Warren / Bowling Green', state:'KY', county:'WARREN',
    url:'https://webgis.bgky.org/server/rest/services/CCPC/CCPC_Parcels/MapServer/0/query',
    bbox:[-86.61,36.77,-86.26,37.11], where:'1=1', ownerWithheld:true,
    f:{pin:['PVA_PARCEL'],address:['ADDRESS'],subdiv:['SUBNAME'],lot:['LOT_NUMBER'],acres:['ACRES'],zoning:['ZONING']},
    search:{address:'ADDRESS',parcel:'PVA_PARCEL'} }
];

const TN_COUNTIES=['Anderson','Bedford','Benton','Bledsoe','Blount','Bradley','Campbell','Cannon','Carroll','Carter','Cheatham','Claiborne','Clay','Cocke','Coffee','Crockett','Cumberland','Decatur','DeKalb','Dickson','Dyer','Fayette','Fentress','Franklin','Gibson','Giles','Grainger','Greene','Grundy','Hamblen','Hancock','Hardeman','Hardin','Hawkins','Haywood','Henderson','Henry','Houston','Humphreys','Jackson','Jefferson','Johnson','Lake','Lauderdale','Lawrence','Lewis','Lincoln','Loudon','Macon','Madison','Marion','Marshall','Maury','McMinn','McNairy','Meigs','Monroe','Moore','Morgan','Obion','Overton','Perry','Pickett','Polk','Putnam','Rhea','Roane','Robertson','Scott','Sequatchie','Sevier','Smith','Stewart','Sullivan','Sumner','Tipton','Trousdale','Unicoi','Union','Van Buren','Warren','Washington','Wayne','Weakley','White','Wilson'];
const MINZOOM = 15;
const LABELZOOM = 17;

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
function wkNorm5(p:any){ p=(p==null?'':String(p)).trim().split(/\s+/)[0]; if(!p) return ''; const a=p.split('.'); const i=a[0].replace(/[^0-9]/g,''); if(i==='') return ''; const dec=((a[1]||'').replace(/[^0-9]/g,'')+'00').substring(0,2); const ii=('000'+i); return ii.substring(ii.length-3)+dec; }
function wkNormMap(m:any){ m=(m==null?'':String(m)).trim().split(/\s+/)[0].toUpperCase(); const mm=m.match(/^([0-9]+)([A-Z]?)$/); if(mm){ const d=('000'+mm[1]); return d.substring(d.length-3)+mm[2]; } return m; }
const WK_TN_SVC='https://services1.arcgis.com/YuVBSS7Y1of2Qud1/arcgis/rest/services/Tennessee_Property_Boundaries_Public_Use/FeatureServer/0/query';
function jurById(id:any){ for(let i=0;i<ZJURS.length;i++){ if(ZJURS[i].id===id) return ZJURS[i]; } return null; }
function jurAt(ll:any){ let best:any=null, ba=Infinity; ZJURS.forEach((j:any)=>{ if(!j.taggable) return; const b=j.bounds; if(ll.lat>=b[0][0]&&ll.lat<=b[1][0]&&ll.lng>=b[0][1]&&ll.lng<=b[1][1]){ const area=(b[1][0]-b[0][0])*(b[1][1]-b[0][1]); if(area<ba){ ba=area; best=j; } } }); return best; }
function nearestJur(ll:any){ let best:any=null, bd=Infinity; ZJURS.forEach((j:any)=>{ if(!j.taggable) return; const b=j.bounds; const cy=(b[0][0]+b[1][0])/2, cx=(b[0][1]+b[1][1])/2; const d=(ll.lat-cy)*(ll.lat-cy)+(ll.lng-cx)*(ll.lng-cx); if(d<bd){ bd=d; best=j; } }); return best; }

// ---- straight-line parcel splitting (half-plane clipping; no library). Rings are [[lng,lat],...]. ----
function ringOpen(r:any){ if(!r||r.length<2) return r||[]; const a=r.slice(); if(a.length>1 && a[0][0]===a[a.length-1][0] && a[0][1]===a[a.length-1][1]) a.pop(); return a; }
function ringBounds(r:any){ let mnx=Infinity,mny=Infinity,mxx=-Infinity,mxy=-Infinity; for(let i=0;i<r.length;i++){ const p=r[i]; if(p[0]<mnx)mnx=p[0]; if(p[0]>mxx)mxx=p[0]; if(p[1]<mny)mny=p[1]; if(p[1]>mxy)mxy=p[1]; } return [mnx,mny,mxx,mxy]; }
function sideOf(P:any,A:any,B:any){ return (B[0]-A[0])*(P[1]-A[1])-(B[1]-A[1])*(P[0]-A[0]); }
// Robust split of a simple polygon ring by an INFINITE line (handles concave / irregular lots â no
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

// ---- UCDD (Upper Cumberland Development District) official zoning overlay (added 2026-06-27) ----
var UCDD_BASE='https://services1.arcgis.com/EMZFDxQzNQloLbAf/arcgis/rest/services';
var UCDD_MINZOOM=14;
var UCDD_ZONING:any[]=[
  {key:'smith',label:'Smith County',service:'Smith_Zoning',layer:0,field:'ZONE',bbox:[-86.2205,36.0193,-85.6925,36.4668]},
  {key:'smithville',label:'Smithville',service:'Smithville_Zoning',layer:0,field:'Zone',bbox:[-85.9373,35.9095,-85.6961,36.0233]},
  {key:'south_carthage',label:'South Carthage',service:'South_Carthage_Zoning',layer:0,field:'Zone',bbox:[-86.0207,36.2134,-85.9001,36.2701]},
  {key:'gordonsville',label:'Gordonsville',service:'Zoning_Gordonsville241011',layer:0,field:'ZONE',bbox:[-85.9930,36.1555,-85.8710,36.2115]},
  {key:'algood',label:'Algood',service:'Zoning_240731',layer:0,field:'Zoning',bbox:[-85.4800,36.1684,-85.4026,36.2297]},
  {key:'baxter',label:'Baxter',service:'Baxter_Zoning_Update_251118',layer:0,field:'Zone_Curre',bbox:[-85.7435,36.1162,-85.5370,36.1870]},
  {key:'livingston',label:'Livingston',service:'Livingston_Zoning',layer:0,field:'Zone',bbox:[-85.3853,36.3566,-85.2633,36.4125]},
  {key:'morrison',label:'Morrison',service:'Morrison_Zoning',layer:0,field:'Zone',bbox:[-86.0024,35.5583,-85.7955,35.6726]},
  {key:'cannon',label:'Cannon County',service:'Cannon_Zoning_250708',layer:0,field:'Zone',bbox:[-86.7203,35.5340,-85.4377,36.0963]},
  {key:'monterey',label:'Monterey',service:'Monterey_Zoning_WFL1',layer:3,field:'Zone_Curre',bbox:[-85.2918,36.1273,-85.2308,36.1554]},
  {key:'spencer',label:'Spencer',service:'Spencer_Zoning_260205',layer:0,field:'ZONE',bbox:[-85.5568,35.6749,-85.3449,35.7994]},
  {key:'sumner_co',label:'Sumner County (unincorp.)',group:'sumner',url:'https://services5.arcgis.com/I2fcoygeHRmx0WzS/arcgis/rest/services/PSALayers/FeatureServer/20',field:'ZONING',bbox:[-86.76,36.24,-86.20,36.66]},
  {key:'gallatin',label:'Gallatin',group:'sumner',url:'https://arcweb.gallatin-tn.gov/arcgis/rest/services/Features/Zoning/FeatureServer/0',field:'ZONING',bbox:[-86.58,36.30,-86.40,36.46]},
  {key:'hendersonville',label:'Hendersonville',group:'sumner',url:'https://gis.hvilletn.org/gisapi/rest/services/HvilleTN_Zoning_UPDATED/MapServer/3',field:'ZONECODE',bbox:[-86.70,36.26,-86.54,36.40]},
  {key:'whitehouse',label:'White House',group:'sumner',url:'https://gis.cityofwhitehouse.com/arcgis/rest/services/WhiteHouseTN_Zoning/FeatureServer/3',field:'ZONECLASS',bbox:[-86.78,36.43,-86.60,36.53]},
  {key:'portland',label:'Portland',group:'sumner',url:'https://services3.arcgis.com/pXGyp7DHTIE4RXOJ/arcgis/rest/services/Zoning_Portland/FeatureServer/0',field:'ZQ45T',bbox:[-86.55,36.55,-86.47,36.61]},
  {key:'millersville',label:'Millersville',group:'sumner',url:'https://services.arcgis.com/jcrrmnzMsBOEAFJp/arcgis/rest/services/Millersville_Zoning_view/FeatureServer/5',field:'ZONE_2021',bbox:[-86.87,36.34,-86.68,36.47]},
  {key:'goodlettsville',label:'Goodlettsville',group:'sumner',url:'https://services8.arcgis.com/qSbT66zM7qttH0fv/arcgis/rest/services/ZONINGARGISMAP/FeatureServer/2',field:'ZONECLASS',bbox:[-86.75,36.27,-86.65,36.39]},
  {key:'westmoreland',label:'Westmoreland',group:'sumner',url:'https://services3.arcgis.com/pXGyp7DHTIE4RXOJ/arcgis/rest/services/Westmoreland_Zoning/FeatureServer/0',field:'Zone_Curre',bbox:[-86.31,36.53,-86.21,36.61]}
];
var UCDD_COLORS:any={
  smith:{'R-1':[242,223,12],'I-1':[142,4,189],'A-1':[33,222,36],'C-1':[255,0,8],'CRT':[255,136,0],'GRD':[135,72,4],'R-C':[255,0,217],'SCA':[3,255,217]},
  smithville:{'C-1':[255,0,8],'C-2':[138,6,6],'I-1':[181,9,175],'R-1':[240,208,5],'R-2':[252,146,31]},
  south_carthage:{'C-1':[255,0,8],'C-2':[94,7,10],'I-1':[120,5,173],'M-2':[11,214,204],'R-1':[247,232,17],'R-2':[255,171,3]},
  gordonsville:{'A-1':[29,209,13],'C-1':[255,0,4],'C-2':[145,1,4],'I-1':[245,5,245],'I-2':[127,0,255],'R-1':[255,230,0],'R-2':[255,153,0],'M-1':[0,234,255]},
  algood:{'R-1':[255,235,10],'R-2':[252,173,0],'R-3':[0,182,242],'R-D':[23,130,4],'C-A':[255,17,0],'C-B':[148,4,4],'C-C':[0,240,52],'I-1':[241,7,245],'I-2':[169,9,237]},
  baxter:{'C-1':[255,166,167],'C-2':[242,0,4],'CBD':[161,26,28],'I-1':[237,43,182],'R-1':[255,242,0],'R-2':[39,245,12],'R-3':[20,158,206],'R-M':[242,126,10]},
  livingston:{'C-1':[255,0,8],'C-2':[255,154,140],'C-3':[143,0,5],'C-M':[255,145,0],'I-1':[154,47,247],'I-2':[227,39,202],'R-1':[235,231,19],'R-2':[28,230,21]},
  morrison:{'A-1':[28,230,21],'C-1':[255,0,8],'C-2':[115,31,34],'I-1':[190,15,209],'R-1':[242,227,12],'R-2':[250,160,15]},
  cannon:{'A-1':[255,238,5],'C-1':[75,199,8],'I-1':[206,29,209],'R-1':[227,5,8]},
  spencer:{'A-1':[130,204,2],'C-1':[255,0,0],'C-2':[166,3,3],'I-1':[174,0,255],'R-1':[255,234,0],'R-2':[245,136,2]},
  monterey:{'C-1':[255,190,190],'C-2':[255,0,0],'I-1':[169,0,230],'R-1':[255,255,0],'R-2':[230,152,0],'R-R':[163,255,115]},
  sumner_co:{'AGRICULTURAL RESERVE (AR)':[38,115,0],'COMMERCIAL GENERAL (CG)':[255,190,190],'COMMERCIAL NEIGHBORHOOD (CN)':[215,158,158],'COMMERCIAL SERVICES (CS)':[255,211,127],'HEAVY INDUSTRIAL (HI)':[57,163,49],'INSTITUTIONAL (IN)':[0,169,230],'LIGHT INDUSTRIAL (LI)':[169,0,230],'PLANNED UNIT DEVELOPMENT (PUD)':[115,255,223],'RURAL RESIDENTIAL (RR)':[215,158,158],'SUBURBAN RESIDENTIAL (SR)':[255,255,0],'RURAL PRESERVATION (RP)':[180,215,158]},
  gallatin:{'A':[85,255,0],'CC':[168,0,0],'CG':[230,0,0],'CG(PUD)':[255,0,197],'CS':[255,190,232],'CS(PUD)':[245,122,122],'CSL':[223,115,255],'COMMERCIAL SERVICES (CS)':[255,211,127],'GO':[0,37,167],'IG':[197,0,255],'IR':[132,0,168],'INSTITUTIONAL (IN)':[0,169,230],'LIGHT INDUSTRIAL (LI)':[169,0,230],'MPO':[255,0,0],'MRO':[230,76,0],'MRO(PUD)':[255,122,0],'MU':[74,164,254],'MUG':[0,255,197],'OR':[255,190,190],'PBP':[255,0,77],'PGC':[214,157,188],'PNC':[230,0,169],'PLANNED UNIT DEVELOPMENT (PUD)':[115,255,223],'R-06':[205,137,102],'R-06(PRD)':[255,255,190],'R-06(PUD)':[215,176,158],'R-08':[245,162,122],'R-08(PRD)':[255,167,127],'R-10':[230,127,0],'R-10(PRD)':[245,202,122],'R-15':[255,235,190],'R-15(PRD)':[255,170,0],'R-15(PUD)':[255,211,127],'R-20':[255,255,0],'R-20(PRD)':[255,235,175],'R-20(PUD)':[230,230,0],'R-40':[255,255,115],'RURAL RESIDENTIAL (RR)':[180,215,158],'SP':[159,164,238]},
  hendersonville:{'DN':[115,0,0],'ER':[193,227,91],'GC':[230,0,0],'HC':[168,0,230],'I':[0,115,255],'MFR':[138,101,59],'MXC':[255,85,0],'MXR':[204,168,102],'NC':[255,166,128],'O':[255,117,225],'RR':[165,255,117],'SR-1':[230,230,0],'SR-2':[255,170,0],'SR-3':[20,158,206]},
  whitehouse:{'R-10':[233,255,190],'R-15':[255,235,175],'R-20':[137,205,102],'C-1':[255,190,190],'R-TC':[255,255,0],'C-1R':[215,176,158],'C-2':[255,0,0],'C-4':[168,56,0],'C-5':[115,0,0],'C-6':[255,170,0],'I-1':[190,232,255],'I-2':[115,178,255],'NCRPUD':[205,170,102],'SRPUD':[68,137,112]},
  portland:{'GCS':[255,190,190],'R15':[168,168,0],'RS20':[85,255,0],'R7.5':[255,255,190],'R10':[255,255,115],'RS40':[112,168,0],'RM1':[211,255,190],'CBD':[255,0,0],'R40':[255,255,255],'IR':[115,178,255],'OPS':[255,115,223],'IG':[190,210,255],'HCD':[223,115,255],'RS15':[163,255,115],'RMHP':[233,255,190],'NSD':[255,190,232],'PUD':[190,255,232],'RM1 - PUD':[133,133,133]},
  millersville:{'NC':[253,127,111],'O':[255,0,197],'MXC':[255,170,0],'GC':[230,0,0],'HC':[169,0,230],'RR':[40,184,67],'ER':[162,255,31],'SR-1':[255,255,0],'SR-2':[230,152,0],'THR':[66,99,0],'MXR':[183,129,74],'MFR':[115,38,0],'I':[0,112,255],'PD':[153,153,153],'H':[26,26,26]},
  goodlettsville:{'A':[252,225,207],'CC':[215,252,199],'CG':[179,193,252],'CPUD':[215,252,252],'CPUDL':[252,207,192],'CS':[194,179,252],'CSL':[182,252,232],'GOPUD':[179,252,187],'HDRPUD':[252,227,184],'IC':[252,182,218],'IG':[210,252,220],'IR':[182,243,252],'LDRPUD':[252,188,182],'MDRPUD':[192,252,204],'NMOVRLY':[251,252,179],'OP':[219,252,182],'R10':[252,179,196],'R15':[184,210,252],'R25':[252,251,212],'R40':[236,202,252],'R7':[222,212,252],'RC1PUD':[215,234,252],'RLSPUD':[252,215,230],'ROPUD':[251,182,252]},
  westmoreland:{'R1':[0,38,115],'R2':[252,146,31],'C1':[237,81,81],'C2':[20,158,206],'I1':[56,168,0]}
};
var UCDD_FALLBACK=['#e6194B','#3cb44b','#ffe119','#4363d8','#f58231','#911eb4','#42d4f4','#f032e6','#bfef45','#469990'];
function ucddRgbaArr(a:any,al:number){ return 'rgba('+a[0]+','+a[1]+','+a[2]+','+al+')'; }
function ucddColor(key:string,zone:any){ var z=(zone==null?'':(''+zone)).toUpperCase(); var cm=UCDD_COLORS[key]||{}; if(cm[z]) return ucddRgbaArr(cm[z],0.65); var h=0; for(var i=0;i<z.length;i++){ h=(h*31+z.charCodeAt(i))>>>0; } return UCDD_FALLBACK[h%UCDD_FALLBACK.length]; }
function ucddHit(bb:any,vw:any){ return !(bb[0]>vw[2]||bb[2]<vw[0]||bb[1]>vw[3]||bb[3]<vw[1]); }
// Authoritative UCDD district names (from each community's adopted zoning ordinance, resolved 2026-06-28).
// Empty/blank = no published ordinance found (South Carthage, Livingston, Spencer) or code not in the
// adopted ordinance (Smith SCA/GRD/CRT) -> legend shows the code only, no invented name.
var UCDD_NAMES:any={
  smith:{'A-1':'Agriculture District','C-1':'General Commercial District','R-1':'Residential District','R-C':'Residential-Commercial District','I-1':'General Industrial District'},
  smithville:{'R-1':'Residential - Low Density','R-2':'Residential - High Density','C-1':'Local Commercial District','C-2':'Central Business District','I-1':'Light Industrial District'},
  south_carthage:{},
  gordonsville:{'A-1':'Agricultural District','C-1':'Limited Commercial District','C-2':'Highway Commercial District','I-1':'Light Industrial District','I-2':'Heavy Industrial District','M-1':'Mixed Residential-Commercial District','R-1':'Low Density Residential District','R-2':'High Density Residential District'},
  algood:{'R-D':'Single-Family & Duplex Residential District','R-1':'Low Density Residential District','R-2':'Medium Density Residential District','R-3':'High Density Residential District','C-A':'General Commercial District','C-B':'Central Business District','C-C':'Planned Commercial District','I-1':'Light Industrial District','I-2':'Heavy Industrial District'},
  baxter:{'R-1':'Low Density Residential District','R-2':'High Density Residential District','R-M':'Residential Medium Density District','CBD':'Central Business District','C-1':'Central Commercial District','C-2':'General Commercial District','I-1':'Light Industrial District'},
  livingston:{},
  morrison:{'A-1':'Agriculture District','R-1':'Low Density Residential District','R-2':'High Density Residential District','C-1':'Central Commercial District','C-2':'General Commercial District','I-1':'Light Industrial District'},
  cannon:{'A-1':'Agricultural District','R-1':'Single-Family Low Density Residential District','C-1':'Commercial District','I-1':'General Industrial District'},
  monterey:{'C-1':'Limited Commercial District','C-2':'General Commercial District','I-1':'Light Industrial District','R-1':'Low Density Residential District','R-2':'High Density Residential District','R-R':'Rural Residential District'},
  spencer:{},
  sumner_co:{},
  gallatin:{'A':'Agricultural District','CC':'Core Commercial District','CG':'Commercial General District','CG(PUD)':'Commercial General District (PUD)','CS':'Commercial Services District','CS(PUD)':'Commercial Services District (PUD)','CSL':'Commercial Services Limited District','GO':'General Office District','IG':'General Industrial District','IR':'Restrictive Industrial District','MPO':'Medical-Professional Office District','MRO':'Multiple Residential and Office District','MRO(PUD)':'Multiple Residential and Office District (PUD)','MU':'Mixed Use District','MUG':'Mixed Use General District','OR':'Office Residential District','PBP':'Planned Business Park District','PGC':'Planned General Commercial District','PNC':'Planned Neighborhood Commercial District','R-06':'High Density Residential District','R-06(PRD)':'High Density Residential District (PRD)','R-06(PUD)':'High Density Residential District (PUD)','R-08':'Medium Density Residential District','R-08(PRD)':'Medium Density Residential District (PRD)','R-10':'Medium Density Residential District','R-10(PRD)':'Medium Density Residential District (PRD)','R-15':'Low-Medium Density Residential District','R-15(PRD)':'Low-Medium Density Residential District (PRD)','R-15(PUD)':'Low-Medium Density Residential District (PUD)','R-20':'Low Density Residential District','R-20(PRD)':'Low Density Residential District (PRD)','R-20(PUD)':'Low Density Residential District (PUD)','R-40':'Low Density Residential District','SP':'Specific Plan District'},
  hendersonville:{'DN':'Dockside Neighborhood','ER':'Estate Residential','GC':'General Commercial','HC':'Heavy Commercial','I':'Industrial','MFR':'Multi-Family Residential','MXC':'Mixed Commercial','MXR':'Mixed Residential','NC':'Neighborhood Commercial','O':'Office','RR':'Rural Residential','SR-1':'Suburban Residential 1','SR-2':'Suburban Residential 2','SR-3':'Suburban Residential 3'},
  whitehouse:{'R-10':'High Density Residential','R-15':'Medium Density Residential','R-20':'Low Density Residential','C-1':'Central Business Commercial','R-TC':'High Density Town Center Commercial','C-1R':'Central Business - Gateway Infill Residential','C-2':'Interstate Sign District','C-4':'Office / Professional','C-5':'Limited Office / Professional','C-6':'Town Center Commercial','I-1':'Light Industrial','I-2':'Heavy Industrial','NCRPUD':'Neighborhood Center','SRPUD':'Suburban Residential'},
  portland:{'CBD':'Central Business District','GCS':'General Commercial Service District','HCD':'Heavy Commercial Distribution','IG':'General Industrial District','IR':'Restrictive Industrial District','NSD':'Neighborhood Service District','OPS':'Office / Professional Service District','PUD':'Planned Unit Development','R10':'Low Density Residential','R15':'Low Density Residential','R40':'Low Density Residential','R7.5':'Medium Density Residential','RM1':'High Density Residential','RM1 - PUD':'High Density Residential (PUD)','RMHP':'Mobile Home Park Residential','RS15':'Single-Family Low Density Residential','RS20':'Single-Family Low Density Residential','RS40':'Single-Family Low Density Residential'},
  millersville:{'NC':'Neighborhood Commercial','O':'Office','MXC':'Mixed Commercial','GC':'General Commercial','HC':'Heavy Commercial','RR':'Rural Residential','ER':'Estate Residential','SR-1':'Suburban Residential 1','SR-2':'Suburban Residential 2','THR':'Townhome Residential','MXR':'Mixed Residential','MFR':'Multi-Family Residential','I':'Industrial','PD':'Planned Development','H':'Historic and Landmarks'},
  goodlettsville:{'A':'Agricultural','CC':'Commercial Core','CG':'Commercial General','CPUD':'Commercial Planned Unit Development','CPUDL':'Commercial Planned Unit Development Limited','CS':'Commercial Services','CSL':'Commercial Services Limited','GOPUD':'General Office Planned Unit Development','HDRPUD':'High Density Residential Planned Unit Development','IC':'Industrial Commercial','IG':'Industrial General','IR':'Industrial Restricted','LDRPUD':'Low Density Residential Planned Unit Development','MDRPUD':'Medium Density Residential Planned Unit Development','NMOVRLY':'North Main Street Overlay District','OP':'Office Professional','R10':'Medium Density Residential','R15':'Medium Density Residential','R25':'Low Density Residential','R40':'Low Density Residential','R7':'High Density Residential','RC1PUD':'Regional Center Planned Unit Development High Intensity','RLSPUD':'Residential Limited Scale Planned Unit Development','ROPUD':'Restricted Office Planned Unit Development'},
  westmoreland:{'R1':'Low Density Residential','R2':'Medium Density Residential','C1':'Central Business','C2':'Highway Service','I1':'General Industrial'}
};
function ucddName(key:string,code:any){ var z=(code==null?'':(''+code)).toUpperCase(); var m=UCDD_NAMES[key]||{}; return m[z]||''; }
// point-in-polygon: is [lng,lat] inside a GeoJSON feature (Polygon/MultiPolygon, honoring holes)? Reuses pointInRing.
function featHitLL(feat:any,lng:number,lat:number){ if(!feat||!feat.geometry) return false; var g=feat.geometry; var p=[lng,lat];
  function inPoly(poly:any){ if(!poly||!poly.length||!pointInRing(p,poly[0])) return false; for(var h=1;h<poly.length;h++){ if(pointInRing(p,poly[h])) return false; } return true; }
  if(g.type==='Polygon') return inPoly(g.coordinates);
  if(g.type==='MultiPolygon'){ for(var i=0;i<g.coordinates.length;i++){ if(inPoly(g.coordinates[i])) return true; } return false; }
  return false; }

function normalize(attrs:any, src:any){ const n:any={src:src}; n.pin=pick(attrs,src.f.pin); n.owner=pick(attrs,src.f.owner); n.owner2=pick(attrs,src.f.owner2); n.address=pick(attrs,src.f.address); n.mail=pick(attrs,src.f.mail); n.subdiv=pick(attrs,src.f.subdiv); n.lot=pick(attrs,src.f.lot); n.acres=pick(attrs,src.f.acres); n.zoning=pick(attrs,src.f.zoning); n.parcelno=pick(attrs,src.f.parcelno); n.assr=pick(attrs,src.f.assr); n.tpad=pick(attrs,src.f.tpad); const gm=(n.tpad.match(/gislink=([^&]+)/)||[])[1]; n.gislink=gm?decodeURIComponent(gm):pick(attrs,src.f.gislinkf); n.deedBook=pick(attrs,src.f.deedBook); n.deedPage=pick(attrs,src.f.deedPage); n.legalref=pick(attrs,src.f.legalref); n.deedref=pick(attrs,src.f.deedref); n.state=src.state; n.county=src.county||pick(attrs,[src.countyField]); if(n.county) n.county=n.county.toUpperCase().replace(/ COUNTY$/,'').trim(); return n; }
function parseBookPage(n:any){ if(n.deedBook&&n.deedPage&&/\d/.test(n.deedBook)&&/\d/.test(n.deedPage)) return {book:n.deedBook.replace(/[^0-9A-Za-z]/g,''),page:n.deedPage.replace(/[^0-9A-Za-z]/g,'')}; const ref=n.legalref||''; const m=ref.match(/^\s*([0-9A-Za-z]+)\s*[-\/]\s*([0-9A-Za-z]+)\s*$/); if(m) return {book:m[1],page:m[2]}; return null; }
function tsNameUrl(owner:string){ const name=(owner||'').split(',')[0].trim(); return TS_BASE+'nameSearch.php?'+qs({nameType:'2',searchType:'PA',indexType:'BOTH',p1:name,p2:'',expandAll:'on',startDate:'',endDate:'',itype:'0',executeSearch:'Execute Search'}); }
function tsBookPageUrl(bp:any){ return TS_BASE+'bookPageSearch.php?'+qs({book:bp.book,page:bp.page,fileNumber:'',executeSearch:'Execute Search'}); }
function outFieldsFor(s:any){ const set:any={}; ['pin','owner','owner2','address','mail','subdiv','lot','acres','zoning','assr','tpad','gislinkf','deedBook','deedPage','legalref','deedref','parcelno'].forEach((k)=>{ (s.f[k]||[]).forEach((fn:string)=>{ set[fn]=1; }); }); if(s.countyField) set[s.countyField]=1; return Object.keys(set).join(',')||'*'; }
function bboxIntersect(a:any,b:any){ return !(b[0]>a[2]||b[2]<a[0]||b[1]>a[3]||b[3]<a[1]); }

// ============================ Coverage / Projects layer (WIP survey jobs) ============================
// Ported from the standalone Coverage Map (2026-06-15): status-colored project pins + Status/Deadline/
// County/JobType/Year filters + search + one-click job-folder links, reading the WIP list live. Toggled
// from the toolbar (Projects Off/On). Status colors = the WIP "Project Status" column's exact colors.
const PALETTE:any = { 'Fielding':'Blue','Crew Assigned':'Magenta','Fielding Complete':'Cyan','Drafting':'Purple','Drafting Complete':'DarkOrange','Waiting on Client':'#008080','Planning Approval':'LimeGreen','HOLD':'Black','Pending Bill':'Brown','Billed':'Red','Paid - Closeout':'Green','Dropped Project':'Gray','Initial Research':'Pink','Onsite Meeting':'Pink','Waiting on Signatures':'Pink','Plat Submitted':'Pink' };
const STATUS_ORDER:any = ['Initial Research','Fielding','Crew Assigned','Fielding Complete','Drafting','Drafting Complete','Onsite Meeting','Waiting on Client','Waiting on Signatures','Plat Submitted','Planning Approval','HOLD','Pending Bill','Billed','Paid - Closeout','Dropped Project'];
const DEFAULT_STATUS_ON:any = ['Fielding','Crew Assigned','Drafting','Drafting Complete','Planning Approval','Waiting on Client'];
const DEADLINE_ORDER:any = ['Overdue','Due in 14 days','Due in 30 days','Due later','No date','Completed'];
// ---- Inquiries layer (IQ list) â amber triangles on each inquiry's parcel; mirrors the Projects layer ----
const IQ_INQUIRIES_GUID = 'a2da06ea-55d3-4221-9988-035800aa59a5';
const IQ_STATUS_ORDER:any = ['Initial Research','Created Quote','Draft Quote','Sent Quote','Attempted Contact','Declined Quote','(blank)'];
const IQ_DEFAULT_STATUS_ON:any = ['Initial Research','Draft Quote','Created Quote'];
const IQ_COLOR = '#f59e0b';
const IQ_EXCLUDE_STATUS = 'Accepted Quote';
const IQ_QUOTES_REL = '/References/Quotes';
const IQ_TRI_SVG = '<svg width="20" height="18" viewBox="0 0 20 18" xmlns="http://www.w3.org/2000/svg"><path d="M10 1.6 L18.7 16.4 L1.3 16.4 Z" fill="'+IQ_COLOR+'" stroke="#ffffff" stroke-width="1.7" stroke-linejoin="round"/></svg>';
function colorFor(s:any){ return PALETTE[s]||'Pink'; }
function deadlineBucket(status:any, fnlt:any){ if(status==='Paid - Closeout'||status==='Dropped Project') return 'Completed'; if(!fnlt) return 'No date'; const d=new Date(fnlt); if(isNaN(d.getTime())) return 'No date'; const now=new Date(); now.setHours(0,0,0,0); const days=(d.getTime()-now.getTime())/86400000; if(days<0) return 'Overdue'; if(days<=14) return 'Due in 14 days'; if(days<=30) return 'Due in 30 days'; return 'Due later'; }

export interface IPropertyDeedMapWebPartProps { title: string; workerUrl: string; zoneListTitle: string; zoningAssetBase: string; projectsListGuid: string; inquiriesListGuid: string; }

export default class PropertyDeedMapWebPart extends BaseClientSideWebPart<IPropertyDeedMapWebPartProps> {
  private map:any; private parcelLayer:any; private hiLayer:any; private labels:any; private bases:any;
  private POP:any = {}; private pseq=0;
  private inflight:any[] = []; private loadTimer:any = null; private rzTimer:any = null;
  private loadedBounds:any = null; private loadedZoom:number = -1;
  private zoneByPin:any = {}; private zoningView=true; private zoningEdit=false;
  private zTarget:any = null; private loadSeq=0; private tagJur:string='auto';
  private workedByPin:any={}; private workView=false; private workEdit=false; private wTarget:any=null; private _collW=false; private _workLoaded=false; private workedWipIds:any={}; private _workCount=0; private _workUnresolved=0; private _wipPick:any[]=[]; private _workColorMode:string='flat'; private pWorkYearOn:any={};
  private selFeat:any=null; private selN:any=null; private selLayer:any=null; private labelLayer:any=null; private workedGeomLayer:any=null; private _workGeomLoaded=false; private _folderCache:any={}; private _printMap:any=null;
  private splitState:any=null; private splitLayer:any=null; private splitTmp:any[]=[]; private splitMarkers:any[]=[]; private _splitClick:any=null; private _splitDrawPopup:any=null;
  private femaLayer:any=null; private _femaOn=false; private areasLayer:any=null; private _areasRenderer:any=null; private areas:any[]=[]; private _areasOn=false;
  private ucddLayer:any=null; private _ucddSeq=0; private _ucddCount=0; private _ucddBounds:any=null; private _ucddZoom:number=-1; private _ucddCache:any={}; private _ucddRenderer:any=null;
  private areaState:any=null; private areaMarkers:any[]=[]; private areaLine:any=null; private _areaClick:any=null;
  private projects:any[]=[]; private projectLayer:any=null; private _projRenderer:any=null; private _projOn=false; private _projLoaded=false;
  private inquiries:any[]=[]; private inqLayer:any=null; private _inqOn=false; private _inqLoaded=false; private inqGeo:any={}; private _quoteFolderCache:any={};
  private iqStatusOn:any={}; private iqCountyOn:any={}; private iqYearOn:any={}; private iqSearch=''; private _collI=false; private _iqDimColl:any={status:false,county:true,year:true}; private _iqLocated=0; private _iqUnplaced=0;
  private pStatusOn:any={}; private pCountyOn:any={}; private pTypeOn:any={}; private pYearOn:any={}; private pDeadlineOn:any={}; private pSearch='';
  private jurShow:any={RBS:true,Lafayette:true,Macon:true,smith:false,smithville:false,south_carthage:false,gordonsville:false,algood:false,baxter:false,livingston:false,morrison:false,cannon:false,monterey:false,spencer:false}; private _collZ=false; private _collUcdd:boolean=false; private _collP=false; private _collLegend=false; private _dimColl:any={status:false,deadline:false,county:true,type:false,year:true};

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
  private get projectsListGuid(): string { return this.properties.projectsListGuid || 'ecfa34b1-214a-4b6a-a661-0d074800714e'; }
  private get inquiriesListGuid(): string { return this.properties.inquiriesListGuid || IQ_INQUIRIES_GUID; }

  public render(): void {
    this.domElement.innerHTML = `
      <style>
        .dls-pm{font-family:'Segoe UI',Arial,sans-serif;color:#0f172a;width:100%;box-sizing:border-box;}
        .dls-pm .dls-inq{background:none;border:none;}
        .dls-pm .dls-inq svg{display:block;filter:drop-shadow(0 1px 1.5px rgba(0,0,0,.5));}
        .dls-pm .dls-inqpop table{border-collapse:collapse;margin:2px 0 4px;}
        .dls-pm .dls-inqpop td{font-size:12px;padding:1px 0;vertical-align:top;}
        .dls-pm .dls-inqpop td.k{color:#64748b;padding-right:8px;white-space:nowrap;font-size:11px;}
        .dls-pm .dls-inq-links{display:flex;flex-direction:column;gap:2px;margin-top:2px;}
        @media (min-width:1300px) and (orientation:landscape){ .dls-pm{width:98vw;position:relative;left:50%;margin-left:-49vw;} }
        .dls-pm .bar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;background:#1f2a37;color:#fff;padding:7px 10px;border-radius:8px 8px 0 0;border-bottom:3px solid #f59e0b;}
        .dls-pm .adv-tg{background:#33445a;color:#fff;border:none;border-radius:6px;padding:5px 9px;font-size:11.5px;cursor:pointer;font-weight:600;}
        .dls-pm .adv-tg.on{background:#f59e0b;color:#1a1205;}
        .dls-pm #advpanel{display:none;background:#eef2f7;border-bottom:1px solid #cbd5e1;padding:8px 10px;}
        .dls-pm #advpanel.open{display:block;}
        .dls-pm #advpanel .advwrap{display:flex;flex-wrap:wrap;align-items:flex-end;gap:8px 10px;}
        .dls-pm #advpanel .advf{display:flex;flex-direction:column;gap:2px;}
        .dls-pm #advpanel label{font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.03em;color:#475569;}
        .dls-pm #advpanel input,.dls-pm #advpanel select{font-size:12px;padding:4px 6px;border:1px solid #94a3b8;border-radius:5px;background:#fff;color:#0f172a;box-sizing:border-box;}
        .dls-pm #advpanel input{width:104px;}
        .dls-pm #advpanel select#adv-county{width:154px;}
        .dls-pm #advpanel .adv-go{background:#16a34a;color:#fff;border:none;border-radius:5px;padding:6px 12px;font-size:12px;font-weight:600;cursor:pointer;}
        .dls-pm #advpanel .adv-clear{background:#f1f5f9;border:1px solid #cbd5e1;border-radius:5px;padding:6px 10px;font-size:12px;cursor:pointer;color:#0f172a;}
        .dls-pm #advpanel .adv-hint{flex-basis:100%;font-size:10px;color:#64748b;margin-top:2px;}
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
        .dls-pm #wmode{background:#2a4a33;color:#fff;}
        .wk-row{margin:3px 0}
        .wk-row select,.wk-row input{width:100%;box-sizing:border-box;padding:4px;border:1px solid #ccc;border-radius:4px;font:12px sans-serif}
        .wk-or{font-size:11px;color:#888;margin:6px 0 2px;text-align:center}
        .wk-cur{font-size:12px;margin:2px 0}
        .dls-wk{margin:2px 0}
        .dls-plabel{color:#444;font:600 10px/1 system-ui,sans-serif;text-shadow:0 0 2px #fff,0 0 2px #fff,0 0 2px #fff;white-space:nowrap;transform:translate(-50%,-50%);}
        .dls-plabel.print{color:#333;font-size:9px;}
        .dls-print-modal{position:fixed;inset:0;background:rgba(20,24,33,.55);z-index:99998;display:flex;align-items:center;justify-content:center;}
        .dls-print-modal .dlsheet{width:8.5in;height:11in;background:#fff;box-shadow:0 8px 40px rgba(0,0,0,.4);display:flex;flex-direction:column;padding:.3in;box-sizing:border-box;}
        .dls-print-modal .dlhd{text-align:center;font:700 16px Georgia,serif;margin-bottom:6px;}
        .dls-print-modal .dlmap{flex:1 1 auto;border:1px solid #999;}
        .dls-print-modal .dlft{display:flex;justify-content:space-between;align-items:flex-end;margin-top:8px;font:11px/1.35 Arial,sans-serif;}
        .dls-print-modal .dlinfo div{margin:1px 0;}
        .dls-print-modal .dlscale{text-align:right;}
        .dls-print-modal .sbar{display:inline-flex;align-items:center;gap:6px;}
        .dls-print-modal .sbar .sb{height:6px;background:#333;border:1px solid #333;}
        .dls-print-modal .srat{font-size:10px;color:#444;margin-top:2px;}
        .dls-print-modal .dpx{position:absolute;top:14px;right:24px;display:flex;gap:8px;}
        .dls-print-modal .dpx button{padding:8px 14px;border:0;border-radius:6px;font:600 13px sans-serif;cursor:pointer;background:#1565ff;color:#fff;}
        .dls-print-modal .dpx button#dlsPrintClose{background:#555;}
        @media print {
          @page { size: letter portrait; margin: 0; }
          html, body { margin:0 !important; padding:0 !important; background:#fff !important; }
          body > *:not(.dls-print-modal){ display:none !important; }
          .dls-print-modal{ position:static !important; inset:auto !important; display:block !important; background:#fff !important; width:auto !important; height:auto !important; }
          .dls-print-modal .dlsheet{ box-shadow:none !important; width:8.5in !important; height:10.7in !important; margin:0 auto !important; padding:0.3in !important; box-sizing:border-box !important; }
          .dls-print-modal .dpx{ display:none !important; }
          .dls-print-modal .leaflet-control-container{ display:none !important; }
        }
        .dls-pm #zlegend{position:absolute;z-index:900;top:8px;right:8px;background:rgba(255,255,255,.98);border:1px solid #cbd5e1;border-radius:8px;font-size:11px;width:242px;max-height:calc(100% - 22px);overflow:auto;box-shadow:0 6px 20px rgba(0,0,0,.16);display:none;}
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
        .zp .zp-jurs{display:flex;flex-wrap:wrap;align-items:center;gap:3px;margin:0 0 7px;font-size:10.5px;color:#475569;}
        .zp .zp-jurbtn{border:1px solid #cbd5e1;border-radius:5px;padding:2px 7px;font-size:10.5px;font-weight:600;background:#f1f5f9;color:#334155;cursor:pointer;}
        .zp .zp-jurbtn.on{background:#1f2a37;color:#fff;border-color:#1f2a37;}
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
        .dls-pm #proj{background:#33445a;color:#fff;}
        .dls-pm #plegend{display:none;}
        .dls-pm .pp-ct{color:#9fb0c3;font-size:10px;font-weight:400;}
        .dls-pm .pp-search{width:100%;box-sizing:border-box;font-size:11.5px;padding:5px 7px;border:1px solid #cbd5e1;border-radius:5px;margin-bottom:6px;}
        .dls-pm .pp-reset{width:100%;font-size:11px;padding:5px;margin-bottom:8px;background:#33445a;color:#fff;border:none;border-radius:5px;cursor:pointer;font-weight:600;}
        .dls-pm .pp-all{color:#2563eb;cursor:pointer;font-weight:600;text-transform:none;letter-spacing:0;font-size:10px;margin-left:auto;}
        .dls-pm .pp-row{display:flex;align-items:center;gap:6px;padding:2px 3px;border-radius:4px;cursor:pointer;font-size:11px;}
        .dls-pm .pp-row:hover{background:#f1f5f9;} .dls-pm .pp-row.off{opacity:.4;}
        .dls-pm .pp-dot{width:11px;height:11px;border-radius:50%;border:1px solid rgba(0,0,0,.3);flex:none;}
        .dls-pm .pp-sq{width:11px;height:11px;border-radius:3px;background:#94a3b8;flex:none;}
        .dls-pm .pp-nm{flex:1;color:#0f172a;} .dls-pm .pp-n{color:#64748b;font-size:10px;}
        .dls-pm .lp-sec{border-top:1px solid #e2e8f0;}
        .dls-pm .lp-sec:first-child{border-top:none;}
        .dls-pm .lp-hd{display:flex;align-items:center;gap:6px;padding:7px 10px;font-weight:700;font-size:12px;cursor:pointer;background:#f1f5f9;color:#1f2a37;-webkit-user-select:none;user-select:none;}
        .dls-pm .lp-sec:first-child .lp-hd{border-radius:8px 8px 0 0;}
        .dls-pm .lp-hd .tw{font-size:9px;color:#64748b;display:inline-block;}
        .dls-pm .lp-sec.coll .lp-hd .tw{transform:rotate(-90deg);}
        .dls-pm .lp-bd{padding:7px 10px;}
        .dls-pm .lp-sec.coll .lp-bd{display:none;}
        .dls-pm .lp-sub{margin:6px 0;}
        .dls-pm .lp-subhd{display:flex;align-items:center;gap:5px;font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.03em;color:#475569;cursor:pointer;margin-bottom:3px;-webkit-user-select:none;user-select:none;}
        .dls-pm .lp-subhd .tw{font-size:8px;color:#94a3b8;display:inline-block;}
        .dls-pm .lp-sub.coll .lp-subhd .tw{transform:rotate(-90deg);}
        .dls-pm .lp-sub.coll .lp-subbd{display:none;}
        .dls-pm .zsh{display:flex;flex-wrap:wrap;align-items:center;gap:4px 8px;margin-bottom:7px;font-size:11px;color:#334155;}
        .dls-pm .zsh .zshl{display:flex;align-items:center;gap:3px;cursor:pointer;}
        .dls-pop b{font-size:13px;} .dls-pop .m{color:#6b7280;margin:2px 0;font-size:11px;}
        .dls-pop a.dls-pop-a{display:inline-block;margin-top:6px;padding:5px 9px;background:#C2410C;color:#fff;border-radius:6px;text-decoration:none;font-weight:600;font-size:11px;}
      </style>
      <div class="dls-pm">
        <div class="bar">
          <strong>${esc(this.properties.title) || 'DLS Property &amp; Deed Map'}</strong>
          <select id="area" title="Dataset to search / jump to"></select>
          <select id="mode"><option value="owner">Owner</option><option value="address">Address</option><option value="parcel">Parcel ID</option></select>
          <input id="q" placeholder="Search owner name&hellip;" />
          <button id="go">Search</button>
          <button id="clear" class="ghost">Clear</button>
          <button id="adv-toggle" class="adv-tg" title="Advanced search: county, map, parcel, subdivision">Advanced &#9662;</button>
          <span class="sp"></span>
          <select id="base"><option value="aerial">Aerial</option><option value="streets" selected>Streets</option><option value="topo">Topo</option></select>
          <select id="zmode" title="Zoning layer (View / Edit)"><option value="off" selected>Zoning: Off</option><option value="view">Zoning: View</option><option value="edit">Zoning: Edit (tag lots)</option></select>
          <select id="proj" title="Survey projects layer (WIP)"><option value="off">Projects: Off</option><option value="on" selected>Projects: On</option></select>
          <select id="inq" title="Inquiries layer (open quotes on their parcels)"><option value="off">Inquiries: Off</option><option value="on" selected>Inquiries: On</option></select>
          <select id="wmode" title="Work history layer (surveyed parcels)"><option value="off" selected>Work history: Off</option><option value="view">Work history: View</option><option value="edit">Work history: Edit (mark surveyed)</option></select>
          <button id="fs" class="ghost" title="Full screen (Esc to exit)">Full screen</button>
          <span id="status">Loading&hellip;</span>
        </div>
        <div id="advpanel"><div class="advwrap">
          <div class="advf"><label>County</label><select id="adv-county"></select></div>
          <div class="advf"><label>Control Map</label><input id="adv-map" placeholder="e.g. 71" /></div>
          <div class="advf"><label>Group</label><input id="adv-group" placeholder="opt." /></div>
          <div class="advf"><label>Parcel</label><input id="adv-parcel" placeholder="e.g. 12 or 12.00" /></div>
          <div class="advf"><label>Subdivision</label><input id="adv-subdiv" placeholder="name contains&hellip;" /></div>
          <div class="advf"><label>Sort by</label><select id="adv-sort"><option value="parcel">Parcel</option><option value="owner">Owner</option><option value="address">Address</option><option value="acres">Acreage</option></select></div>
          <button id="adv-go" class="adv-go">Search</button>
          <button id="adv-clear" class="adv-clear">Clear</button>
          <div class="adv-hint">Tennessee statewide (86 counties). Enter a Control Map + Parcel to jump to the exact lot, or a Subdivision name to list matches. (Owner &amp; Address use the bar above.)</div>
        </div></div>
        <div class="stage">
          <div id="map"></div>
          <div id="results"><h4><span id="rtitle">Results</span><span class="x" id="rclose">&times;</span></h4><div id="rlist"></div></div>
          <div id="legend"><b>Parcels load at zoom ${MINZOOM}+</b> &mdash; pan/zoom to your area.<br/><span class="src" id="legsrc">Active data: &mdash;</span><div class="disc">Reference only &mdash; not a boundary survey, title opinion, or zoning determination. Parcel &amp; owner data are pulled live from each assessor and may lag.</div></div>
          <div id="zlegend"></div>
          <div id="plegend"></div>
        </div>
      </div>`;

    const $ = (s:string)=>this.domElement.querySelector(s) as any;
    const areaSel = $('#area');
    SOURCES.forEach((s)=>{ const o=document.createElement('option'); o.value=s.id; o.textContent=s.label; areaSel.appendChild(o); });
    $('#mode').onchange = (e:any)=>{ const p:any={owner:'Search owner nameâ¦',address:'Search street addressâ¦',parcel:'Search parcel IDâ¦'}; $('#q').placeholder=p[e.target.value]; };
    $('#go').onclick = ()=>this.runSearch();
    $('#q').addEventListener('keydown',(e:any)=>{ if(e.key==='Enter') this.runSearch(); });
    const advCounty = $('#adv-county');
    if(advCounty){ const o0=document.createElement('option'); o0.value=''; o0.textContent='â County â'; advCounty.appendChild(o0); TN_COUNTIES.forEach((c:string)=>{ const o=document.createElement('option'); o.value=c; o.textContent=c; advCounty.appendChild(o); }); }
    const advToggle = $('#adv-toggle');
    if(advToggle) advToggle.onclick = ()=>{ const p=$('#advpanel'); const open=p.classList.toggle('open'); advToggle.classList.toggle('on', open); };
    const advGo = $('#adv-go'); if(advGo) advGo.onclick = ()=>this.runAdvancedSearch();
    const advClear = $('#adv-clear'); if(advClear) advClear.onclick = ()=>{ ['#adv-county','#adv-map','#adv-group','#adv-parcel','#adv-subdiv'].forEach((sl:string)=>{ const el=$(sl); if(el) el.value=''; }); this.setStatus('Advanced search cleared'); };
    ['#adv-map','#adv-group','#adv-parcel','#adv-subdiv'].forEach((sl:string)=>{ const el=$(sl); if(el) el.addEventListener('keydown',(ev:any)=>{ if(ev.key==='Enter') this.runAdvancedSearch(); }); });
    $('#clear').onclick = ()=>{ this.hiLayer.clearLayers(); $('#results').style.display='none'; this.clearSelection(); };
    $('#rclose').onclick = ()=>{ $('#results').style.display='none'; };
    $('#base').onchange = (e:any)=>this.setBase(e.target.value);
    $('#zmode').onchange = (e:any)=>this.setZoningMode(e.target.value);
    $('#proj').onchange = (e:any)=>this.setProjectsMode(e.target.value==='on');
    $('#inq').onchange = (e:any)=>this.setInquiriesMode(e.target.value==='on');
    $('#wmode').onchange = (e:any)=>this.setWorkMode(e.target.value);
    $('#fs').onclick = ()=>this.toggleFs();
    this.buildMap();
    this.setZoningMode('off');   // default Zoning OFF (Projects on, Work history off)
    this.setProjectsMode(true);   // Projects layer ON by default (one Master Map)
    this.setInquiriesMode(true);  // Inquiries layer ON by default (early-stage quotes only)
  }

  private buildMap(): void {
    const mapEl = this.domElement.querySelector('#map');
    this.map = L.map(mapEl,{minZoom:6,maxZoom:20}).setView([36.521,-86.029],16);
    this.bases = {
      aerial: L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',{maxZoom:20,attribution:'Imagery Â© Esri'}),
      streets: L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}',{maxZoom:20,attribution:'Â© Esri'}),
      topo: L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer/tile/{z}/{y}/{x}',{maxZoom:20,attribution:'Â© Esri'})
    };
    this.bases.streets.addTo(this.map);
    this.labels = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}',{maxZoom:20,opacity:.9});
    this.map.createPane('zoning'); this.map.getPane('zoning').style.zIndex='350'; this.map.getPane('zoning').style.pointerEvents='none';
    this.map.createPane('zsplit'); this.map.getPane('zsplit').style.zIndex='420'; this.map.getPane('zsplit').style.pointerEvents='none';
    this.splitLayer = L.layerGroup().addTo(this.map);
    this.map.createPane('areas'); this.map.getPane('areas').style.zIndex='430'; this.map.getPane('areas').style.pointerEvents='none';
    this._areasRenderer = L.svg({pane:'areas'}); this._areasRenderer.addTo(this.map);
    this.areasLayer = L.layerGroup().addTo(this.map);
    this.map.createPane('projects'); this.map.getPane('projects').style.zIndex='500';
    this._projRenderer = L.svg({pane:'projects'}); this._projRenderer.addTo(this.map);
    this.projectLayer = L.layerGroup().addTo(this.map);
    this.map.createPane('inquiries'); this.map.getPane('inquiries').style.zIndex='510'; this.map.on('popupopen',(e:any)=>this.onPopupOpen(e));
    this.inqLayer = L.layerGroup().addTo(this.map);
    const FemaTiles:any = L.TileLayer.extend({ getTileUrl:function(coords:any){ const map=this._map; const ts=this.getTileSize(); const nw=map.unproject(L.point(coords.x*ts.x,coords.y*ts.y),coords.z); const se=map.unproject(L.point((coords.x+1)*ts.x,(coords.y+1)*ts.y),coords.z); const a=L.CRS.EPSG3857.project(nw), b=L.CRS.EPSG3857.project(se); const bbox=Math.min(a.x,b.x)+','+Math.min(a.y,b.y)+','+Math.max(a.x,b.x)+','+Math.max(a.y,b.y); return 'https://hazards.fema.gov/arcgis/rest/services/public/NFHL/MapServer/export?bbox='+bbox+'&bboxSR=3857&imageSR=3857&size='+ts.x+','+ts.y+'&dpi=96&format=png32&transparent=true&f=image'; } });
    this.femaLayer = new FemaTiles('', {tileSize:256, opacity:0.55, pane:'zoning', maxZoom:20, attribution:'Flood data Â© FEMA NFHL'});
    ZJURS.forEach((j:any)=>{ j._layer = L.imageOverlay(this.zoningAssetBase+j.file, j.bounds, {opacity:j.opacity, interactive:false, pane:'zoning'}); j._on = false; });
    this.map.createPane('worked'); this.map.getPane('worked').style.zIndex='470';
    this.map.createPane('sel'); this.map.getPane('sel').style.zIndex='480'; this.map.getPane('sel').style.pointerEvents='none';
    this.map.createPane('labels'); this.map.getPane('labels').style.zIndex='550'; this.map.getPane('labels').style.pointerEvents='none';
    this.map.createPane('ucdd'); this.map.getPane('ucdd').style.zIndex='360'; this.map.getPane('ucdd').style.pointerEvents='none';
    this._ucddRenderer=L.canvas({pane:'ucdd',padding:0.5});
    this.ucddLayer=L.layerGroup([],{pane:'ucdd'});
    this.parcelLayer = L.geoJSON(null,{ style:(ft:any)=>this.parcelStyle(ft), onEachFeature:(ft:any,layer:any)=>this.onFeat(ft,layer) }).addTo(this.map);
    this.selLayer = L.geoJSON(null,{ pane:'sel', style:{color:'#1565ff',weight:3,fillColor:'#4a90e2',fillOpacity:0.22} }).addTo(this.map);
    this.labelLayer = L.layerGroup().addTo(this.map);
    this.workedGeomLayer = L.geoJSON(null,{ pane:'worked', style:(ft:any)=>this.workedStyle(ft), onEachFeature:(ft:any,layer:any)=>this.onWorkedFeat(ft,layer) });
    this.hiLayer = L.geoJSON(null,{ style:{color:'#ff2d55',weight:3,fill:false} }).addTo(this.map);
    this.map.on('moveend',()=>{ clearTimeout(this.loadTimer); this.loadTimer=setTimeout(()=>this.maybeLoad(),250); });
    this.map.on('zoomend',()=>{ this.renderLabels(); });
    this.setStatus('Pan/zoom to your area â parcels load at zoom '+MINZOOM+'+');
    setTimeout(()=>{ try{ this.map.invalidateSize(); }catch(e){} this.loadParcels(); },400);
    window.addEventListener('resize',()=>{ clearTimeout(this.rzTimer); this.rzTimer=setTimeout(()=>{ try{ if(this.map) this.map.invalidateSize(); }catch(e){} },200); });
    this.loadZoning();
    this.loadWorked();
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

  // Only reload if the view left the already-loaded (padded) area â keeps popups open and avoids flicker on small pans / popup auto-pan.
  private maybeLoad(): void {
    this.loadUcdd();   // self-guards: only when Zoning on + a UCDD sublayer checked + view left cache
    const z=this.map.getZoom();
    if(z>=MINZOOM && this.loadedBounds && this.loadedZoom===z && this.loadedBounds.contains(this.map.getBounds())) return;
    this.loadParcels();
  }

  private loadParcels(): void {
    this.inflight.forEach((c)=>{ try{c.abort();}catch(e){} }); this.inflight=[];
    const mySeq = ++this.loadSeq;   // anti-flicker: a stale load must not clobber a newer one
    const legsrc = this.domElement.querySelector('#legsrc') as any;
    if(this.map.getZoom()<MINZOOM){ this.parcelLayer.clearLayers(); this.loadedBounds=null; this.loadedZoom=-1; this.setStatus('Zoom in to load parcels (zoom â¥ '+MINZOOM+')'); if(legsrc) legsrc.textContent='Active data: â'; return; }
    const srcs=this.activeSources();
    if(srcs.length===0){ this.parcelLayer.clearLayers(); this.loadedBounds=null; this.setStatus('No parcel source covers this view'); return; }
    if(legsrc) legsrc.textContent='Active data: '+srcs.map((s)=>s.label.replace(/^..? â /,'')).join(', ');
    const pb=this.map.getBounds().pad(0.2); this.loadedBounds=pb; this.loadedZoom=this.map.getZoom();
    const env=[pb.getWest(),pb.getSouth(),pb.getEast(),pb.getNorth()].join(',');
    this.setStatus('Loading parcelsâ¦');   // keep the OLD parcels on screen until the new set is ready (no flash)
    let got=0, done=0; const errs:string[]=[]; const acc:any[]=[];
    const short=(s:any)=>s.label.replace(/^..? â /,'');
    const finish=()=>{ if(mySeq!==this.loadSeq) return; if(done===srcs.length){ this.parcelLayer.clearLayers(); if(acc.length) this.parcelLayer.addData(acc); this.setStatus(got+' parcels'+(errs.length?'  Â· unavailable: '+errs.join('; '):'')); this.renderLabels(); } };
    const httpsPage = (typeof location!=='undefined' && location.protocol==='https:');
    srcs.forEach((s)=>{
      if(httpsPage && /^http:\/\//i.test(s.url)){ errs.push(short(s)+' (HTTP-only â needs HTTPS proxy)'); done++; finish(); return; }
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
    // Standalone MAP popup (NOT bound to the parcel layer) so a parcel reload can't close it â no re-clicking.
    layer.on('click',(ev:any)=>{
      if(this.splitState||this.areaState) return;
      const src=SOURCES.filter((s)=>s.id===feat.properties.__src)[0]||SOURCES[0];
      const n=normalize(feat.properties,src);
      const ll=(ev&&ev.latlng)||(layer.getBounds&&layer.getBounds().getCenter());
      if(this.workEdit){ this.openWorkPicker(n, ll, feat); return; }
      if(this.zoningEdit){ this.openZonePicker(n, ll, feat); return; }
      n.ucdd=this.ucddZonesAt(ll);
      this.selectParcel(feat, n);
      L.popup({maxWidth:320,autoPanPadding:[24,24]}).setLatLng(ll).setContent(this.popupHtml(n)).openOn(this.map);
    });
  }

  private ucddZonesAt(ll:any): any[] {
    const out:any[]=[]; if(!ll||!this._ucddCache) return out;
    const lng=ll.lng, lat=ll.lat;
    for(let i=0;i<UCDD_ZONING.length;i++){ const c=UCDD_ZONING[i]; const feats=this._ucddCache[c.key]; if(!feats||!feats.length) continue;
      for(let k=0;k<feats.length;k++){ if(featHitLL(feats[k],lng,lat)){ const zone=feats[k].properties[c.field]; out.push({label:c.label,zone:(zone==null?'':(''+zone)),name:ucddName(c.key,zone)}); break; } } }
    return out;
  }

  private popupHtml(n:any): string {
    const id='_p'+(this.pseq++); this.POP[id]=n;
    let rows='';
    const row=(k:string,v:any)=>{ if(v) rows+='<tr><td class="k">'+k+'</td><td>'+esc(v)+'</td></tr>'; };
    let owner=n.owner+(n.owner2?'; '+n.owner2:'');
    if(n.src.ownerWithheld && !owner) owner='<i>(owner not published by county)</i>';
    rows+='<tr><td class="k">Owner</td><td>'+(owner||'â')+'</td></tr>';
    row('Address',n.address); row('Parcel',n.pin);
    if(n.acres) row('Acres', (+n.acres? (+n.acres).toFixed(2):n.acres));
    if(n.subdiv) row('Subdiv', n.subdiv+(n.lot?'  Lot '+n.lot:''));
    if(n.zoning) row('Zoning', n.zoning);
    const zt=this.zoneByPin[pinKey(n.pin)]; if(zt){ const zj=jurById(zt.jur)||ZJURS[0]; if(zt.split&&zt.pieces){ row('Zone ('+(zt.jur||'')+')', zt.pieces.map((p:any)=>p.z||'blank').join(' / ')+' (split lot)'); } else { row('Zone ('+(zt.jur||'')+')', zt.zone+' â '+((zj.names&&zj.names[zt.zone])||'')+(zt.flood?' Â· Floodplain':'')); } }
    if(n.ucdd && n.ucdd.length){ for(let ui=0;ui<n.ucdd.length;ui++){ const uz=n.ucdd[ui]; row('Zone ('+uz.label+')', uz.zone+(uz.name?' â '+uz.name:'')+' Â· official zoning'); } }
    if(this.workView){ const wj=this.workedByPin[pinKey(n.pin)]; if(wj&&wj.length){ let wb=''; for(let wi=0;wi<wj.length;wi++){ const w=wj[wi]; wb+='<div class="dls-wk"><b>'+esc(w.job||'')+'</b>'+(w.name?' '+esc(w.name):'')+(w.job?' <a class="dls-pop-a" href="#" data-act="wfolder" data-arg="'+esc(w.job)+'">Open folder &#8599;</a>':'')+'</div>'; } rows+='<tr><td class="k">Surveyed</td><td>'+wb+'</td></tr>'; } }
    if(n.deedBook||n.deedPage) row('Deed','Bk '+n.deedBook+' Pg '+n.deedPage);
    else if(n.legalref) row('Deed ref', n.legalref);
    else if(n.deedref) row('Deed ref', n.deedref);
    return '<div class="lp"><h3>'+esc(n.county||'')+(n.state?', '+n.state:'')+'</h3><div class="co">'+esc(n.src.label)+'</div><table>'+rows+'</table>'+'<div class="valbox" id="vb_'+id+'" data-gis="'+esc(n.gislink||'')+'" data-acres="'+esc(n.acres||'')+'"></div>'+'<div style="margin:6px 0"><button class="zbtn2" data-act="print" data-id="'+id+'">&#128424; Print lot sheet</button></div>'+this.deedSection(n,id)+'</div>';
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
      out+='<div class="note">Opens US Title Search (your session) and surfaces the latest warranty-deed book/page to enter â that site has no direct deep-link.</div>';
    } else if(n.state==='TN'){
      out+='<span class="note">No deed site mapped for '+esc(n.county)+' â use the assessor link.</span><br/>';
    } else if(n.state==='KY'){
      out+='<a class="btn ts payg" href="'+TS_BASE+'countySelect.php" target="_blank" rel="noopener">TitleSearcher (KY) Â· pick county</a> <span class="note">KY not yet mapped</span><br/>';
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

  private onPopupOpen(e:any): void {
    try{
      const root:any = (e && e.popup && e.popup.getElement) ? e.popup.getElement() : null;
      if(!root) return;
      const box:any = root.querySelector('.valbox[data-gis]');
      if(!box || box.getAttribute('data-done')==='1') return;
      const gis = box.getAttribute('data-gis')||'';
      if(!gis || !this.workerUrl){ box.style.display='none'; return; }
      box.setAttribute('data-done','1');
      box.innerHTML = '<div class="vbx-hd">Sale &amp; assessor value</div><div style="color:#64748b;font-size:11px">Loading…</div>';
      const acres = parseFloat(box.getAttribute('data-acres')||'') || 0;
      const self = this;
      fetch(this.workerUrl+'?gislink='+encodeURIComponent(gis)).then((r:any)=>r.json())
        .then((d:any)=>{ if(!d || !d.ok){ box.style.display='none'; return; } const h=self.valBoxHtml(d, acres); if(h){ box.innerHTML=h; } else { box.style.display='none'; } })
        .catch(()=>{ box.style.display='none'; });
    }catch(err){}
  }

  private valBoxHtml(d:any, acres:number): string {
    const v:any = d.values || {};
    const a:any = d.acres || {};
    const ac:number = acres>0 ? acres : ((a.calc && a.calc>0) ? a.calc : ((a.deed && a.deed>0) ? a.deed : 0));
    const grp = (x:any)=>{ const n=Math.round(Number(x)); if(!isFinite(n)) return ''; return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ','); };
    const money = (x:any)=> (x==null||x==='') ? '' : ('$'+grp(x));
    const perAc = (val:any)=> (ac>0 && Number(val)>0) ? ('$'+grp(Number(val)/ac)+'/ac') : '';
    let rows = '';
    const line = (k:string, val:string, sub:string)=>{ if(!val) return; rows += '<tr><td class="vk">'+k+'</td><td>'+val+(sub?(' <span class="vsub">'+sub+'</span>'):'')+'</td></tr>'; };

    const qs:any = d.qualifiedSale;
    const bestSale:any = d.best;
    if(qs && qs.price){
      const vac = /^V/i.test(qs.vi||'');
      line('Last sale', esc(qs.price)+(qs.date?(' · '+esc(qs.date)):''), 'WD · '+esc(qs.qual||'Accepted')+(vac?' · vacant':' · improved'));
      const sp = perAc(Number(String(qs.price).replace(/[^0-9.]/g,'')));
      if(sp) line('Sale $/acre', sp, vac?'':'(incl. buildings)');
    } else if(bestSale && bestSale.price && bestSale.price!=='$0'){
      line('Last sale', esc(bestSale.price)+(bestSale.date?(' · '+esc(bestSale.date)):''), esc(bestSale.type||'')+(bestSale.qual?(' · '+esc(bestSale.qual)):''));
    }

    line('Land value', money(v.landMarketValue), perAc(v.landMarketValue));
    line('Total appraisal', money(v.totalMarketAppraisal), (v.improvementValue?('impr '+money(v.improvementValue)):''));
    line('Assessment', money(v.assessment), (v.assessmentPct?(v.assessmentPct+'%'):''));

    if(!rows) return '';

    let callout='';
    const landPer = perAc(v.landMarketValue);
    if(landPer){ callout = '<div class="vbx-per">'+landPer+' <span class="vsub">assessor land value / acre'+(ac>0?(' · '+(Math.round(ac*100)/100)+' ac'):'')+'</span></div>'; }

    let hist='';
    const sales:any[] = d.sales || [];
    if(sales.length){
      let hr='';
      for(let i=0;i<sales.length;i++){ const sx:any=sales[i]; hr += '<tr><td>'+esc(sx.date||'')+'</td><td>'+esc(sx.price||'')+'</td><td>'+esc((sx.type||'').replace(/ -.*/,''))+'</td><td>'+esc((sx.qual||'').replace(/ -.*/,''))+'</td></tr>'; }
      hist = '<details class="vbx-hist"><summary>Sale history ('+sales.length+')</summary><table class="vbx-ht"><tr><th>Date</th><th>Price</th><th>Type</th><th>Qual</th></tr>'+hr+'</table></details>';
    }

    const css = '<style>'
      + '.valbox{border-top:1px solid #e2e8f0;margin-top:6px;padding-top:6px;}'
      + '.valbox .vbx-hd{font-size:11px;color:#64748b;margin-bottom:3px;font-weight:700;text-transform:uppercase;letter-spacing:.02em;}'
      + '.valbox table{border-collapse:collapse;font-size:12px;width:100%;} .valbox td{padding:1px 6px 1px 0;vertical-align:top;}'
      + '.valbox td.vk{color:#64748b;white-space:nowrap;} .valbox .vsub{color:#94a3b8;font-size:10px;}'
      + '.valbox .vbx-per{margin:5px 0 2px;font-size:15px;font-weight:800;color:#0f766e;}'
      + '.valbox .vbx-hist{margin-top:4px;font-size:11px;} .valbox .vbx-hist summary{cursor:pointer;color:#475569;}'
      + '.valbox .vbx-ht{margin-top:3px;} .valbox .vbx-ht th{color:#64748b;text-align:left;font-weight:600;padding-right:6px;} .valbox .vbx-ht td{padding-right:6px;}'
      + '</style>';

    return css + '<div class="vbx-hd">Sale &amp; assessor value</div>' + callout + '<table>'+rows+'</table>' + hist;
  }

    // ======================= Work history layer (surveyed parcels) =======================
  private setWorkMode(v:string): void {
    this.workView=(v==='view'||v==='edit');
    this.workEdit=(v==='edit');
    if(this.workEdit && this.zoningEdit){ this.zoningEdit=false; const _zs=this.domElement.querySelector('#zmode') as any; if(_zs) _zs.value='view'; this.zoningView=true; }
    if(this.workView){ if(this.workedGeomLayer && !this.map.hasLayer(this.workedGeomLayer)) this.workedGeomLayer.addTo(this.map); this.ensureWorkedGeom(); } else { if(this.workedGeomLayer && this.map.hasLayer(this.workedGeomLayer)) this.map.removeLayer(this.workedGeomLayer); }
    this.buildLegend();
    this.restyleParcels();
    if(v==='edit') this.setStatus('Work history EDIT â click a lot, then pick its WIP job (or type an older one) to mark it surveyed.');
    else if(v==='view') this.setStatus('Work history VIEW â surveyed lots are outlined and filled green. '+this._workCount+' lots.');
    else this.setStatus('Work history off.');
    this.map.closePopup();
  }

  private workColor(w:any): any {
    if(this._workColorMode==='year' || this._workColorMode==='month'){
      let yr=0; for(let i=0;i<w.length;i++){ const t=String(w[i].job||''); if(/^[0-9]{6}/.test(t)){ const y=2000+parseInt(t.substring(0,2),10); if(y>yr) yr=y; } }
      const pal:any={2021:'#7c3aed',2022:'#2563eb',2023:'#0891b2',2024:'#16a34a',2025:'#ca8a04',2026:'#dc2626'};
      const c=pal[yr]||'#16a34a'; return {outline:'#0a3d1f',fill:c};
    }
    return {outline:'#0a5a27',fill:'#16a34a'};
  }

  private countWorkedLots(m:any): number { return Object.keys(m).length; }

  private loadWorked(): void {
    if(this._workLoaded) return;
    this.spGet(this.workedApi()+"/items?$select=Id,ParcelID,JobNumber,JobName,FolderURL,County,TaxMap,ParcelNo,WIPItemId,Source,Notes&$top=5000").then((d:any)=>{
      const items=(d&&d.value)||[]; const m:any={}; const ids:any={}; let unres=0;
      for(let i=0;i<items.length;i++){ const it=items[i]; if(it.WIPItemId!=null) ids[it.WIPItemId]=1; if(it.ParcelID){ const k=pinKey(it.ParcelID); if(!m[k]) m[k]=[]; m[k].push({id:it.Id,job:it.JobNumber,name:it.JobName,folder:it.FolderURL,county:it.County,src:it.Source,wipId:it.WIPItemId,notes:it.Notes}); } else { unres++; } }
      this.workedByPin=m; this.workedWipIds=ids; this._workCount=this.countWorkedLots(m); this._workUnresolved=unres; this._workLoaded=true;
      const fc:any={}; for(let fi=0;fi<items.length;fi++){ const it2=items[fi]; if(it2.JobNumber&&it2.FolderURL) fc[it2.JobNumber]=it2.FolderURL; } this._folderCache=fc;
      this._workGeomLoaded=false; if(this.workView){ this.ensureWorkedGeom(); this.buildLegend(); }
      this.selfHealWorked();
    }).catch(()=>{});
  }

  private selfHealWorked(): void {
    const url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists(guid'" + this.projectsListGuid + "')/items?$top=500&$select=Id,Title,JobLabel,FolderURL,County,Tax_x0020_Map,Parcel_x0020_Number";
    this.spGet(url).then((d:any)=>{
      const items=(d&&d.value)||[]; const todo:any[]=[]; const pk:any[]=[];
      for(let i=0;i<items.length;i++){ const x=items[i];
        pk.push({num:(x.Title==null?'':String(x.Title)),name:x.JobLabel||'',folder:x.FolderURL||'',county:x.County||''});
        if(this.workedWipIds[x.Id]) continue;
        if(!x.Tax_x0020_Map||!x.Parcel_x0020_Number) continue;
        todo.push(x);
      }
      this._wipPick=pk;
      if(todo.length) this.healNext(todo,0,0);
    }).catch(()=>{});
  }

  private healNext(todo:any[], i:number, made:number): void {
    if(i>=todo.length){ if(made>0){ this._workLoaded=false; this.loadWorked(); } return; }
    const x=todo[i]; const self=this;
    const ccm=String(x.County||'').match(/([0-9]{3})/); const cc=ccm?ccm[1]:'';
    const map=wkNormMap(x.Tax_x0020_Map); const p5=wkNorm5(x.Parcel_x0020_Number);
    const next=(n:number)=>self.healNext(todo,i+1,made+n);
    if(!cc||!p5){ this.createWorked(x,'','Auto-unresolved','missing county/parcel',()=>next(1)); return; }
    const q=WK_TN_SVC+"?where="+encodeURIComponent("PARCELID LIKE '"+cc+" "+map+"%' AND PARCELID LIKE '%"+p5+" %'")+"&outFields=PARCELID&returnGeometry=false&resultRecordCount=50&f=json";
    fetch(q).then((r:any)=>r.json()).then((j:any)=>{ const fs=(j&&j.features)||[]; const mm:any[]=[]; for(let k=0;k<fs.length;k++){ const pid=fs[k].attributes.PARCELID; if(pid.substring(0,3)===cc&&pid.substring(4,8).replace(/\s+$/,'')===map&&pid.substring(11,16)===p5) mm.push(pid); } if(mm.length===1){ self.createWorked(x,mm[0],'Auto','',()=>next(1)); } else { self.createWorked(x,'','Auto-unresolved',(mm.length>1?('group ambiguous: '+mm.length+' lots'):'no exact parcel match'),()=>next(1)); } }).catch(()=>next(0));
  }

  private createWorked(x:any, parcelId:string, source:string, notes:string, cb:any): void {
    const body:any={Title:(parcelId||String(x.Title||x.Id)),ParcelID:parcelId||'',JobNumber:String(x.Title||''),JobName:x.JobLabel||'',FolderURL:x.FolderURL||'',County:String(x.County||''),TaxMap:String(x.Tax_x0020_Map||''),ParcelNo:String(x.Parcel_x0020_Number||''),WIPItemId:x.Id,Source:source,Notes:notes||''};
    this.spPost(this.workedApi()+'/items', body).then(()=>{ this.workedWipIds[x.Id]=1; if(cb)cb(); }).catch(()=>{ if(cb)cb(); });
  }

  private buildWorkPanel(): void {
    const el=this.domElement.querySelector('#workHost') as any; if(!el) return;
    let h='<div class="zrow"><span class="zsw" style="background:#16a34a;border:2px solid #0a5a27"></span> Surveyed lot</div>';
    h+='<div class="zdisc">'+this._workCount+' parcels marked surveyed (live from DLS Worked Parcels). ';
    if(this._workUnresolved) h+='<b>'+this._workUnresolved+'</b> job(s) need a lot picked (review worklist). ';
    h+='New WIP jobs with a Tax Map &amp; Parcel are added automatically when the map loads. Switch to <b>Edit</b> to mark a lot surveyed by hand.</div>';
    h+='<div class="lp-sub" data-wyearsec><div class="lp-subhd"><span class="tw">&#9662;</span> Filter by year <span class="pp-all" data-wyall>all</span></div><div class="lp-subbd"><div id="wYear"></div></div></div>';
    el.innerHTML=h;
    const c=this.domElement.querySelector('#wCount'); if(c) c.textContent=this._workShownCount()+' lots'; this.renderWorkYearSection(); const _wself=this; const _wh=el.querySelector('[data-wyearsec] .lp-subhd') as any; if(_wh) _wh.addEventListener('click',function(e:any){ const a=(e.target&&e.target.closest)?e.target.closest('[data-wyall]'):null; if(a){ _wself.workYearAllOn(); return; } const sub=this.parentNode; if(sub) sub.classList.toggle('coll'); });
  }

  private openWorkPicker(n:any, ll:any, _feat?:any): void {
    const pin=pinKey(n.pin); if(!pin){ this.setStatus('This parcel has no ID â cannot mark it.'); return; }
    this.wTarget={pin:pin, raw:String(n.pin).trim(), ll:ll};
    this.renderWorkPicker();
  }

  private renderWorkPicker(): void {
    const t=this.wTarget; if(!t) return;
    const cur=this.workedByPin[t.pin]||[];
    let curHtml=''; for(let i=0;i<cur.length;i++){ const w=cur[i]; curHtml+='<div class="wk-cur"><b>'+esc(w.job||'')+'</b> '+esc(w.name||'')+' <button class="zp-clear" data-act="wdel" data-arg="'+w.id+'">remove</button></div>'; }
    let opts='<option value="">â choose a WIP job â</option>'; for(let i=0;i<this._wipPick.length;i++){ const p=this._wipPick[i]; opts+='<option value="'+i+'">'+esc((p.num||'')+(p.name?(' '+p.name):''))+'</option>'; }
    const html='<div class="zp"><div class="zp-h">Mark surveyed</div>'
      +'<div class="zp-pin">Parcel: <b>'+esc(t.raw)+'</b></div>'
      +(curHtml?('<div class="zp-cur">Already recorded:'+curHtml+'</div>'):'')
      +'<div class="wk-row"><select id="wkJob">'+opts+'</select></div>'
      +'<div class="wk-or">â or type an older (pre-WIP) job â</div>'
      +'<div class="wk-row"><input id="wkNum" placeholder="Job #"/></div>'
      +'<div class="wk-row"><input id="wkName" placeholder="Job name"/></div>'
      +'<div class="wk-row"><input id="wkFolder" placeholder="Folder URL (optional)"/></div>'
      +'<div style="margin-top:6px"><button class="zbtn2" data-act="wsave">Mark surveyed</button></div>'
      +'<div class="zp-note">Marks this exact lot surveyed and links the job folder. Full history â adding a second job keeps the first.</div></div>';
    L.popup({maxWidth:300,autoPanPadding:[24,24]}).setLatLng(t.ll||this.map.getCenter()).setContent(html).openOn(this.map);
  }

  private saveWorked(): void {
    const t=this.wTarget; if(!t||!t.pin) return;
    const jobSel=this.domElement.querySelector('#wkJob') as any;
    let num='', name='', folder='', county='';
    if(jobSel && jobSel.value!==''){ const p=this._wipPick[+jobSel.value]; if(p){ num=p.num; name=p.name; folder=p.folder; county=p.county; } }
    if(!num){ const ni=this.domElement.querySelector('#wkNum') as any; const na=this.domElement.querySelector('#wkName') as any; const fo=this.domElement.querySelector('#wkFolder') as any; num=ni?(ni.value||'').trim():''; name=na?(na.value||'').trim():''; folder=fo?(fo.value||'').trim():''; }
    if(!num){ this.setStatus('Pick a WIP job or type a job number first.'); return; }
    const body:any={Title:t.raw,ParcelID:t.raw,JobNumber:num,JobName:name,FolderURL:folder,County:county,TaxMap:'',ParcelNo:'',Source:'Manual',Notes:'tagged on map'};
    this.spPost(this.workedApi()+'/items', body).then((r:any)=>{ if(r.status>=200&&r.status<300) return r.json(); throw new Error('HTTP '+r.status); }).then((d:any)=>{ const k=t.pin; if(!this.workedByPin[k]) this.workedByPin[k]=[]; this.workedByPin[k].push({id:d&&d.Id,job:num,name:name,folder:folder,county:county,src:'Manual'}); this._workCount=this.countWorkedLots(this.workedByPin); this.restyleParcels(); if(this.workView) this.buildLegend(); this.setStatus('Marked '+t.raw+' surveyed â '+num); this.map.closePopup(); }).catch((e:any)=>this.setStatus('Save failed: '+e));
  }

  private clearWorked(rowId:string): void {
    const t=this.wTarget; const id=parseInt(rowId,10); if(!id) return;
    this.spPost(this.workedApi()+'/items('+id+')', null, {'X-HTTP-Method':'DELETE','IF-MATCH':'*'}).then((r:any)=>{ if(r.status>=200&&r.status<300){ if(t&&this.workedByPin[t.pin]){ const arr=this.workedByPin[t.pin]; const na:any[]=[]; for(let i=0;i<arr.length;i++){ if(arr[i].id!==id) na.push(arr[i]); } if(na.length) this.workedByPin[t.pin]=na; else delete this.workedByPin[t.pin]; } this._workCount=this.countWorkedLots(this.workedByPin); this.restyleParcels(); if(this.workView) this.buildLegend(); if(t) this.renderWorkPicker(); this.setStatus('Removed.'); } else this.setStatus('Remove failed ('+r.status+')'); }).catch((e:any)=>this.setStatus('Remove failed: '+e));
  }

  // ================= v20: selection highlight =================
  private selectParcel(feat:any, n:any): void {
    try{ this.selFeat=feat; this.selN=n; if(this.selLayer){ this.selLayer.clearLayers(); if(feat&&feat.geometry) this.selLayer.addData(feat); } }catch(e){}
  }
  private clearSelection(): void { this.selFeat=null; this.selN=null; if(this.selLayer) this.selLayer.clearLayers(); }

  // ================= v20: parcel-number labels =================
  private parcelLabelText(ft:any): string {
    const src=SOURCES.filter((s)=>s.id===ft.properties.__src)[0]||SOURCES[0];
    const pno=pick(ft.properties, src.f.parcelno);
    if(pno){ const m=pno.match(/^0*([0-9]+)\.([0-9]+)$/); if(m) return m[1]+'.'+m[2]; return pno.replace(/^0+(?=[0-9])/,''); }
    const pid=pick(ft.properties, src.f.pin);
    if(pid && /^[0-9]{3} /.test(pid) && pid.length>=16){ const ip=parseInt(pid.substring(11,14),10); const dp=pid.substring(14,16); if(!isNaN(ip)) return ip+'.'+dp; }
    return '';
  }
  private renderLabels(): void {
    if(!this.labelLayer) return; this.labelLayer.clearLayers();
    if(this.map.getZoom()<LABELZOOM) return;
    const self=this;
    try{ this.parcelLayer.eachLayer(function(ly:any){ const ft=ly.feature; if(!ft) return; const txt=self.parcelLabelText(ft); if(!txt) return; const r=outerRing(ft.geometry); if(!r||!r.length) return; const c=centroid(r); const mk=L.marker([c[1],c[0]],{pane:'labels',interactive:false,icon:L.divIcon({className:'dls-plabel',html:txt,iconSize:[0,0]})}); self.labelLayer.addLayer(mk); }); }catch(e){}
  }

  // ================= v20: worked-lots overview (by-ID, any zoom) =================
  private workedStyle(ft:any): any { const w=this.workedByPin[pinKey(ft.properties.PARCELID)]; if(!this.workedVisibleByYear(w||[])) return {stroke:false,fill:false,pane:'worked'}; const c=this.workColor(w||[]); return {stroke:true,color:c.outline,weight:2.5,fill:true,fillColor:c.fill,fillOpacity:0.45,pane:'worked'}; }
  private workedYearsOf(w:any): any { const ys:any={}; if(w){ for(let i=0;i<w.length;i++){ const t=String(w[i].job||''); if(/^[0-9]{6}/.test(t)){ ys['20'+t.substring(0,2)]=1; } else { ys['(blank)']=1; } } } return ys; }
  private workedVisibleByYear(w:any): boolean { if(!w||!w.length) return true; const ys=this.workedYearsOf(w); const ks=Object.keys(ys); if(!ks.length) return true; for(let i=0;i<ks.length;i++){ if(this.pWorkYearOn[ks[i]]!==false) return true; } return false; }
  private applyWorkedYearFilter(): void { try{ if(this.workedGeomLayer) this.workedGeomLayer.setStyle((ft:any)=>this.workedStyle(ft)); }catch(e){} }
  private workedYearCounts(): any { const m:any={}; const pins=Object.keys(this.workedByPin); for(let i=0;i<pins.length;i++){ const ys=this.workedYearsOf(this.workedByPin[pins[i]]); const ks=Object.keys(ys); for(let j=0;j<ks.length;j++){ m[ks[j]]=(m[ks[j]]||0)+1; } } return m; }
  private _workShownCount(): number { const pins=Object.keys(this.workedByPin); let n=0; for(let i=0;i<pins.length;i++){ if(this.workedVisibleByYear(this.workedByPin[pins[i]])) n++; } return n; }
  private renderWorkYearSection(): void { const el=this.domElement.querySelector('#wYear') as any; if(!el) return; const counts=this.workedYearCounts(); const keys=Object.keys(counts); for(let i=0;i<keys.length;i++){ if(this.pWorkYearOn[keys[i]]===undefined) this.pWorkYearOn[keys[i]]=true; } keys.sort((a:any,b:any)=>b.localeCompare(a)); const self=this; el.innerHTML=''; for(let i=0;i<keys.length;i++){ const k=keys[i]; const row=document.createElement('div'); row.className='pp-row'+(this.pWorkYearOn[k]===false?' off':''); row.innerHTML='<span class="pp-sq"></span><span class="pp-nm">'+esc(k)+'</span><span class="pp-n">'+counts[k]+'</span>'; row.onclick=function(){ self.pWorkYearOn[k]=self.pWorkYearOn[k]===false; self.applyWorkedYearFilter(); self.renderWorkYearSection(); self.updateWorkCountLabel(); }; el.appendChild(row); } }
  private workYearAllOn(): void { const ks=Object.keys(this.pWorkYearOn); for(let i=0;i<ks.length;i++){ this.pWorkYearOn[ks[i]]=true; } this.applyWorkedYearFilter(); this.renderWorkYearSection(); this.updateWorkCountLabel(); }
  private updateWorkCountLabel(): void { const c=this.domElement.querySelector('#wCount') as any; if(c) c.textContent=this._workShownCount()+' lots'; }
  private onWorkedFeat(feat:any, layer:any): void {
    layer.on('click',(ev:any)=>{
      if(this.splitState||this.areaState) return;
      const n=normalize(feat.properties, SOURCES[0]);
      const ll=(ev&&ev.latlng)||(layer.getBounds&&layer.getBounds().getCenter());
      if(this.workEdit){ this.openWorkPicker(n, ll, feat); return; }
      if(this.zoningEdit){ this.openZonePicker(n, ll, feat); return; }
      n.ucdd=this.ucddZonesAt(ll);
      this.selectParcel(feat, n);
      L.popup({maxWidth:320,autoPanPadding:[24,24]}).setLatLng(ll).setContent(this.popupHtml(n)).openOn(this.map);
    });
  }
  private ensureWorkedGeom(): void {
    if(this._workGeomLoaded || !this.workedGeomLayer) return;
    this._workGeomLoaded=true; const self=this;
    this.spGet(this.workedApi()+"/items?$select=ParcelID&$top=5000").then((d:any)=>{
      const seen:any={}; const list:string[]=[]; const items=(d&&d.value)||[];
      for(let i=0;i<items.length;i++){ const p=items[i].ParcelID; if(p&&!seen[pinKey(p)]){ seen[pinKey(p)]=1; list.push(p); } }
      self.setStatus('Loading '+list.length+' surveyed lotsâ¦'); self.fetchWorkedChunks(list,0,[]);
    }).catch(()=>{ self._workGeomLoaded=false; });
  }
  private fetchWorkedChunks(list:string[], i:number, acc:any[]): void {
    const self=this;
    if(i>=list.length){ try{ self.workedGeomLayer.clearLayers(); if(acc.length) self.workedGeomLayer.addData(acc); }catch(e){} self.setStatus(acc.length+' surveyed lots shown.'); return; }
    const chunk=list.slice(i,i+25); const inlist=chunk.map((x)=>"'"+x.replace(/'/g,"''")+"'").join(',');
    const url=WK_TN_SVC+'?where='+encodeURIComponent('PARCELID IN ('+inlist+')')+'&outFields='+encodeURIComponent(outFieldsFor(SOURCES[0]))+'&returnGeometry=true&outSR=4326&resultRecordCount=2000&f=json';
    this.arcgisFetch(url).then((d:any)=>{ const feats=esriToFeatures(d); for(let k=0;k<feats.length;k++){ feats[k].properties.__src='tn'; acc.push(feats[k]); } self.fetchWorkedChunks(list,i+25,acc); }).catch(()=>{ self.fetchWorkedChunks(list,i+25,acc); });
  }

  // ================= v20: folder resolve by Job# (hybrid, no PA) =================
  private openFolderByJob(jobNo:string): void {
    if(!jobNo){ this.setStatus('No job number on this record.'); return; }
    const w=window.open('','_blank'); try{ if(w) w.document.write('<p style="font:14px/1.4 sans-serif;padding:18px;color:#333">Opening job folderâ¦</p>'); }catch(e){}
    const open=(url:string)=>{ try{ if(w) w.location=url; else window.open(url,'_blank'); }catch(e){} };
    const cached=this._folderCache[jobNo]; const self=this;
    if(cached){ const m=String(cached).match(/https?:\/\/[^/]+(\/.*)$/); const sr=m?decodeURIComponent(m[1]):''; if(sr){ const ex=this.context.pageContext.web.absoluteUrl+"/_api/web/GetFolderByServerRelativeUrl('"+sr.replace(/'/g,"''")+"')/Exists"; this.spGet(ex).then((d:any)=>{ if(d&&(d.value===true||d.Exists===true)){ open(cached); self.setStatus('Opened job folder.'); } else self.resolveFolder(jobNo,open); }).catch(()=>self.resolveFolder(jobNo,open)); return; } }
    this.resolveFolder(jobNo, open);
  }
  private resolveFolder(jobNo:string, open:any): void {
    const base=this.context.pageContext.web.absoluteUrl; const host=base.replace(/\/sites\/.*$/,'');
    const spm=base.match(/(\/sites\/[^/]+)/); const sp=spm?spm[1]:''; const self=this;
    const yr='20'+jobNo.substring(0,2); const mm=jobNo.substring(2,4);
    const parents=[yr, yr+'/'+mm, 'Archive', 'Archive/'+yr, 'Archive/'+yr+'/'+mm];
    const tryNext=(idx:number)=>{
      if(idx>=parents.length){ open(base+"/Shared Documents/Forms/AllItems.aspx"); self.setStatus('Folder for '+jobNo+' not found â opened the library.'); return; }
      const parentRel=(sp+'/Shared Documents/'+parents[idx]).replace(/ /g,'%20');
      const u=base+"/_api/web/GetFolderByServerRelativeUrl('"+parentRel.replace(/'/g,"''")+"')/Folders?$select=Name,ServerRelativeUrl&$filter=startswith(Name,'"+jobNo+"')&$top=1";
      self.spGet(u).then((d:any)=>{ const v=(d&&d.value)||[]; if(v.length){ const fu=host+v[0].ServerRelativeUrl; self._folderCache[jobNo]=fu; open(fu); self.setStatus('Opened folder for '+jobNo); } else tryNext(idx+1); }).catch(()=>tryNext(idx+1));
    };
    tryNext(0);
  }

  // ================= v20: print lot sheet (8.5x11, centered + neighbors) =================
  private openPrintSheet(n:any, feat:any): void {
    if(!feat||!feat.geometry){ this.setStatus('Click a lot first, then Print.'); return; }
    const old=document.getElementById('dls-print') as any; if(old){ try{ old.parentNode.removeChild(old); }catch(e){} }
    const host=document.createElement('div'); host.id='dls-print'; host.className='dls-print-modal';
    const short=this.parcelLabelText(feat)||n.pin||'';
    const dt=new Date(); const dstr=(dt.getMonth()+1)+'/'+dt.getDate()+'/'+dt.getFullYear();
    host.innerHTML=''
      +'<div class="dpx"><button id="dlsPrintGo">&#128424; Print</button><button id="dlsPrintClose">Close</button></div>'
      +'<div class="dlsheet">'
      +'<div class="dlhd">'+esc(n.county||'')+' County &mdash; Parcel: '+esc(short)+'</div>'
      +'<div id="dlsPrintMap" class="dlmap"></div>'
      +'<div class="dlft"><div class="dlinfo">'
      +'<div><b>Date:</b> '+esc(dstr)+'</div>'
      +'<div><b>County:</b> '+esc(n.county||'')+'</div>'
      +'<div><b>Owner:</b> '+esc((n.owner||'')+(n.owner2?'; '+n.owner2:''))+'</div>'
      +'<div><b>Address:</b> '+esc(n.address||'')+'</div>'
      +'<div><b>Parcel ID:</b> '+esc(n.pin||'')+'</div>'
      +'<div><b>Deeded Acreage:</b> '+esc(n.acres||'0')+'</div>'
      +'</div><div class="dlscale" id="dlsPrintScale"></div></div>'
      +'</div>';
    document.body.appendChild(host); const self=this;
    (host.querySelector('#dlsPrintClose') as any).onclick=()=>{ try{ if(self._printMap){ self._printMap.remove(); self._printMap=null; } if(host.parentNode) host.parentNode.removeChild(host); }catch(e){} };
    (host.querySelector('#dlsPrintGo') as any).onclick=()=>{ try{ window.print(); }catch(e){} };
    const pm=L.map(host.querySelector('#dlsPrintMap'),{zoomControl:false,attributionControl:false,minZoom:6,maxZoom:21}); this._printMap=pm;
    L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}',{maxNativeZoom:19,maxZoom:21}).addTo(pm);
    const r=outerRing(feat.geometry); const c=r?centroid(r):null;
    const gj=L.geoJSON(feat,{interactive:false,style:{color:'#1565ff',weight:3,fillColor:'#4a90e2',fillOpacity:0.25}}).addTo(pm);
    if(c) pm.setView([c[1],c[0]],17);
    let _loaded=false;
    const frame=()=>{ try{ pm.invalidateSize(); pm.fitBounds(gj.getBounds(),{padding:[60,60],maxZoom:21}); if(!_loaded){ _loaded=true; self.loadPrintParcels(pm,feat); } self.updatePrintScale(pm,host); }catch(e){} };
    pm.on('moveend zoomend',()=>self.updatePrintScale(pm,host));
    setTimeout(frame,300); setTimeout(frame,1000);
  }
  private loadPrintParcels(pm:any, subj:any): void {
    const b=pm.getBounds().pad(0.15); const env=[b.getWest(),b.getSouth(),b.getEast(),b.getNorth()].join(','); const src=SOURCES[0]; const self=this;
    const url=src.url+'?'+qs({where:'1=1',geometry:env,geometryType:'esriGeometryEnvelope',inSR:4326,spatialRel:'esriSpatialRelIntersects',outFields:outFieldsFor(src),returnGeometry:true,outSR:4326,resultRecordCount:2000,f:'json'});
    this.arcgisFetch(url).then((d:any)=>{ const feats=esriToFeatures(d);
      L.geoJSON({type:'FeatureCollection',features:feats} as any,{interactive:false,style:{color:'#d08b27',weight:1,fill:false}}).addTo(pm);
      for(let i=0;i<feats.length;i++){ feats[i].properties.__src='tn'; const txt=self.parcelLabelText(feats[i]); if(!txt) continue; const rr=outerRing(feats[i].geometry); if(!rr||!rr.length) continue; const cc=centroid(rr); L.marker([cc[1],cc[0]],{interactive:false,icon:L.divIcon({className:'dls-plabel print',html:txt,iconSize:[0,0]})}).addTo(pm); }
      L.geoJSON(subj,{interactive:false,style:{color:'#1565ff',weight:3,fillColor:'#4a90e2',fillOpacity:0.28}}).addTo(pm);
    }).catch(()=>{});
  }
  private updatePrintScale(pm:any, host:any): void {
    try{ const z=pm.getZoom(); const lat=pm.getCenter().lat; const mpp=156543.03392*Math.cos(lat*Math.PI/180)/Math.pow(2,z);
      const ratio=Math.round(mpp/0.0002645833); const m2ft=3.28084; const targetFt=mpp*150*m2ft;
      const nice=[50,100,200,300,500,1000,2000,5000]; let ft=nice[0]; for(let i=0;i<nice.length;i++){ if(nice[i]<=targetFt) ft=nice[i]; }
      const px=Math.round((ft/m2ft)/mpp); const el=host.querySelector('#dlsPrintScale');
      if(el) el.innerHTML='<div class="sbar"><div class="sb" style="width:'+px+'px"></div><span>'+ft+' ft</span></div><div class="srat">Scale &asymp; 1:'+ratio+'</div>';
    }catch(e){}
  }

  private onAct(act:string, id:string, arg:string): void {
    if(act==='deedGo') this.deedGo(id);
    else if(act==='deedName') this.deedName(id);
    else if(act==='deedGoUS') this.deedGoUS(id);
    else if(act==='cpf'){ const e=this.POP[id]; if(e) this.copyText(e[arg]||''); }
    else if(act==='zset') this.saveZone(arg);
    else if(act==='zjur') this.switchZoneJur(arg);
    else if(act==='zclear') this.clearZone();
    else if(act==='zsplitopen') this.startSplitMenu();
    else if(act==='zsplit') this.beginSplit(arg);
    else if(act==='zpz'){ const p=(arg||'').split('|'); this.setPieceZone(+p[0], p[1]); }
    else if(act==='zsplitsave') this.saveSplit();
    else if(act==='zsplitcancel') this.cancelSplit();
    else if(act==='zareafinish') this.finishAreaDraw();
    else if(act==='zareacancel') this.cancelAreaDraw();
    else if(act==='zareasave') this.saveArea(arg);
    else if(act==='wsave') this.saveWorked();
    else if(act==='wdel') this.clearWorked(arg);
    else if(act==='print') this.openPrintSheet(this.selN||this.POP[id], this.selFeat);
    else if(act==='wfolder') this.openFolderByJob(arg);
    else if(act==='qfolder') this.openQuoteFolder(arg);
  }

  private openDeferred(): any { const w=window.open('','_blank'); try{ if(w) w.document.write('<p style="font:14px/1.4 sans-serif;padding:18px;color:#333">Looking up the latest deedâ¦</p>'); }catch(e){} return w; }
  private tsCountyThen(w:any,cnum:string,url:string): void { if(!w) return; w.location=TS_BASE+'countySearchPage.php?cnum='+cnum; setTimeout(()=>{ try{w.location=url;}catch(e){} },1600); }

  private deedGo(id:string): void {
    const e=this.POP[id]; if(!e||!e.tsCnum) return;
    const w=this.openDeferred();
    const nameFallback=()=>this.tsCountyThen(w,e.tsCnum,tsNameUrl(e.owner||''));
    if(e.localBP){ this.tsCountyThen(w,e.tsCnum,tsBookPageUrl(e.localBP)); return; }
    if(e.gislink && this.workerUrl){
      fetch(this.workerUrl+'?gislink='+encodeURIComponent(e.gislink)).then((r)=>r.json())
        .then((d:any)=>{ if(d&&d.ok&&d.best&&d.best.book){ this.setStatus('Latest deed: '+(d.best.type||'')+' Bk '+d.best.book+' Pg '+d.best.page); this.tsCountyThen(w,e.tsCnum,tsBookPageUrl(d.best)); } else { this.setStatus('No book/page found â using owner-name search'); nameFallback(); } })
        .catch(()=>nameFallback());
    } else nameFallback();
  }
  private deedName(id:string): void { const e=this.POP[id]; if(!e||!e.tsCnum) return; this.tsCountyThen(this.openDeferred(),e.tsCnum,tsNameUrl(e.owner||'')); }
  private deedGoUS(id:string): void {
    const e=this.POP[id]; window.open(US_BASE,'_blank');
    const show=(bp:any)=>{ this.setStatus('US Title Search Â· '+(e.county||'')+': Begin Search â Book/Page â Book '+bp.book+'  Page '+bp.page+(bp.type?'  ('+bp.type+')':'')); this.copyText(bp.book+' '+bp.page); };
    if(e.localBP){ show(e.localBP); return; }
    if(e.gislink && this.workerUrl){
      fetch(this.workerUrl+'?gislink='+encodeURIComponent(e.gislink)).then((r)=>r.json())
        .then((d:any)=>{ if(d&&d.ok&&d.best&&d.best.book) show(d.best); else this.setStatus('No book/page found â use name search in US Title Search'); })
        .catch(()=>this.setStatus('Deed lookup unavailable â use name search in US Title Search'));
    } else this.setStatus('Opened US Title Search â Begin Search.');
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
    this.setStatus('Searching '+src.label.replace(/^..? â /,'')+'â¦');
    const url=src.url+'?'+qs({where:where,outFields:outFieldsFor(src),returnGeometry:true,outSR:4326,resultRecordCount:60,f:'json'});
    this.arcgisFetch(url).then((d:any)=>{ if(d.error) throw new Error(d.error.message||'error'); const feats=esriToFeatures(d); feats.forEach((f:any)=>{ f.properties.__src=src.id; }); this.showResults(feats,src); })
      .catch((e:any)=>this.setStatus('Search failed: '+e.message));
  }

  private runAdvancedSearch(): void {
    const $ = (s:string)=>this.domElement.querySelector(s) as any;
    const src = SOURCES.filter((s)=>s.id==='tn')[0];
    const county = ($('#adv-county').value||'').trim();
    const mapRaw = ($('#adv-map').value||'').trim();
    const grpRaw = ($('#adv-group').value||'').trim();
    const parRaw = ($('#adv-parcel').value||'').trim();
    const subRaw = ($('#adv-subdiv').value||'').trim();
    const sortKey = ($('#adv-sort').value||'parcel');
    if(!county){ this.setStatus('Pick a county for Advanced search'); return; }
    if(!mapRaw && !parRaw && !subRaw){ this.setStatus('Enter a Control Map, Parcel, or Subdivision (county alone is too broad)'); return; }
    const sql=(x:string)=>String(x).replace(/'/g,"''");
    const map=wkNormMap(mapRaw); const p5=wkNorm5(parRaw);
    const grp=grpRaw?grpRaw.toUpperCase():'';
    const parts:string[]=["UPPER(COUNTY_NAME)='"+sql(county.toUpperCase())+"'"];
    if(map) parts.push("PARCELID LIKE '___ "+sql(map)+"%'");
    if(p5) parts.push("PARCELID LIKE '%"+sql(p5)+" %'");
    if(subRaw) parts.push("UPPER(SUBDIV) LIKE '%"+sql(subRaw.toUpperCase())+"%'");
    const where=parts.join(' AND ');
    const sortFld:any={parcel:'PARCELID',owner:'OWNER',address:'ADDRESS',acres:'DEEDAC DESC'};
    const orderBy=sortFld[sortKey]||'PARCELID';
    this.setStatus('Searching '+county+'â¦');
    const url=src.url+'?'+qs({where:where,outFields:outFieldsFor(src),orderByFields:orderBy,returnGeometry:true,outSR:4326,resultRecordCount:100,f:'json'});
    this.arcgisFetch(url).then((d:any)=>{
      if(d.error) throw new Error(d.error.message||'error');
      let feats=esriToFeatures(d);
      feats=feats.filter((f:any)=>{ const pid=String(f.properties.PARCELID||'');
        if(map && pid.substring(4,8).replace(/\s+$/,'')!==map) return false;
        if(p5 && pid.substring(11,16)!==p5) return false;
        if(grp && pid.substring(8,11).replace(/\s+$/,'')!==grp) return false;
        return true; });
      feats.forEach((f:any)=>{ f.properties.__src='tn'; });
      this.showResults(feats,src);
      if(feats.length===1){ this.gotoFeature(feats[0], normalize(feats[0].properties, src)); }
    }).catch((e:any)=>this.setStatus('Search failed: '+e.message));
  }

  private showResults(feats:any[], src:any): void {
    const box=this.domElement.querySelector('#results') as any; const list=this.domElement.querySelector('#rlist') as any;
    (this.domElement.querySelector('#rtitle') as any).textContent=feats.length+' result'+(feats.length===1?'':'s');
    list.innerHTML='';
    if(feats.length===0){ list.innerHTML='<div class="rrow"><span>No matches.</span></div>'; box.style.display='block'; this.setStatus('No matches'); return; }
    feats.forEach((f:any)=>{ const n=normalize(f.properties,src); const r=document.createElement('div'); r.className='rrow'; r.innerHTML='<b>'+esc(n.owner||n.address||n.pin||'(parcel)')+'</b><span>'+esc([n.address,n.pin].filter(Boolean).join(' Â· '))+'</span>'; r.onclick=()=>this.gotoFeature(f,n); list.appendChild(r); });
    box.style.display='block'; this.setStatus(feats.length+' result(s)');
  }

  private gotoFeature(f:any, n:any): void {
    this.hiLayer.clearLayers(); this.hiLayer.addData(f);
    try{ this.map.fitBounds(this.hiLayer.getBounds(),{maxZoom:18,padding:[40,40]}); }catch(e){}
    this.selectParcel(f, n);
    L.popup({maxWidth:320}).setLatLng(this.hiLayer.getBounds().getCenter()).setContent(this.popupHtml(n)).openOn(this.map);
  }

  // ======================= RBS zoning layer =======================
  private cfg(): any { return (SPHttpClient as any).configurations.v1; }
  private listApi(): string { return this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + this.zoneListTitle.replace(/'/g,"''") + "')"; }
  private workedApi(): string { return this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('DLS Worked Parcels')"; }
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
    }).catch(()=>{ /* list missing / no access â zoning just stays empty */ });
  }

  private featPin(ft:any): string { const src=SOURCES.filter((s)=>s.id===ft.properties.__src)[0]||SOURCES[0]; return pinKey(pick(ft.properties, src.f.pin)); }
  private parcelStyle(ft:any): any {
    if(this.zoningView){ const z=this.zoneByPin[this.featPin(ft)]; if(z && this.jurShow[z.jur]!==false){ if(z.split) return {color:'#6b5300',weight:1,fillColor:'#000',fillOpacity:0.001}; const j=jurById(z.jur)||ZJURS[0]; const c=(j.colors&&j.colors[z.zone])||'#888'; return {color:'#6b5300',weight:1,fillColor:c,fillOpacity:0.55}; } }
    return {color:'#ffd24d',weight:1,fillColor:'#000',fillOpacity:0.001};
  }
  private restyleParcels(): void { try{ if(this.parcelLayer) this.parcelLayer.setStyle((ft:any)=>this.parcelStyle(ft)); }catch(e){} }

  // ---- ONE combined legend: collapsible Zoning + Projects sections (in #zlegend) ----
  private buildLegend(): void {
    const el=this.domElement.querySelector('#zlegend') as any; if(!el) return;
    const showZ=this.zoningView, showP=this._projOn, showW=this.workView, showI=this._inqOn;
    if(!showZ && !showP && !showW && !showI){ el.style.display='none'; el.innerHTML=''; return; }
    el.style.display='block';
    let h='';
    if(showZ) h+='<div class="lp-sec'+(this._collZ?' coll':'')+'"><div class="lp-hd" data-sec="Z"><span class="tw">&#9662;</span> Zoning</div><div class="lp-bd" id="zoneHost"></div></div>';
    if(showP) h+='<div class="lp-sec'+(this._collP?' coll':'')+'"><div class="lp-hd" data-sec="P"><span class="tw">&#9662;</span> Projects <span id="pCount" class="pp-ct"></span></div><div class="lp-bd" id="projHost"></div></div>';
    if(showW) h+='<div class="lp-sec'+(this._collW?' coll':'')+'"><div class="lp-hd" data-sec="W"><span class="tw">&#9662;</span> Work history <span id="wCount" class="pp-ct"></span></div><div class="lp-bd" id="workHost"></div></div>';
    if(showI) h+='<div class="lp-sec'+(this._collI?' coll':'')+'"><div class="lp-hd" data-sec="I"><span class="tw">&#9662;</span> Inquiries <span id="iCount" class="pp-ct"></span></div><div class="lp-bd" id="inqHost"></div></div>';
    el.innerHTML=h;
    const self=this;
    const hds=el.querySelectorAll('.lp-hd'); for(let i=0;i<hds.length;i++){ hds[i].addEventListener('click',function(){ const s=this.getAttribute('data-sec'); if(s==='Z') self._collZ=!self._collZ; else if(s==='P') self._collP=!self._collP; else if(s==='W') self._collW=!self._collW; else if(s==='I') self._collI=!self._collI; if(this.parentNode) this.parentNode.classList.toggle('coll'); }); }
    if(showZ) this.buildZPanel();
    if(showP){ this.buildProjectPanel(); this.renderProjectPins(); }
    if(showW){ this.buildWorkPanel(); }
    if(showI){ this.buildInquiryPanel(); this.renderInquiryPins(); }
  }

  private buildZPanel(): void {
    const el=this.domElement.querySelector('#zoneHost') as any; if(!el) return;
    const self=this;
    let h='<div class="zsh">Show: ';
    ZJURS.forEach((j:any)=>{ if(!j.taggable) return; h+='<label class="zshl"><input type="checkbox" data-jsh="'+j.id+'"'+(this.jurShow[j.id]!==false?' checked':'')+'> '+esc(j.name)+'</label>'; });
    h+='</div>';
    h+='<div class="lp-sub'+(this._collUcdd?' coll':'')+'" data-ucddsub="1"><div class="lp-subhd"><span class="tw">&#9662;</span> Official zoning (live)'+(this._ucddCount?' <span class="pp-ct">'+this._ucddCount+' lots</span>':'')+'</div><div class="lp-subbd">';
    h+='<div class="zdisc" style="margin:2px 0 4px">Auto-loads official county/city zoning for whatever is on screen &mdash; scroll the map &amp; the district legend below follows your view.</div><div class="zsh">';
    h+='</div><div class="zdisc">Official county/city zoning, lot-by-lot. Zoom in (&#8805; '+UCDD_MINZOOM+').</div></div></div>';
    h+='<div class="ztag">Tag lots as: <select id="ztagjur"><option value="auto">Auto-detect</option>';
    ZJURS.forEach((j:any)=>{ if(j.taggable) h+='<option value="'+j.id+'">'+esc(j.name)+'</option>'; });
    h+='</select></div>';
    h+='<div class="lp-sub'+(this._collLegend?' coll':'')+'" data-leg="1"><div class="lp-subhd"><span class="tw">&#9662;</span> District legend</div><div class="lp-subbd">';
    ZJURS.forEach((j:any)=>{ if(!j.taggable) return; h+='<div class="zjh">'+esc(j.name)+'</div>'; j.zones.forEach((z:string)=>{ h+='<div class="zi"><span class="zsw" style="background:'+j.colors[z]+'"></span>'+z+' &middot; '+esc(j.names[z])+'</div>'; }); });
for(var k=0;k<UCDD_ZONING.length;k++){ var uu=UCDD_ZONING[k]; var ufe=this._ucddCache&&this._ucddCache[uu.key]; if(!ufe||!ufe.length) continue; h+='<div class="zjh">'+esc(uu.label)+'</div>'; var seen:any={}; for(var fi=0;fi<ufe.length;fi++){ var vv=ufe[fi].properties[uu.field]; if(vv==null) continue; var vs=(''+vv).replace(/^\s+|\s+$/g,''); if(!vs||seen[vs.toUpperCase()]) continue; seen[vs.toUpperCase()]=1; var nm=ucddName(uu.key,vs); h+='<div class="zi"><span class="zsw" style="background:'+ucddColor(uu.key,vs)+'"></span>'+esc(vs)+(nm?' &middot; '+esc(nm):'')+'</div>'; } }
    h+='</div></div>';
    h+='<div class="zdiv"></div><div class="zjh">Other layers</div>';
    h+='<div class="zrow"><label><input type="checkbox" id="zfema"'+(this._femaOn?' checked':'')+'> FEMA flood (NFHL)</label><span class="zacc exact">live</span><input type="range" min="20" max="100" value="'+(this.femaLayer&&this.femaLayer.options?Math.round(this.femaLayer.options.opacity*100):55)+'" id="zfemaop"></div>';
    h+='<div class="zrow"><label><input type="checkbox" id="zareas"'+(this._areasOn?' checked':'')+'> Drawn areas (historic dist.)</label></div>';
    h+='<button class="zbtn2" id="zdrawarea" style="margin-top:4px">Draw an area&hellip;</button>';
    h+='<div class="zdisc">Local tags are colored by district; UCDD layers are official &amp; live. Use "Show" to toggle each entity. Reference only &mdash; confirm zoning with the city/county.</div>';
    el.innerHTML=h;
    const jsh=el.querySelectorAll('[data-jsh]'); for(let i2=0;i2<jsh.length;i2++){ jsh[i2].addEventListener('change',function(){ self.jurShow[this.getAttribute('data-jsh')]=this.checked; self.restyleParcels(); self.buildSplitLayer(); }); }
    const ush=el.querySelectorAll('[data-ush]'); for(let i3=0;i3<ush.length;i3++){ ush[i3].addEventListener('change',function(){ self.jurShow[this.getAttribute('data-ush')]=this.checked; self._ucddBounds=null; self.loadUcdd(); }); }
    const ts=el.querySelector('#ztagjur') as any; if(ts){ ts.value=this.tagJur; ts.addEventListener('change',function(e:any){ self.tagJur=e.target.value; }); }
    const uh=el.querySelector('[data-ucddsub] .lp-subhd') as any; if(uh) uh.addEventListener('click',function(){ self._collUcdd=!self._collUcdd; if(this.parentNode) this.parentNode.classList.toggle('coll'); });
    const lh=el.querySelector('[data-leg] .lp-subhd') as any; if(lh) lh.addEventListener('click',function(){ self._collLegend=!self._collLegend; if(this.parentNode) this.parentNode.classList.toggle('coll'); });
    const fm=el.querySelector('#zfema') as any; if(fm) fm.addEventListener('change',function(e:any){ self._femaOn=!!e.target.checked; self.applyFema(); });
    const fo=el.querySelector('#zfemaop') as any; if(fo) fo.addEventListener('input',function(e:any){ if(self.femaLayer) self.femaLayer.setOpacity((+e.target.value)/100); });
    const ar=el.querySelector('#zareas') as any; if(ar) ar.addEventListener('change',function(e:any){ self._areasOn=!!e.target.checked; self.buildAreasLayer(); });
    const da=el.querySelector('#zdrawarea') as any; if(da) da.addEventListener('click',function(){ self.startAreaDraw(); });
  }

  private loadUcdd(): void {
    if(!this.ucddLayer || !this._ucddRenderer) return;
    var self=this;
    if(!this.zoningView){ if(this.map.hasLayer(this.ucddLayer)){ this.ucddLayer.clearLayers(); this.map.removeLayer(this.ucddLayer); } this._ucddBounds=null; this._ucddCount=0; this._ucddCache={}; return; }
    var b=this.map.getBounds(); var vw=[b.getWest(),b.getSouth(),b.getEast(),b.getNorth()];
    var z=this.map.getZoom();
    var on:any[]=[];
    for(var i=0;i<UCDD_ZONING.length;i++){ var cc=UCDD_ZONING[i]; if(ucddHit(cc.bbox,vw)) on.push(cc); }
    if(on.length===0 || z<UCDD_MINZOOM){ this.ucddLayer.clearLayers(); this._ucddCount=0; this._ucddBounds=null; this._ucddCache={}; if(this.map.hasLayer(this.ucddLayer)) this.map.removeLayer(this.ucddLayer); this.buildLegend(); return; }
    if(!this.map.hasLayer(this.ucddLayer)) this.ucddLayer.addTo(this.map);
    if(this._ucddBounds && this._ucddZoom===z && this._ucddBounds.contains(b)) return;
    var pb=b.pad(0.4); this._ucddBounds=pb; this._ucddZoom=z;
    var env=[pb.getWest(),pb.getSouth(),pb.getEast(),pb.getNorth()].join(',');
    var res=(b.getEast()-b.getWest())/Math.max(1,this.map.getSize().x); var off=res>0?res:0.00001;
    var seq= ++this._ucddSeq; var cache:any={}; var tmp:any[]=[]; var tot=0; var pending=on.length; var commit=function(){ if(self._ucddSeq!==seq) return; self._ucddCache=cache; self.ucddLayer.clearLayers(); for(var ti=0;ti<tmp.length;ti++){ self.ucddLayer.addLayer(tmp[ti]); } self._ucddCount=tot; self.buildLegend(); };
    for(var k2=0;k2<on.length;k2++){
      (function(c:any){
        var url=(c.url?c.url:(UCDD_BASE+'/'+c.service+'/FeatureServer/'+c.layer))+'/query?'+qs({where:'1=1',geometry:env,geometryType:'esriGeometryEnvelope',inSR:4326,spatialRel:'esriSpatialRelIntersects',outFields:c.field,returnGeometry:true,outSR:4326,maxAllowableOffset:off,geometryPrecision:6,resultRecordCount:2000,f:'json'});
        self.arcgisFetch(url).then(function(d:any){ if(self._ucddSeq===seq && d && !d.error){ var feats=esriToFeatures(d); cache[c.key]=feats; var lyr=L.geoJSON(feats,{pane:'ucdd',renderer:self._ucddRenderer,style:function(ft:any){ return {color:'#444',weight:0.4,fillColor:ucddColor(c.key,ft.properties[c.field]),fillOpacity:0.55}; }}); tmp.push(lyr); tot+=feats.length; } pending--; if(pending<=0) commit(); }).catch(function(){ pending--; if(pending<=0) commit(); });
      })(on[k2]);
    }
  }

  private setZoningMode(v:string): void {
    this.zoningView = (v==='view'||v==='edit');
    this.zoningEdit = (v==='edit');
    if(this.zoningEdit && this.workEdit){ this.workEdit=false; const _ws=this.domElement.querySelector('#wmode') as any; if(_ws) _ws.value='view'; this.workView=true; }
    this.buildLegend();
    this.applyOverlays();
    this.restyleParcels();
    this.buildSplitLayer();
    this.applyFema();
    this.buildAreasLayer();
    this.loadUcdd();
    if(v==='edit') this.setStatus('Zoning EDIT â set "Tag lots as" if needed, click a lot, choose its zone. Reference only.');
    else if(v==='view') this.setStatus('Zoning VIEW â tagged lots are colored by their district.');
    else this.setStatus('Zoning off.');
    this.map.closePopup();
  }

  private applyOverlays(): void {
    ZJURS.forEach((j:any)=>{ if(!j._layer) return; const show=this.zoningView && j._on; if(show){ if(!this.map.hasLayer(j._layer)){ j._layer.addTo(this.map); j._layer.bringToFront(); } } else if(this.map.hasLayer(j._layer)){ this.map.removeLayer(j._layer); } });
  }

  private openZonePicker(n:any, ll:any, feat?:any): void {
    const pin=pinKey(n.pin); if(!pin){ this.setStatus('This parcel has no ID â cannot tag it.'); return; }
    const j = (this.tagJur && this.tagJur!=='auto') ? jurById(this.tagJur) : (jurAt(ll) || nearestJur(ll));
    if(!j){ this.setStatus('No zoning jurisdiction available to tag.'); return; }
    this.zTarget={pin:pin, raw:String(n.pin).trim(), jur:j.id, ll:ll, ring:outerRing(feat&&feat.geometry)};
    this.renderZonePicker();
  }

  // Re-pick the jurisdiction for the lot being tagged (RBS / Lafayette / Macon) â lets a city lot that the
  // bounding box auto-detected as "Macon County" be tagged with the correct city's districts instead.
  private switchZoneJur(jurId:string): void {
    const j=jurById(jurId); if(!j||!this.zTarget) return;
    this.zTarget.jur=jurId; this.renderZonePicker();
    this.setStatus('Jurisdiction set to '+j.name+' â now pick the district.');
  }

  private renderZonePicker(): void {
    const t=this.zTarget; if(!t||!t.jur) return; const j=jurById(t.jur); if(!j) return;
    const cur=this.zoneByPin[t.pin];
    const curTxt = cur? (cur.split&&cur.pieces? esc(cur.pieces.map((p:any)=>p.z||'blank').join(' / '))+' (split, '+esc(cur.jur)+')' : esc(cur.zone)+' ('+esc(cur.jur)+')'+(cur.flood?' + Floodplain':'')) : 'â';
    let jb=''; ZJURS.forEach((jj:any)=>{ if(!jj.taggable) return; jb+='<button class="zp-jurbtn'+(jj.id===t.jur?' on':'')+'" data-act="zjur" data-arg="'+esc(jj.id)+'">'+esc(jj.name)+'</button>'; });
    let g=''; j.zones.forEach((z:string)=>{ g+='<button class="zbtn" data-act="zset" data-arg="'+z+'" style="background:'+j.colors[z]+'">'+z+'<small>'+esc(j.names[z])+'</small></button>'; });
    const html='<div class="zp"><div class="zp-h">Set zoning &middot; '+esc(j.name)+'</div>'
      +'<div class="zp-pin">Parcel: <b>'+esc(t.raw)+'</b></div>'
      +'<div class="zp-cur">Current: <b>'+curTxt+'</b></div>'
      +'<div class="zp-jurs"><span>Jurisdiction:</span> '+jb+'</div>'
      +'<div class="zp-grid">'+g+'</div>'
      +'<label class="zp-fl"><input type="checkbox" id="zpFlood"'+(cur&&cur.flood?' checked':'')+'> In 1% floodplain</label>'
      +'<div style="margin-top:4px"><button class="zp-clear" data-act="zsplitopen">Split lot&hellip;</button> <button class="zp-clear" data-act="zclear">Clear</button></div>'
      +'<div class="zp-note">If this lot is inside a city, pick the right jurisdiction above. '+esc(j.name)+(j.accuracy==='approx'?' overlay is approximate â verify boundary lots. ':' ')+'Reference only â not an official determination.</div></div>';
    L.popup({maxWidth:300,autoPanPadding:[24,24]}).setLatLng(t.ll||this.map.getCenter()).setContent(html).openOn(this.map);
  }

  private saveZone(zone:string): void {
    const t=this.zTarget; if(!t||!t.pin||!t.jur) return;
    const j=jurById(t.jur); if(!j||j.zones.indexOf(zone)<0) return;
    const fl=this.domElement.querySelector('#zpFlood') as any; const flood=!!(fl&&fl.checked);
    const cur=this.zoneByPin[t.pin];
    const done=(id:number)=>{ this.zoneByPin[t.pin]={zone:zone,flood:flood,id:id,jur:t.jur}; this.restyleParcels(); this.setStatus('Saved '+t.raw+' â '+zone+' ('+t.jur+')'+(flood?' + floodplain':'')); this.map.closePopup(); };
    if(cur && cur.id){
      this.spPost(this.listApi()+'/items('+cur.id+')', {Zone:zone,Floodplain:flood,Jurisdiction:t.jur}, {'X-HTTP-Method':'MERGE','IF-MATCH':'*'})
        .then((r:any)=>{ if(r.status>=200&&r.status<300) done(cur.id); else this.setStatus('Save failed ('+r.status+') â check list permissions'); })
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
    if(!st.rings || st.rings.length<2){ this.setStatus('That line did not divide the lot â try again.'); this.cancelSplit(); return; }
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
    Object.keys(m).forEach((pin)=>{ const z=m[pin]; if(z&&z.split&&z.pieces&&self.jurShow[z.jur]!==false){ z.pieces.forEach((p:any)=>{ if(!p.z||!p.r||p.r.length<3) return; const latlngs=p.r.map((c:any)=>[c[1],c[0]]); self.splitLayer.addLayer(L.polygon(latlngs,{color:'#6b5300',weight:1,fillColor:self.jurColor(z.jur,p.z),fillOpacity:0.55,interactive:false,pane:'zsplit'})); }); } });
  }

  // ======================= Projects (Coverage Map) layer =======================
  private loadProjects(): void {
    if(this._projLoaded){ this.buildLegend(); return; }
    const url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists(guid'" + this.projectsListGuid + "')/items?$top=500&$select=Id,Title,JobLabel,Lat,Lng,FolderURL,County,ProjectStatus,JobTypeStandard,Property_x0020_Address,FNLTDate";
    this.setStatus('Loading projectsâ¦');
    this.spGet(url).then((d:any)=>{
      const items=(d&&d.value)||[]; const arr:any[]=[];
      items.forEach((x:any)=>{ if(x.Lat==null||x.Lng==null) return; const t=(x.Title==null?'':String(x.Title)); const yr=/^\d{6}/.test(t)?('20'+t.substring(0,2)):'(blank)'; const county=(x.County==null?'':String(x.County)).replace(/^\d+\s*-\s*/,'').replace(/,?\s*(TN|KY)\b.*$/i,'').trim(); arr.push({ Id:x.Id, Title:t, JobLabel:x.JobLabel||'', Lat:x.Lat, Lng:x.Lng, FolderURL:x.FolderURL||'', County:county, Status:x.ProjectStatus||'', JobType:x.JobTypeStandard||'', Address:x.Property_x0020_Address||'', Year:yr, Deadline:deadlineBucket(x.ProjectStatus||'', x.FNLTDate) }); });
      this.projects=arr; this._projLoaded=true;
      arr.forEach((j:any)=>{ const s=j.Status||'(no status)'; if(this.pStatusOn[s]===undefined) this.pStatusOn[s]=DEFAULT_STATUS_ON.indexOf(s)>=0; const c=j.County||'(blank)'; if(this.pCountyOn[c]===undefined) this.pCountyOn[c]=true; const ty=j.JobType||'(blank)'; if(this.pTypeOn[ty]===undefined) this.pTypeOn[ty]=true; const y=j.Year||'(blank)'; if(this.pYearOn[y]===undefined) this.pYearOn[y]=true; const dl=j.Deadline||'No date'; if(this.pDeadlineOn[dl]===undefined) this.pDeadlineOn[dl]=true; });
      this.buildLegend(); this.setStatus(this.projects.length+' projects loaded');
    }).catch((e:any)=>{ this.setStatus('Projects: could not load WIP jobs ('+e+')'); });
  }

  private projVisible(j:any): boolean {
    if(this.pStatusOn[j.Status||'(no status)']===false) return false;
    if(this.pCountyOn[j.County||'(blank)']===false) return false;
    if(this.pTypeOn[j.JobType||'(blank)']===false) return false;
    if(this.pYearOn[j.Year||'(blank)']===false) return false;
    if(this.pDeadlineOn[j.Deadline||'No date']===false) return false;
    if(this.pSearch){ const hay=(j.JobLabel+' '+j.Title+' '+j.Address+' '+j.County).toLowerCase(); if(hay.indexOf(this.pSearch)<0) return false; }
    return true;
  }

  private renderProjectPins(): void {
    if(!this.projectLayer) return; this.projectLayer.clearLayers();
    if(!this._projOn){ this.updateProjCount(0); return; }
    let shown=0; const self=this;
    this.projects.forEach((j:any)=>{
      if(!self.projVisible(j)) return; shown++;
      const st=j.Status||'(no status)';
      const label=j.JobLabel||(j.Title+(j.County?' â '+j.County:''));
      const folder=j.FolderURL?'<a class="dls-pop-a" href="'+esc(j.FolderURL)+'" target="_blank" rel="noopener">Open project folder &#8599;</a>':'';
      const html='<div class="dls-pop"><b>'+esc(label)+'</b>'+(j.Status?'<div class="m">'+esc(j.Status)+(j.JobType?' &middot; '+esc(j.JobType):'')+'</div>':'')+(j.Address?'<div class="m">'+esc(j.Address)+'</div>':'')+folder+'</div>';
      const m=L.circleMarker([j.Lat,j.Lng],{radius:7,fillColor:colorFor(st),color:'#fff',weight:2,fillOpacity:0.9,pane:'projects',renderer:self._projRenderer});
      m.bindPopup(html,{maxWidth:260}); m.addTo(self.projectLayer);
    });
    this.updateProjCount(shown);
  }
  private updateProjCount(shown:number): void { const c=this.domElement.querySelector('#pCount'); if(c) c.textContent=shown+' of '+this.projects.length+' jobs'; }

  private setProjectsMode(on:boolean): void {
    this._projOn=on;
    const sel=this.domElement.querySelector('#proj') as any; if(sel && sel.value!==(on?'on':'off')) sel.value=on?'on':'off';
    if(on){ this.loadProjects(); }
    else { if(this.projectLayer) this.projectLayer.clearLayers(); this.buildLegend(); this.setStatus('Projects layer off.'); }
  }

  private projCounts(keyFn:any): any { const m:any={}; this.projects.forEach((j:any)=>{ const k=keyFn(j); m[k]=(m[k]||0)+1; }); return m; }

  private buildProjectPanel(): void {
    const el=this.domElement.querySelector('#projHost') as any; if(!el) return;
    let h='<input id="pSearch" class="pp-search" placeholder="Search client / address / job #" value="'+esc(this.pSearch)+'"/>';
    h+='<button class="pp-reset" id="pReset">Reset filters</button>';
    const dims=[['status','Project status'],['deadline','Deadline'],['county','County'],['type','Job type'],['year','Year']];
    for(let i=0;i<dims.length;i++){ const dim=dims[i][0], lbl=dims[i][1]; const cap=dim.charAt(0).toUpperCase()+dim.slice(1); h+='<div class="lp-sub'+(this._dimColl[dim]?' coll':'')+'" data-dimsec="'+dim+'"><div class="lp-subhd"><span class="tw">&#9662;</span> '+lbl+' <span class="pp-all" data-pdim="'+dim+'">all</span></div><div class="lp-subbd"><div id="p'+cap+'"></div></div></div>'; }
    el.innerHTML=h;
    const self=this;
    const si=el.querySelector('#pSearch') as any; if(si) si.oninput=function(){ self.pSearch=(si.value||'').toLowerCase().trim(); self.renderProjectPins(); };
    const rb=el.querySelector('#pReset') as any; if(rb) rb.onclick=function(){ self.projReset(); };
    const subhds=el.querySelectorAll('.lp-subhd'); for(let i=0;i<subhds.length;i++){ subhds[i].addEventListener('click',function(e:any){ const a=(e.target&&e.target.closest)?e.target.closest('[data-pdim]'):null; if(a){ self.projAllOn(a.getAttribute('data-pdim')); return; } const sub=this.parentNode; const dim=sub.getAttribute('data-dimsec'); self._dimColl[dim]=!self._dimColl[dim]; sub.classList.toggle('coll'); }); }
    this.renderProjSections();
  }

  private renderProjSections(): void {
    this.renderProjSection('#pStatus', this.pStatusOn, 'status', true);
    this.renderProjSection('#pDeadline', this.pDeadlineOn, 'deadline', false);
    this.renderProjSection('#pCounty', this.pCountyOn, 'county', false);
    this.renderProjSection('#pType', this.pTypeOn, 'type', false);
    this.renderProjSection('#pYear', this.pYearOn, 'year', false);
  }

  private renderProjSection(sel:string, state:any, dim:string, color:boolean): void {
    const el=this.domElement.querySelector(sel) as any; if(!el) return;
    const keyFn = dim==='status'?(j:any)=>j.Status||'(no status)' : dim==='county'?(j:any)=>j.County||'(blank)' : dim==='type'?(j:any)=>j.JobType||'(blank)' : dim==='year'?(j:any)=>j.Year||'(blank)' : (j:any)=>j.Deadline||'No date';
    const counts=this.projCounts(keyFn);
    const keys=Object.keys(counts).sort((a:any,b:any)=>{ if(dim==='status'){ const ia=STATUS_ORDER.indexOf(a),ib=STATUS_ORDER.indexOf(b); return (ia<0?99:ia)-(ib<0?99:ib); } if(dim==='deadline'){ const ia=DEADLINE_ORDER.indexOf(a),ib=DEADLINE_ORDER.indexOf(b); return (ia<0?99:ia)-(ib<0?99:ib); } if(dim==='year') return b.localeCompare(a); return a.localeCompare(b); });
    const self=this; el.innerHTML='';
    keys.forEach((k:string)=>{ const row=document.createElement('div'); row.className='pp-row'+(state[k]===false?' off':''); const mark=color?'<span class="pp-dot" style="background:'+colorFor(k)+'"></span>':'<span class="pp-sq"></span>'; row.innerHTML=mark+'<span class="pp-nm">'+esc(k||'(blank)')+'</span><span class="pp-n">'+counts[k]+'</span>'; row.onclick=function(){ state[k]=state[k]===false; self.renderProjectPins(); self.renderProjSections(); }; el.appendChild(row); });
  }

  private projAllOn(dim:string): void { const st = dim==='status'?this.pStatusOn:dim==='county'?this.pCountyOn:dim==='type'?this.pTypeOn:dim==='year'?this.pYearOn:dim==='deadline'?this.pDeadlineOn:null; if(!st) return; Object.keys(st).forEach((k)=>{ st[k]=true; }); this.renderProjectPins(); this.renderProjSections(); }

  private projReset(): void { const self=this; Object.keys(this.pStatusOn).forEach((k)=>{ self.pStatusOn[k]=DEFAULT_STATUS_ON.indexOf(k)>=0; }); [this.pCountyOn,this.pTypeOn,this.pYearOn,this.pDeadlineOn].forEach((st:any)=>{ Object.keys(st).forEach((k)=>{ st[k]=true; }); }); this.pSearch=''; const si=this.domElement.querySelector('#pSearch') as any; if(si) si.value=''; this.renderProjectPins(); this.renderProjSections(); }

  // ======================= Inquiries (IQ list) layer =======================
  private inquiriesApi(): string { return this.context.pageContext.web.absoluteUrl + "/_api/web/lists(guid'" + this.inquiriesListGuid + "')/items"; }

  private setInquiriesMode(on:boolean): void {
    this._inqOn=on;
    const sel=this.domElement.querySelector('#inq') as any; if(sel && sel.value!==(on?'on':'off')) sel.value=on?'on':'off';
    if(on){ this.loadInquiries(); }
    else { if(this.inqLayer) this.inqLayer.clearLayers(); this.buildLegend(); this.setStatus('Inquiries layer off.'); }
  }

  private iqStatus(v:any): string { if(v==null) return ''; if(typeof v==='object') return String(v.Value||v.value||''); return String(v); }
  private iqCountyName(v:any): string { let c:any=v; if(c&&typeof c==='object') c=c.Value||c.value||''; c=(c==null?'':String(c)); return c.replace(/^\d+\s*-\s*/,'').replace(/,?\s*(TN|KY)\b.*$/i,'').replace(/ COUNTY$/i,'').trim(); }
  private iqMulti(v:any): string { if(v==null) return ''; if(Object.prototype.toString.call(v)==='[object Array]') return v.join(', '); if(typeof v==='object'&&v.results) return v.results.join(', '); return String(v); }
  private iqUrl(v:any): string { if(!v) return ''; if(typeof v==='object') return String(v.Url||v.url||''); return String(v); }
  private iqIsTn(name:string): boolean { const u=(name||'').toUpperCase(); for(let i=0;i<TN_COUNTIES.length;i++){ if(TN_COUNTIES[i].toUpperCase()===u) return true; } return false; }

  private loadInquiries(): void {
    if(this._inqLoaded){ this.buildLegend(); return; }
    const cols='Id,Title,QuoteNumber,QuoteStatus,JobNumber,County,TaxMap,ParcelNumber,EstimatedAcreage,QuotedAmount,SurveyProjectType,PropertyOwnerName,PropertyOwnerLastName,PrimaryStreet,FollowUpDate,EstimatorFileUrl,EstimatorFileName,ProjectContractUrl,ProjectContractName';
    const url=this.inquiriesApi()+'?$top=2000&$select='+cols;
    this.setStatus('Loading inquiriesâ¦'); const self=this;
    this.spGet(url).then((d:any)=>{
      const items=(d&&d.value)||[]; const arr:any[]=[];
      for(let i=0;i<items.length;i++){ const x=items[i];
        const st=self.iqStatus(x.QuoteStatus); if(st===IQ_EXCLUDE_STATUS) continue;
        const qn=(x.QuoteNumber==null?'':String(x.QuoteNumber)); const ym=qn.match(/^Q(\d{2})/); const yr=ym?('20'+ym[1]):'(blank)';
        const owner=((x.PropertyOwnerName||'')+' '+(x.PropertyOwnerLastName||'')).replace(/\s+/g,' ').trim();
        arr.push({ Id:x.Id, client:(x.Title==null?'':String(x.Title)), quote:qn, status:st||'(blank)', county:self.iqCountyName(x.County),
          taxMap:(x.TaxMap==null?'':String(x.TaxMap)), parcel:(x.ParcelNumber==null?'':String(x.ParcelNumber)),
          acres:x.EstimatedAcreage, amount:x.QuotedAmount, ptype:self.iqMulti(x.SurveyProjectType),
          owner:owner, street:(x.PrimaryStreet==null?'':String(x.PrimaryStreet)), follow:x.FollowUpDate||'',
          estUrl:self.iqUrl(x.EstimatorFileUrl), conUrl:self.iqUrl(x.ProjectContractUrl), year:yr });
      }
      self.inquiries=arr; self._inqLoaded=true;
      for(let j=0;j<arr.length;j++){ const q=arr[j];
        const s=q.status||'(blank)'; if(self.iqStatusOn[s]===undefined) self.iqStatusOn[s]=IQ_DEFAULT_STATUS_ON.indexOf(s)>=0;
        const c=q.county||'(blank)'; if(self.iqCountyOn[c]===undefined) self.iqCountyOn[c]=true;
        const y=q.year||'(blank)'; if(self.iqYearOn[y]===undefined) self.iqYearOn[y]=true;
      }
      self.setStatus('Locating '+arr.length+' inquiries on their parcelsâ¦'); self.resolveInquiryParcels();
    }).catch((e:any)=>{ this.setStatus('Inquiries: could not load IQ list ('+e+')'); });
  }

  // Resolve each inquiry's County+TaxMap+Parcel to its exact lot via the TN statewide service (free, like Work history).
  private resolveInquiryParcels(): void {
    const self=this; const byCounty:any={};
    for(let i=0;i<this.inquiries.length;i++){ const q=this.inquiries[i];
      if(this.inqGeo[q.Id]) continue;
      const map=wkNormMap(q.taxMap); const p5=wkNorm5(q.parcel);
      if(!q.county || !map || !p5) continue;
      if(!self.iqIsTn(q.county)) continue;   // statewide service covers TN 86 counties (metros/KY skipped)
      const key=q.county.toUpperCase(); if(!byCounty[key]) byCounty[key]={county:q.county,items:[]}; byCounty[key].items.push({q:q,map:map,p5:p5});
    }
    // Chunk each county's inquiries so the OR'd PARCELID query URL stays under the ArcGIS GET limit
    // (Macon reached 38 inquiries -> a single query exceeded the URL limit, server returned HTML, the whole county failed to resolve).
    const groups:any[]=[]; const keys=Object.keys(byCounty); const CH=15;
    for(let ci=0;ci<keys.length;ci++){ const g=byCounty[keys[ci]]; for(let s=0;s<g.items.length;s+=CH){ groups.push({county:g.county,items:g.items.slice(s,s+CH)}); } }
    let pending=groups.length;
    if(pending===0){ self.finishInquiryResolve(); return; }
    for(let k=0;k<groups.length;k++){ this.resolveCountyParcels(groups[k], function(){ pending--; if(pending<=0) self.finishInquiryResolve(); }); }
  }

  private resolveCountyParcels(grp:any, done:any): void {
    const self=this; const sql=(x:string)=>String(x).replace(/'/g,"''");
    const ors:string[]=[]; for(let i=0;i<grp.items.length;i++){ const it=grp.items[i]; ors.push("(PARCELID LIKE '___ "+sql(it.map)+"%' AND PARCELID LIKE '%"+sql(it.p5)+" %')"); }
    const where="UPPER(COUNTY_NAME)='"+sql(grp.county.toUpperCase())+"' AND ("+ors.join(' OR ')+")";
    const url=WK_TN_SVC+'?'+qs({where:where,outFields:outFieldsFor(SOURCES[0]),returnGeometry:true,outSR:4326,resultRecordCount:1000,f:'json'});
    this.arcgisFetch(url).then((d:any)=>{ const feats=esriToFeatures(d);
      for(let g=0;g<grp.items.length;g++){ const it=grp.items[g]; if(self.inqGeo[it.q.Id]) continue;
        for(let f=0;f<feats.length;f++){ const pid=String((feats[f].properties as any).PARCELID||'');
          if(pid.substring(4,8).replace(/\s+$/,'')!==it.map) continue;
          if(pid.substring(11,16)!==it.p5) continue;
          (feats[f].properties as any).__src='tn'; const r=outerRing(feats[f].geometry); if(!r||!r.length) continue;
          self.inqGeo[it.q.Id]={feat:feats[f], center:centroid(r)}; break;
        }
      }
      done();
    }).catch(()=>{ done(); });
  }

  private finishInquiryResolve(): void {
    let loc=0; for(let i=0;i<this.inquiries.length;i++){ if(this.inqGeo[this.inquiries[i].Id]) loc++; }
    this._iqLocated=loc; this._iqUnplaced=this.inquiries.length-loc;
    this.buildLegend(); this.setStatus(this.inquiries.length+' inquiries ('+loc+' located on parcels).');
  }

  private iqVisible(q:any): boolean {
    if(this.iqStatusOn[q.status||'(blank)']===false) return false;
    if(this.iqCountyOn[q.county||'(blank)']===false) return false;
    if(this.iqYearOn[q.year||'(blank)']===false) return false;
    if(this.iqSearch){ const hay=(q.client+' '+q.owner+' '+q.street+' '+q.quote+' '+q.county).toLowerCase(); if(hay.indexOf(this.iqSearch)<0) return false; }
    return true;
  }

  private renderInquiryPins(): void {
    if(!this.inqLayer) return; this.inqLayer.clearLayers();
    if(!this._inqOn){ this.updateInqCount(0); return; }
    let shown=0; const self=this;
    const icon=L.divIcon({className:'dls-inq',html:IQ_TRI_SVG,iconSize:[20,18],iconAnchor:[10,9]});
    for(let i=0;i<this.inquiries.length;i++){ const q=this.inquiries[i];
      if(!self.iqVisible(q)) continue; const g=self.inqGeo[q.Id]; if(!g) continue; shown++;
      const m=L.marker([g.center[1],g.center[0]],{icon:icon,pane:'inquiries',title:q.client||q.quote});
      m.bindPopup(self.inquiryPopupHtml(q),{maxWidth:300});
      m.on('click',function(){ try{ self.selLayer.clearLayers(); self.selLayer.addData(g.feat); self.selFeat=g.feat; self.selN=normalize((g.feat.properties as any),SOURCES[0]); }catch(e){} });
      m.addTo(self.inqLayer);
    }
    this.updateInqCount(shown);
  }
  private updateInqCount(shown:number): void { const c=this.domElement.querySelector('#iCount'); if(c) c.textContent=shown+' shown Â· '+this._iqLocated+' located Â· '+this._iqUnplaced+' not located'; }

  private inquiryPopupHtml(q:any): string {
    let rows='';
    const row=(k:string,v:any)=>{ if(v||v===0) rows+='<tr><td class="k">'+k+'</td><td>'+esc(v)+'</td></tr>'; };
    row('Quote #', q.quote); row('Status', q.status);
    if(q.owner) row('Owner', q.owner);
    if(q.street) row('Address', q.street);
    if(q.county) row('County', q.county);
    if(q.acres||q.acres===0) row('Est. acreage', q.acres);
    if(q.amount||q.amount===0) row('Quoted', '$'+(+q.amount).toLocaleString());
    if(q.ptype) row('Type', q.ptype);
    if(q.follow){ const fd=new Date(q.follow); if(!isNaN(fd.getTime())) row('Follow-up', (fd.getMonth()+1)+'/'+fd.getDate()+'/'+fd.getFullYear()); }
    let links='';
    if(q.quote) links+='<a class="dls-pop-a" href="#" data-act="qfolder" data-arg="'+esc(q.quote)+'">Open quote folder &#8599;</a>';
    if(q.estUrl) links+='<a class="dls-pop-a" href="'+esc(q.estUrl)+'" target="_blank" rel="noopener">Estimator file &#8599;</a>';
    if(q.conUrl) links+='<a class="dls-pop-a" href="'+esc(q.conUrl)+'" target="_blank" rel="noopener">Contract &#8599;</a>';
    const head=esc(q.client||'(inquiry)')+(q.county?' â '+esc(q.county):'');
    return '<div class="dls-pop dls-inqpop"><b>'+head+'</b><table>'+rows+'</table>'+(links?'<div class="dls-inq-links">'+links+'</div>':'')+'</div>';
  }

  // Resolve the quote's folder under References/Quotes by QuoteNumber (startswith), like the job-folder resolver.
  private openQuoteFolder(quoteNo:string): void {
    if(!quoteNo){ this.setStatus('No quote number on this inquiry.'); return; }
    const base=this.context.pageContext.web.absoluteUrl; const host=base.replace(/\/sites\/.*$/,'');
    const spm=base.match(/(\/sites\/[^/]+)/); const sp=spm?spm[1]:''; const self=this;
    const w=window.open('','_blank'); try{ if(w) w.document.write('<p style="font:14px/1.4 sans-serif;padding:18px;color:#333">Opening quote folderâ¦</p>'); }catch(e){}
    const open=(url:string)=>{ try{ if(w) w.location=url; else window.open(url,'_blank'); }catch(e){} };
    const libUrl=host+(sp+IQ_QUOTES_REL).replace(/ /g,'%20');
    const cached=this._quoteFolderCache[quoteNo]; if(cached){ open(cached); this.setStatus('Opened quote folder.'); return; }
    const parentRel=(sp+IQ_QUOTES_REL).replace(/ /g,'%20');
    const u=base+"/_api/web/GetFolderByServerRelativeUrl('"+parentRel.replace(/'/g,"''")+"')/Folders?$select=Name,ServerRelativeUrl&$filter=startswith(Name,'"+quoteNo.replace(/'/g,"''")+"')&$top=1";
    this.spGet(u).then((d:any)=>{ const v=(d&&d.value)||[]; if(v.length){ const fu=host+v[0].ServerRelativeUrl; self._quoteFolderCache[quoteNo]=fu; open(fu); self.setStatus('Opened quote folder for '+quoteNo); } else { open(libUrl); self.setStatus('Quote folder for '+quoteNo+' not found â opened the Quotes library.'); } }).catch(()=>{ open(libUrl); });
  }

  // ---- Inquiries filter panel (mirrors the Projects panel) ----
  private buildInquiryPanel(): void {
    const el=this.domElement.querySelector('#inqHost') as any; if(!el) return;
    let h='<input id="iSearch" class="pp-search" placeholder="Search client / owner / address / quote #" value="'+esc(this.iqSearch)+'"/>';
    h+='<button class="pp-reset" id="iReset">Reset filters</button>';
    const dims=[['status','Quote status'],['county','County'],['year','Year']];
    for(let i=0;i<dims.length;i++){ const dim=dims[i][0], lbl=dims[i][1]; const cap=dim.charAt(0).toUpperCase()+dim.slice(1); h+='<div class="lp-sub'+(this._iqDimColl[dim]?' coll':'')+'" data-idimsec="'+dim+'"><div class="lp-subhd"><span class="tw">&#9662;</span> '+lbl+' <span class="pp-all" data-idim="'+dim+'">all</span></div><div class="lp-subbd"><div id="iq'+cap+'"></div></div></div>'; }
    el.innerHTML=h; const self=this;
    const si=el.querySelector('#iSearch') as any; if(si) si.oninput=function(){ self.iqSearch=(si.value||'').toLowerCase().trim(); self.renderInquiryPins(); };
    const rb=el.querySelector('#iReset') as any; if(rb) rb.onclick=function(){ self.iqReset(); };
    const subhds=el.querySelectorAll('.lp-subhd'); for(let i=0;i<subhds.length;i++){ subhds[i].addEventListener('click',function(e:any){ const a=(e.target&&e.target.closest)?e.target.closest('[data-idim]'):null; if(a){ self.iqAllOn(a.getAttribute('data-idim')); return; } const sub=this.parentNode; const dim=sub.getAttribute('data-idimsec'); self._iqDimColl[dim]=!self._iqDimColl[dim]; sub.classList.toggle('coll'); }); }
    this.renderInqSections();
  }
  private iqCounts(keyFn:any): any { const m:any={}; for(let i=0;i<this.inquiries.length;i++){ const k=keyFn(this.inquiries[i]); m[k]=(m[k]||0)+1; } return m; }
  private renderInqSections(): void { this.renderInqSection('#iqStatus', this.iqStatusOn, 'status'); this.renderInqSection('#iqCounty', this.iqCountyOn, 'county'); this.renderInqSection('#iqYear', this.iqYearOn, 'year'); }
  private renderInqSection(sel:string, state:any, dim:string): void {
    const el=this.domElement.querySelector(sel) as any; if(!el) return;
    const keyFn = dim==='status'?(q:any)=>q.status||'(blank)' : dim==='county'?(q:any)=>q.county||'(blank)' : (q:any)=>q.year||'(blank)';
    const counts=this.iqCounts(keyFn);
    const keys=Object.keys(counts).sort((a:any,b:any)=>{ if(dim==='status'){ const ia=IQ_STATUS_ORDER.indexOf(a),ib=IQ_STATUS_ORDER.indexOf(b); return (ia<0?99:ia)-(ib<0?99:ib); } if(dim==='year') return b.localeCompare(a); return a.localeCompare(b); });
    const self=this; el.innerHTML='';
    for(let i=0;i<keys.length;i++){ const k=keys[i]; const row=document.createElement('div'); row.className='pp-row'+(state[k]===false?' off':''); row.innerHTML='<span class="pp-sq" style="background:'+IQ_COLOR+'"></span><span class="pp-nm">'+esc(k||'(blank)')+'</span><span class="pp-n">'+counts[k]+'</span>'; row.onclick=function(){ state[k]=state[k]===false; self.renderInquiryPins(); self.renderInqSections(); }; el.appendChild(row); }
  }
  private iqAllOn(dim:string): void { const st = dim==='status'?this.iqStatusOn:dim==='county'?this.iqCountyOn:dim==='year'?this.iqYearOn:null; if(!st) return; const ks=Object.keys(st); for(let i=0;i<ks.length;i++) st[ks[i]]=true; this.renderInquiryPins(); this.renderInqSections(); }
  private iqReset(): void { const self=this; const ks=Object.keys(this.iqStatusOn); for(let i=0;i<ks.length;i++){ self.iqStatusOn[ks[i]]=IQ_DEFAULT_STATUS_ON.indexOf(ks[i])>=0; } [this.iqCountyOn,this.iqYearOn].forEach((st:any)=>{ const k2=Object.keys(st); for(let j=0;j<k2.length;j++) st[k2[j]]=true; }); this.iqSearch=''; const si=this.domElement.querySelector('#iSearch') as any; if(si) si.value=''; this.renderInquiryPins(); this.renderInqSections(); }

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
      PropertyPaneTextField('zoningAssetBase',{label:'Zoning overlays folder URL'}),
      PropertyPaneTextField('projectsListGuid',{label:'WIP / Projects list GUID'}),
      PropertyPaneTextField('inquiriesListGuid',{label:'Inquiries (IQ) list GUID'})
    ]}]}]};
  }
}
