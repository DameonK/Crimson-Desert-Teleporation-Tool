"""
Crimson Desert — Teleport Utility
=======================================

Requirements:
  - Python 3.10+
  - pymem  (pip install pymem)
  - Run as Administrator

Features:
  F5  Teleport to map destination marker
  F6  Save map marker as waypoint
  F8  Abort / return to pre-teleport position
  10s invulnerability after each teleport
  Waypoint manager with community sharing
"""

VERSION = "2.1.0"

import ctypes
import ctypes.wintypes
import struct
import os
import sys
import json
import time
import atexit
import tkinter as tk
from tkinter import ttk
from urllib.request import urlopen
from urllib.parse import quote_plus

# Add local pylibs folder to path (for pymem, Pillow, etc.)
_pylibs = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pylibs")
if os.path.isdir(_pylibs) and _pylibs not in sys.path:
    sys.path.insert(0, _pylibs)

try:
    import webview
    _HAS_WEBVIEW = True
except ImportError:
    _HAS_WEBVIEW = False

# ── Admin check ──────────────────────────────────────────────────────

def _is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False

if not _is_admin():
    ctypes.windll.user32.MessageBoxW(
        0,
        "This program must be run as Administrator.\n\n"
        "Right-click and select 'Run as administrator'.",
        "Crimson Desert Teleporter Tool", 0x10)
    sys.exit(1)

try:
    import pymem
    import pymem.process
except ImportError:
    ctypes.windll.user32.MessageBoxW(
        0, "pymem is required.\n\nInstall with:\n  pip install pymem",
        "Missing Dependency", 0x10)
    sys.exit(1)

# ── Constants ────────────────────────────────────────────────────────

PROCESS_NAME = "CrimsonDesert.exe"
SAVE_DIR = os.path.join(os.environ.get("LOCALAPPDATA", ""), "CD_Teleport")
SAVE_FILE = os.path.join(SAVE_DIR, "cd_waypoints.json")

# Google Sheets integration (from CE table)
SHARED_CSV_URL = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vRuCTPOpKood_wCToItMFiGYMjL4FxP6CAOWxNzcZKoNI3WUU06OmBqyECASUJ8SUSqh2KvPXaG-s6-"
    "/pub?gid=1303005004&single=true&output=csv"
)
FORM_SUBMIT_URL = (
    "https://docs.google.com/forms/d/e/"
    "1FAIpQLScdrT1RU4EKKOsbCpt5j2BUTpJEocbc7L4xR53lCDzpjrDfbQ/formResponse"
)
FORM_FIELDS = {
    "name": "entry.2135530741",
    "x": "entry.1438084253",
    "y": "entry.2086854493",
    "z": "entry.1815075034",
}

INVULN_SECONDS = 10
HEIGHT_BOOST = 10.0

# World map bounds (calibrated from real in-game positions)
WORLD_BOUNDS = {"x_min": -13628.0, "x_max": -2028.0,
                "z_min": -7624.0,  "z_max": 4276.0}
SETTINGS_FILE = os.path.join(SAVE_DIR, "cd_settings.json")

# Virtual key code <-> display name mapping
VK_NAMES = {
    0x70: "F1", 0x71: "F2", 0x72: "F3", 0x73: "F4", 0x74: "F5", 0x75: "F6",
    0x76: "F7", 0x77: "F8", 0x78: "F9", 0x79: "F10", 0x7A: "F11", 0x7B: "F12",
    0x21: "PgUp", 0x22: "PgDn", 0x23: "End", 0x24: "Home", 0x2D: "Insert",
    0x90: "NumLk", 0x60: "Num0", 0x61: "Num1", 0x62: "Num2", 0x63: "Num3",
    0x64: "Num4", 0x65: "Num5", 0x66: "Num6", 0x67: "Num7", 0x68: "Num8",
    0x69: "Num9", 0x6A: "Num*", 0x6B: "Num+", 0x6D: "Num-", 0x6E: "Num.",
    0x6F: "Num/", 0x20: "Space", 0x09: "Tab",
}
# Add letter keys A-Z and number keys 0-9
for _i in range(0x41, 0x5B):
    VK_NAMES[_i] = chr(_i)
for _i in range(0x30, 0x3A):
    VK_NAMES[_i] = str(_i - 0x30)
VK_FROM_NAME = {v: k for k, v in VK_NAMES.items()}

# Modifier keys (not valid as primary keys, only as modifiers)
VK_MOD_CTRL  = 0x11
VK_MOD_ALT   = 0x12
VK_MOD_SHIFT = 0x10
MOD_NAMES = {VK_MOD_CTRL: "Ctrl", VK_MOD_ALT: "Alt", VK_MOD_SHIFT: "Shift"}
MOD_VKS = (VK_MOD_CTRL, VK_MOD_ALT, VK_MOD_SHIFT)

def _hotkey_display(vk, mod=0):
    """Return display string like 'Ctrl+F5' or 'F5'."""
    key_name = VK_NAMES.get(vk, f"0x{vk:02X}")
    if mod and mod in MOD_NAMES:
        return f"{MOD_NAMES[mod]}+{key_name}"
    return key_name

DEFAULT_HOTKEYS = {
    "teleport": {"vk": 0x74, "mod": 0, "enabled": True},   # F5
    "save":     {"vk": 0x75, "mod": 0, "enabled": True},    # F6
    "abort":    {"vk": 0x77, "mod": 0, "enabled": True},    # F8
}

def _load_settings():
    try:
        with open(SETTINGS_FILE, 'r') as f:
            return json.load(f)
    except Exception:
        return {}

def _save_settings(data):
    os.makedirs(SAVE_DIR, exist_ok=True)
    with open(SETTINGS_FILE, 'w') as f:
        json.dump(data, f, indent=2)

# ── Windows API setup ────────────────────────────────────────────────

k32 = ctypes.windll.kernel32

# VirtualAllocEx must return 64-bit pointer
k32.VirtualAllocEx.restype = ctypes.c_ulonglong
k32.VirtualAllocEx.argtypes = [
    ctypes.c_void_p, ctypes.c_ulonglong, ctypes.c_size_t,
    ctypes.c_ulong, ctypes.c_ulong,
]
k32.VirtualFreeEx.argtypes = [
    ctypes.c_void_p, ctypes.c_ulonglong, ctypes.c_size_t, ctypes.c_ulong,
]

MEM_COMMIT  = 0x1000
MEM_RESERVE = 0x2000
MEM_RELEASE = 0x8000
PAGE_EXECUTE_READWRITE = 0x40


# ── Interactive Map ──────────────────────────────────────────────────

MAP_URL = "https://mapgenie.io/crimson-desert/maps/pywel"

# JavaScript injected after map page loads.
# Adds a transparent teleporter overlay (player marker, waypoints,
# destination marker, control bar) on top of the real website map.
_MAP_INJECT_JS = r"""(function(){
// Coordinate conversion (game <-> MapGenie Leaflet latlng)
// Default calibration — overridden by saved calibration from settings
var MG_LZ=0.0000408926,MG_OZ=0.776343;  // lat = MG_LZ * gameZ + MG_OZ
var MG_LX=0.0000416254,MG_OX=-0.400665; // lng = MG_LX * gameX + MG_OX
var _calMode=false;

function g2ll(gx,gz){return [MG_LZ*gz+MG_OZ, MG_LX*gx+MG_OX];}
function ll2g(lat,lng){return [(lng-MG_OX)/MG_LX, (lat-MG_OZ)/MG_LZ];}
// Called from Python to update calibration constants
window.setCalibration=function(lz,oz,lx,ox){MG_LZ=lz;MG_OZ=oz;MG_LX=lx;MG_OX=ox;reposAll();};

// State
var _map=null, dest=null, curH=1200;
var playerGX=null, playerGZ=null;
var waypoints=[];
var wpEls=[];

// Inject CSS
var sty=document.createElement('style');
sty.textContent=
 '#tp-bar{position:fixed;top:0;left:0;right:0;z-index:100000;display:flex;align-items:center;'+
 'padding:6px 12px;background:rgba(24,24,37,.92);border-bottom:1px solid #313244;gap:8px;'+
 'font-family:"Segoe UI",sans-serif;font-size:13px;color:#cdd6f4;backdrop-filter:blur(8px)}'+
 '#tp-bar label{font-weight:bold}'+
 '#tp-bar .cd{font-family:Consolas,monospace;color:#a6adc8;min-width:70px}'+
 '#tp-bar button{background:#45475a;color:#cdd6f4;border:1px solid #585b70;border-radius:4px;'+
 'padding:4px 12px;cursor:pointer;font-size:12px}'+
 '#tp-bar button:hover{background:#585b70}'+
 '#tp-bar button.pri{background:#89b4fa;color:#1e1e2e;border-color:#89b4fa;font-weight:bold}'+
 '#tp-bar button.pri:hover{background:#74c7ec}'+
 '#tp-bar .sp{flex:1}'+
 '#tp-sts{position:fixed;bottom:0;left:0;right:0;z-index:100000;padding:4px 12px;font-size:11px;'+
 'color:#a6adc8;background:rgba(24,24,37,.92);border-top:1px solid #313244;'+
 'font-family:"Segoe UI",sans-serif;backdrop-filter:blur(8px)}'+
 '#tp-ol{position:fixed;top:0;left:0;width:100%;height:100%;z-index:99999;pointer-events:none}'+
 '.tp-m{position:absolute;transform:translate(-50%,-50%);pointer-events:none}'+
 '.tp-m-p{width:14px;height:14px;background:#ff1e1e;border:2px solid white;border-radius:50%;'+
 'box-shadow:0 0 6px rgba(255,30,30,.5)}'+
 '.tp-m-d{width:14px;height:14px;background:#3278ff;border:2px solid white;border-radius:50%;'+
 'box-shadow:0 0 6px rgba(50,120,255,.5)}'+
 '.tp-m-w{width:8px;height:8px;background:#f9e2af;border:1px solid #b0a088;border-radius:50%}'+
 '#tp-hdlg{display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.6);'+
 'z-index:200000;justify-content:center;align-items:center}'+
 '#tp-hdlg.open{display:flex}'+
 '#tp-hdlg .dlg{background:#1e1e2e;border:1px solid #45475a;border-radius:8px;padding:20px;'+
 'min-width:340px;color:#cdd6f4;font-family:"Segoe UI",sans-serif}'+
 '#tp-hdlg .dlg h3{margin:0 0 8px;font-size:14px}'+
 '#tp-hdlg .dlg .note{color:#6c7086;font-size:12px;margin-bottom:12px}'+
 '#tp-hdlg input{background:#313244;border:1px solid #45475a;color:#cdd6f4;padding:6px 10px;'+
 'border-radius:4px;font-family:Consolas,monospace;font-size:13px;width:100%;margin-bottom:12px;'+
 'box-sizing:border-box}'+
 '#tp-hdlg input:focus{outline:none;border-color:#89b4fa}'+
 '.tp-pre{display:flex;gap:8px;margin-bottom:12px}'+
 '.tp-pre button{flex:1;background:#45475a;color:#cdd6f4;border:1px solid #585b70;'+
 'border-radius:4px;padding:4px 12px;cursor:pointer;font-size:12px}'+
 '.tp-pre button:hover{background:#585b70}'+
 '.tp-acts{display:flex;gap:8px;justify-content:flex-end}'+
 '.tp-acts button{background:#45475a;color:#cdd6f4;border:1px solid #585b70;'+
 'border-radius:4px;padding:4px 12px;cursor:pointer;font-size:12px}'+
 '.tp-acts button:hover{background:#585b70}';
document.head.appendChild(sty);

// Top bar
var bar=document.createElement('div');bar.id='tp-bar';
bar.innerHTML='<label>Destination:</label>'+
 '<span>X:</span><span class="cd" id="tp-dx">---</span>'+
 '<span>Z:</span><span class="cd" id="tp-dz">---</span>'+
 '<span>Height:</span><span class="cd" id="tp-dh">1200</span>'+
 '<span style="width:10px"></span>'+
 '<button class="pri" id="tp-go">Teleport</button>'+
 '<button id="tp-sh">Set Height</button>'+
 '<span class="sp"></span>'+
 '<button id="tp-cal">Calibrate</button>';
document.body.appendChild(bar);

// Status bar
var sts=document.createElement('div');sts.id='tp-sts';
sts.textContent='Teleporter overlay \u2014 waiting for map...';
document.body.appendChild(sts);

// Marker overlay
var ol=document.createElement('div');ol.id='tp-ol';
document.body.appendChild(ol);
var pMk=document.createElement('div');pMk.className='tp-m';
pMk.innerHTML='<div class="tp-m-p"></div>';pMk.style.display='none';ol.appendChild(pMk);
var dMk=document.createElement('div');dMk.className='tp-m';
dMk.innerHTML='<div class="tp-m-d"></div>';dMk.style.display='none';ol.appendChild(dMk);

// Height dialog
var hdlg=document.createElement('div');hdlg.id='tp-hdlg';
hdlg.innerHTML='<div class="dlg"><h3>Set Teleport Height</h3>'+
 '<div class="note">Map height is not automatically derived from map data.</div>'+
 '<input type="number" id="tp-hin" value="1200">'+
 '<div class="tp-pre"><button id="tp-hg">Ground (1200)</button>'+
 '<button id="tp-ha">Abyss (2500)</button></div>'+
 '<div class="tp-acts"><button id="tp-hc">Cancel</button>'+
 '<button id="tp-hok">OK</button></div></div>';
document.body.appendChild(hdlg);

function applyH(v){v=parseFloat(v);if(isNaN(v))return;curH=v;
 document.getElementById('tp-dh').textContent=v.toFixed(0);
 hdlg.classList.remove('open');
 if(window.pywebview)window.pywebview.api.set_height(v);}
document.getElementById('tp-hg').onclick=function(){applyH(1200)};
document.getElementById('tp-ha').onclick=function(){applyH(2500)};
document.getElementById('tp-hc').onclick=function(){hdlg.classList.remove('open')};
document.getElementById('tp-hok').onclick=function(){applyH(document.getElementById('tp-hin').value)};
document.getElementById('tp-hin').addEventListener('keydown',function(e){
 if(e.key==='Enter')applyH(this.value);if(e.key==='Escape')hdlg.classList.remove('open');});
document.getElementById('tp-sh').onclick=function(){
 document.getElementById('tp-hin').value=curH;hdlg.classList.add('open');
 var i=document.getElementById('tp-hin');i.focus();i.select();};
document.getElementById('tp-go').onclick=function(){
 if(!dest){sts.textContent='Set a destination on the map first.';return;}
 if(window.pywebview)window.pywebview.api.teleport();};
document.getElementById('tp-cal').onclick=function(){
 _calMode=true;
 document.getElementById('tp-cal').style.background='#f38ba8';
 document.getElementById('tp-cal').textContent='Click your position...';
 sts.textContent='CALIBRATE: Click on your exact current position on the map.';};

// ── Find Leaflet map instance ──
function isLeaflet(v){try{return v&&typeof v.getCenter==='function'&&typeof v.getZoom==='function'&&typeof v.on==='function'&&(typeof v.latLngToContainerPoint==='function'||typeof v.containerPointToLatLng==='function');}catch(e){return false;}}
// Broader check: any map-like object (Leaflet, MapLibre, or wrapper)
function isMapLike(v){try{return v&&typeof v.getCenter==='function'&&typeof v.getZoom==='function'&&typeof v.on==='function'&&typeof v.getContainer==='function';}catch(e){return false;}}
function wrapLeaflet(lm){
 // Make Leaflet map look like MapLibre for our overlay code
 return {_lm:lm,_isLeaflet:true,
  getCenter:function(){return lm.getCenter();},
  getZoom:function(){return lm.getZoom();},
  getContainer:function(){return lm.getContainer();},
  project:function(lngLat){
   // MapLibre: project([lng,lat]) -> {x,y} pixel. Leaflet: latLngToContainerPoint([lat,lng])
   var pt=lm.latLngToContainerPoint([lngLat[1]!==undefined?lngLat[1]:lngLat.lat,lngLat[0]!==undefined?lngLat[0]:lngLat.lng]);
   return {x:pt.x,y:pt.y};
  },
  on:function(ev,fn){
   if(ev==='click'){
    lm.on('click',function(e){fn({lngLat:{lat:e.latlng.lat,lng:e.latlng.lng}});});
   }else{lm.on(ev,fn);}
   return this;
  }
 };
}

// Strategy L: Leaflet-specific detection (most game map sites use Leaflet)
function extractLeafletFromEl(el){
 // Leaflet stores map ref on container via internal properties
 try{var keys=Object.getOwnPropertyNames(el);
  for(var j=0;j<keys.length;j++){try{var v=el[keys[j]];if(isLeaflet(v))return v;}catch(e){}}
 }catch(e){}
 try{if(isLeaflet(el._leaflet_map))return el._leaflet_map;}catch(e){}
 try{if(isLeaflet(el._map))return el._map;}catch(e){}
 return null;
}
function findLeaflet(){
 // Collect ALL Leaflet map instances from containers, pick the largest (main map)
 var containers=document.querySelectorAll('.leaflet-container');
 if(containers.length>0){
  var best=null,bestArea=0;
  for(var i=0;i<containers.length;i++){
   var lm=extractLeafletFromEl(containers[i]);
   if(lm){
    var rect=containers[i].getBoundingClientRect();
    var area=rect.width*rect.height;
    if(area>bestArea){bestArea=area;best=lm;}
   }
  }
  if(best)return wrapLeaflet(best);
 }
 // Check global L and common globals
 if(typeof L!=='undefined'){
  var names=['map','_map','mapInstance','leafletMap'];
  for(var i=0;i<names.length;i++){try{if(isLeaflet(window[names[i]]))return wrapLeaflet(window[names[i]]);}catch(e){}}
 }
 // Walk window properties for Leaflet maps
 try{var wkeys=Object.getOwnPropertyNames(window);
  for(var i=0;i<wkeys.length;i++){try{var v=window[wkeys[i]];
   if(isLeaflet(v))return wrapLeaflet(v);
   if(v&&typeof v==='object'&&!Array.isArray(v)){
    var vk=Object.keys(v);for(var j=0;j<Math.min(vk.length,30);j++){
     try{if(isLeaflet(v[vk[j]]))return wrapLeaflet(v[vk[j]]);}catch(e){}}
   }
  }catch(e){}}
 }catch(e){}
 return null;
}

// Collect ALL map-like objects from any source
function collectAllMaps(){
 var maps=[];var seen=new WeakSet();
 function add(m,src){if(m&&!seen.has(m)){try{seen.add(m);}catch(e){return;}maps.push({m:m,src:src});}}
 // From .leaflet-container elements
 var containers=document.querySelectorAll('.leaflet-container');
 for(var i=0;i<containers.length;i++){
  var el=containers[i];
  var lm=extractLeafletFromEl(el);
  if(lm){add(lm,'container-'+i);continue;}
  try{var keys=Object.getOwnPropertyNames(el);
   for(var j=0;j<keys.length;j++){try{var v=el[keys[j]];if(isMapLike(v))add(v,'el-prop-'+i);}catch(e){}}
  }catch(e){}
 }
 // From window globals
 var names=['map','_map','mapInstance','glMap','leafletMap','$map'];
 for(var i=0;i<names.length;i++){try{var v=window[names[i]];if(isMapLike(v))add(v,'window.'+names[i]);}catch(e){}}
 // From window property scan (2 levels)
 try{var wkeys=Object.getOwnPropertyNames(window);
  for(var i=0;i<wkeys.length;i++){try{var v=window[wkeys[i]];
   if(v&&typeof v==='object'&&isMapLike(v))add(v,'win.'+wkeys[i]);
   if(v&&typeof v==='object'&&!Array.isArray(v)){
    try{var vk=Object.keys(v);for(var j=0;j<Math.min(vk.length,30);j++){
     try{if(vk[j]&&v[vk[j]]&&isMapLike(v[vk[j]]))add(v[vk[j]],'win.'+wkeys[i]+'.'+vk[j]);}catch(e){}
    }}catch(e){}
   }
  }catch(e){}}
 }catch(e){}
 return maps;
}
function findMap(){
 var all=collectAllMaps();
 if(all.length===0)return null;
 // Pick the map with the largest visible container (main world map)
 var best=null,bestArea=0;
 var diag=[];
 for(var i=0;i<all.length;i++){
  var m=all[i].m;
  try{
   var cont=m.getContainer?m.getContainer():null;
   var area=0;
   if(cont){var r=cont.getBoundingClientRect();area=r.width*r.height;}
   var c=m.getCenter?m.getCenter():{lat:0,lng:0};
   diag.push(all[i].src+'('+Math.round(area)+'px lat='+Number(c.lat).toFixed(3)+')');
   if(area>bestArea){bestArea=area;best=m;}
  }catch(e){}
 }
 window.__tp_mapDiag=diag.join(', ');
 if(best){if(isLeaflet(best))return wrapLeaflet(best);return best;}
 // Fallback: just return the first one
 var m=all[0].m;
 if(isLeaflet(m))return wrapLeaflet(m);return m;
}

function posMk(mk,gx,gz){
 if(!_map)return;
 var ll=g2ll(gx,gz);
 try{var pt=_map.project([ll[1],ll[0]]);
  var rect=_map.getContainer().getBoundingClientRect();
  mk.style.left=(rect.left+pt.x)+'px';
  mk.style.top=(rect.top+pt.y)+'px';
  mk.style.display='';
 }catch(e){mk.style.display='none';}
}
function reposAll(){
 if(playerGX!==null)posMk(pMk,playerGX,playerGZ);
 if(dest)posMk(dMk,dest.x,dest.z);
 wpEls.forEach(function(w){posMk(w.el,w.x,w.z);});
}
function initWithMap(map){
 _map=map;
 map.on('move',reposAll);
 map.on('moveend',reposAll);
 map.on('resize',reposAll);
 map.on('click',function(e){
  if(_calMode){
   _calMode=false;
   document.getElementById('tp-cal').style.background='';
   document.getElementById('tp-cal').textContent='Calibrate';
   sts.textContent='Calibrating... raw('+e.lngLat.lat.toFixed(6)+', '+e.lngLat.lng.toFixed(6)+')';
   if(window.pywebview)window.pywebview.api.calibrate(e.lngLat.lat,e.lngLat.lng);
   return;
  }
  var gc=ll2g(e.lngLat.lat,e.lngLat.lng);
  dest={x:gc[0],z:gc[1]};
  document.getElementById('tp-dx').textContent=gc[0].toFixed(1);
  document.getElementById('tp-dz').textContent=gc[1].toFixed(1);
  posMk(dMk,gc[0],gc[1]);
  sts.textContent='Destination: ('+gc[0].toFixed(1)+', '+gc[1].toFixed(1)+') Height: '+curH;
  if(window.pywebview)window.pywebview.api.on_destination_set(gc[0],gc[1]);
 });
 // Dump coordinate system info for calibration
 var info='Map found';
 try{
  var raw=map._isLeaflet?map._lm:map;
  var cont=raw.getContainer?raw.getContainer():null;
  if(cont){var r=cont.getBoundingClientRect();info+=' ['+Math.round(r.width)+'x'+Math.round(r.height)+']';}
  var c=raw.getCenter?raw.getCenter():{lat:'?',lng:'?'};
  info+=' center('+Number(c.lat).toFixed(3)+','+Number(c.lng).toFixed(3)+') z='+raw.getZoom();
 }catch(e){info+=' (diag: '+e.message+')';}
 if(window.__tp_mapDiag)info+=' | found: '+window.__tp_mapDiag;
 sts.textContent=info+' \u2014 click map to set destination';
 reposAll();
}

// Poll for Leaflet map instance (site may still be loading)
var pc=0;
var pi=setInterval(function(){
 var m=findMap();
 if(m){clearInterval(pi);initWithMap(m);return;}
 pc++;
 var hasLeaflet=!!document.querySelector('.leaflet-container');
 if(pc<=60){
  sts.textContent='Searching for map... ('+pc+') leaflet:'+hasLeaflet;
 }else{clearInterval(pi);
  sts.textContent='Map instance not found \u2014 overlay markers unavailable. Teleporter still works via coordinates.';}
},500);

// API called from Python
window.updatePlayer=function(gx,gz){playerGX=gx;playerGZ=gz;posMk(pMk,gx,gz);};
window.updateWaypoints=function(wps){
 wpEls.forEach(function(w){ol.removeChild(w.el);});wpEls=[];
 wps.forEach(function(wp){if(wp.x===0&&wp.z===0)return;
  var el=document.createElement('div');el.className='tp-m';
  el.innerHTML='<div class="tp-m-w" title="'+(wp.name||'').replace(/"/g,'&quot;')+'"></div>';
  ol.appendChild(el);wpEls.push({el:el,x:wp.x,z:wp.z});
  posMk(el,wp.x,wp.z);
 });
};
window.setHeight=function(h){curH=h;document.getElementById('tp-dh').textContent=h.toFixed(0);};
window.setStatus=function(s){sts.textContent=s;};
})();"""


class _MapBridge:
    """JavaScript bridge for the pywebview map window."""

    def __init__(self, app):
        self._app = app

    def on_destination_set(self, x, z):
        x, z = float(x), float(z)
        self._app._map_dest = (x, self._app._map_height, z)
        self._app.after(0, lambda: self._app.bottom_var.set(
            f"Map dest: ({x:.1f}, {z:.1f})"))

    def teleport(self):
        self._app.after(0, self._app._map_teleport)

    def set_height(self, height):
        height = float(height)
        self._app._map_height = height
        if self._app._map_dest:
            x, _, z = self._app._map_dest
            self._app._map_dest = (x, height, z)
        def _save():
            settings = _load_settings()
            settings["map_height"] = height
            _save_settings(settings)
        self._app.after(0, _save)

    def calibrate(self, lat, lng):
        """Receive a calibration click: lat/lng where the player actually is."""
        lat, lng = float(lat), float(lng)
        app = self._app
        def _do_calibrate():
            # Get current player game coordinates
            eng = app.engine
            if not eng or not eng.pos_addr:
                app.bottom_var.set("Calibration failed: not connected")
                return
            gx, _, gz = eng.read_position()
            # Store calibration point
            settings = _load_settings()
            cal_pts = settings.get("map_cal_points", [])
            cal_pts.append({"gx": gx, "gz": gz, "lat": lat, "lng": lng})
            # Keep last 5 points
            cal_pts = cal_pts[-5:]
            settings["map_cal_points"] = cal_pts
            if len(cal_pts) >= 2:
                # Least-squares fit: lat = lz*gz + oz, lng = lx*gx + ox
                n = len(cal_pts)
                sx = sum(p["gx"] for p in cal_pts)
                sz = sum(p["gz"] for p in cal_pts)
                slat = sum(p["lat"] for p in cal_pts)
                slng = sum(p["lng"] for p in cal_pts)
                sxlng = sum(p["gx"] * p["lng"] for p in cal_pts)
                szlat = sum(p["gz"] * p["lat"] for p in cal_pts)
                sx2 = sum(p["gx"] ** 2 for p in cal_pts)
                sz2 = sum(p["gz"] ** 2 for p in cal_pts)
                det_z = n * sz2 - sz * sz
                det_x = n * sx2 - sx * sx
                if abs(det_z) > 1e-12 and abs(det_x) > 1e-12:
                    lz = (n * szlat - sz * slat) / det_z
                    oz = (slat - lz * sz) / n
                    lx = (n * sxlng - sx * slng) / det_x
                    ox = (slng - lx * sx) / n
                    settings["map_cal"] = {"lz": lz, "oz": oz,
                                           "lx": lx, "ox": ox}
                    _save_settings(settings)
                    # Push new constants to JS
                    if app._map_webview:
                        app._map_webview.evaluate_js(
                            f"setCalibration({lz},{oz},{lx},{ox})")
                    app.bottom_var.set(
                        f"Calibrated with {n} points")
                    return
            _save_settings(settings)
            app.bottom_var.set(
                f"Calibration point saved ({len(cal_pts)}/2 needed)")
        self._app.after(0, _do_calibrate)


# ── TeleportEngine ───────────────────────────────────────────────────

class TeleportEngine:
    """Process attachment, code-cave injection, and memory operations."""

    # AOB signatures (from CE table)
    AOB_ENTITY = b'\x48\x8B\x06\x0F\x11\x88\xB0\x01\x00\x00'
    AOB_POS    = b'\x0F\x11\x99\x90\x00\x00\x00'
    AOB_HEALTH = b'\x48\x8B\x46\x08\x48\x89\xF1'
    AOB_MAP    = b'\xF2\x0F\x11\x02\x8B\x47\x08\x89\x42\x08\x80'
    AOB_WORLD  = b'\x0F\x5C\x1D'  # prefix for world-offset constant

    # Memory block layout
    OFF_TD   = 0x000   # teleportData  (64 bytes)
    OFF_INV  = 0x040   # invulnFlag    (16 bytes)
    OFF_MD   = 0x050   # mapDestData   (16 bytes)
    OFF_CA   = 0x100   # cave A        (128 bytes)
    OFF_CB   = 0x180   # cave B        (128 bytes)
    OFF_CC   = 0x200   # cave C        (128 bytes)
    OFF_CD   = 0x280   # cave D        (128 bytes)
    BLOCK_SZ = 0x1000  # 4096 bytes

    def __init__(self):
        self.pm = None
        self.module = None
        self.attached = False
        self.hooks_installed = False

        self.block = 0          # allocated memory block
        self.td = 0             # teleportData addr
        self.inv = 0            # invulnFlag addr
        self.md = 0             # mapDestData addr

        self.hook_a = 0         # entity hook point (AOB_ENTITY + 3)
        self.hook_b = 0         # position block hook point
        self.hook_c = 0         # health hook point
        self.hook_d = 0         # map dest hook point

        self.orig_bytes = {}    # addr -> original bytes for unhooking
        self.world_offset_addr = 0
        self._trampolines = []  # far-mode trampoline allocations to free
        self._far_mode = False  # True if block is far from hooks

    # ── attach / detach ──────────────────────────────────────────────

    def attach(self):
        self.pm = pymem.Pymem(PROCESS_NAME)
        self.module = pymem.process.module_from_name(
            self.pm.process_handle, PROCESS_NAME)
        self.attached = True

    def detach(self):
        if self.hooks_installed:
            self.uninstall_hooks()
        handle = self.pm.process_handle if self.pm else None
        if self.block and handle:
            k32.VirtualFreeEx(handle, self.block, 0, MEM_RELEASE)
            self.block = 0
        for tramp in self._trampolines:
            if handle:
                try:
                    k32.VirtualFreeEx(handle, tramp, 0, MEM_RELEASE)
                except Exception:
                    pass
        self._trampolines.clear()
        self._far_mode = False
        if self.pm:
            self.pm.close_process()
        self.attached = False

    # ── pattern scanning ─────────────────────────────────────────────

    def _read_module(self):
        base = self.module.lpBaseOfDll
        size = self.module.SizeOfImage
        data = bytearray(size)
        CHUNK = 0x10000
        for off in range(0, size, CHUNK):
            sz = min(CHUNK, size - off)
            try:
                data[off:off + sz] = self.pm.read_bytes(base + off, sz)
            except Exception:
                pass
        return bytes(data), base

    def scan_and_hook(self):
        data, base = self._read_module()

        # Entity hook: AOB + 3
        idx = data.find(self.AOB_ENTITY)
        if idx == -1:
            raise RuntimeError("Entity hook AOB not found — game version mismatch?")
        self.hook_a = base + idx + 3

        # Position block
        idx = data.find(self.AOB_POS)
        if idx == -1:
            raise RuntimeError("Position block AOB not found")
        self.hook_b = base + idx

        # Health hook
        idx = data.find(self.AOB_HEALTH)
        if idx == -1:
            raise RuntimeError("Health hook AOB not found")
        self.hook_c = base + idx

        # Map destination
        idx = data.find(self.AOB_MAP)
        if idx == -1:
            raise RuntimeError("Map dest AOB not found")
        self.hook_d = base + idx

        # World offset: find 0F 5C 1D ?? ?? ?? ?? 0F 11 99 90 00 00 00
        suffix = b'\x0F\x11\x99\x90\x00\x00\x00'
        pos = 0
        while pos < len(data) - 14:
            i = data.find(self.AOB_WORLD, pos)
            if i == -1:
                break
            if data[i + 7:i + 14] == suffix:
                disp = struct.unpack_from('<i', data, i + 3)[0]
                self.world_offset_addr = base + i + 7 + disp
                break
            pos = i + 1

        # Allocate & install
        self._alloc_block()
        self._install_hooks()

    # ── memory allocation ────────────────────────────────────────────

    def _alloc_near(self, handle, near, size):
        """Try to allocate memory within +/-2GB of `near`."""
        for offset in range(0x10000, 0x7FFF0000, 0x10000):
            for addr in [near + offset, near - offset]:
                if addr <= 0:
                    continue
                result = k32.VirtualAllocEx(
                    handle, addr, size,
                    MEM_COMMIT | MEM_RESERVE, PAGE_EXECUTE_READWRITE)
                if result:
                    return result
        return 0

    def _alloc_block(self):
        handle = self.pm.process_handle

        # Try near allocation using each hook address as anchor
        for anchor in [self.hook_a, self.hook_b, self.hook_c, self.hook_d]:
            if not anchor:
                continue
            result = self._alloc_near(handle, anchor, self.BLOCK_SZ)
            if result:
                self.block = result
                self.td  = result + self.OFF_TD
                self.inv = result + self.OFF_INV
                self.md  = result + self.OFF_MD
                self._far_mode = False
                return

        # Fallback: allocate anywhere (let Windows choose)
        result = k32.VirtualAllocEx(
            handle, 0, self.BLOCK_SZ,
            MEM_COMMIT | MEM_RESERVE, PAGE_EXECUTE_READWRITE)
        if not result:
            raise RuntimeError(
                "Could not allocate memory for code caves.\n"
                "Close other programs that hook into the game "
                "(trainers, overlays, cheat tables) and try again.")
        self.block = result
        self.td  = result + self.OFF_TD
        self.inv = result + self.OFF_INV
        self.md  = result + self.OFF_MD
        self._far_mode = True

    # ── hook installation ────────────────────────────────────────────

    def _rel32(self, from_addr, to_addr):
        rel = to_addr - (from_addr + 5)
        if not (-0x80000000 <= rel <= 0x7FFFFFFF):
            raise RuntimeError(
                f"Cave too far for rel32 jump: {from_addr:#x} -> {to_addr:#x}")
        return rel

    def _jmp_patch(self, from_addr, to_addr):
        return b'\xE9' + struct.pack('<i', self._rel32(from_addr, to_addr)) + b'\x90\x90'

    def _abs_jmp(self, target):
        """FF 25 00 00 00 00 + 8-byte address (jmp [rip+0])"""
        return b'\xFF\x25\x00\x00\x00\x00' + struct.pack('<Q', target)

    def _build_cave_a(self):
        """Entity capture + velocity teleport cave."""
        td = self.td
        ret = self.hook_a + 7
        c = bytearray()
        c += b'\x51'                                            # push rcx
        c += b'\x48\xB9' + struct.pack('<Q', td)                # mov rcx, teleportData
        c += b'\x48\x89\x41\x18'                                # mov [rcx+18], rax
        c += b'\xF3\x0F\x10\x80\x90\x00\x00\x00'               # movss xmm0, [rax+90]
        c += b'\xF3\x0F\x11\x41\x20'                            # movss [rcx+20], xmm0
        c += b'\xF3\x0F\x10\x80\x94\x00\x00\x00'               # movss xmm0, [rax+94]
        c += b'\xF3\x0F\x11\x41\x24'                            # movss [rcx+24], xmm0
        c += b'\xF3\x0F\x10\x80\x98\x00\x00\x00'               # movss xmm0, [rax+98]
        c += b'\xF3\x0F\x11\x41\x28'                            # movss [rcx+28], xmm0
        c += b'\x83\x79\x10\x00'                                # cmp [rcx+10], 0
        c += b'\x7E\x06'                                        # jle skip (+6)
        c += b'\x0F\x10\x09'                                    # movups xmm1, [rcx]
        c += b'\xFF\x49\x10'                                    # dec [rcx+10]
        # skip:
        c += b'\x59'                                            # pop rcx
        c += b'\x0F\x11\x88\xB0\x01\x00\x00'                   # movups [rax+1B0], xmm1
        c += self._abs_jmp(ret)
        return bytes(c)

    def _build_cave_b(self):
        """Position block during teleport cave."""
        td = self.td
        ret = self.hook_b + 7
        c = bytearray()
        c += b'\x50'                                            # push rax
        c += b'\x48\xB8' + struct.pack('<Q', td)                # mov rax, teleportData
        c += b'\x83\x78\x10\x00'                                # cmp [rax+10], 0
        c += b'\x7E\x15'                                        # jle doOriginal (+21)
        c += b'\x48\x3B\x48\x18'                                # cmp rcx, [rax+18]
        c += b'\x75\x0F'                                        # jne doOriginal (+15)
        c += b'\x58'                                            # pop rax
        c += self._abs_jmp(ret)                                 # skip original
        # doOriginal:
        c += b'\x58'                                            # pop rax
        c += b'\x0F\x11\x99\x90\x00\x00\x00'                   # movups [rcx+90], xmm3
        c += self._abs_jmp(ret)
        return bytes(c)

    def _build_cave_c(self):
        """Health invulnerability cave."""
        inv = self.inv
        ret = self.hook_c + 7
        c = bytearray()
        c += b'\x53'                                            # push rbx
        c += b'\x48\xBB' + struct.pack('<Q', inv)               # mov rbx, invulnFlag
        c += b'\x80\x3B\x01'                                    # cmp byte [rbx], 1
        c += b'\x5B'                                            # pop rbx
        c += b'\x75\x0F'                                        # jne orig (+15)
        c += b'\x80\x3E\x00'                                    # cmp byte [rsi], 0
        c += b'\x75\x0A'                                        # jne orig (+10)
        c += b'\x53'                                            # push rbx
        c += b'\x48\x8B\x5E\x18'                                # mov rbx, [rsi+18]
        c += b'\x48\x89\x5E\x08'                                # mov [rsi+08], rbx
        c += b'\x5B'                                            # pop rbx
        # originalCode:
        c += b'\x48\x8B\x46\x08'                                # mov rax, [rsi+08]
        c += b'\x48\x89\xF1'                                    # mov rcx, rsi
        c += self._abs_jmp(ret)
        return bytes(c)

    def _build_cave_d(self):
        """Map destination capture cave."""
        md = self.md
        ret = self.hook_d + 7
        c = bytearray()
        c += b'\xF2\x0F\x11\x02'                               # movsd [rdx], xmm0
        c += b'\x8B\x47\x08'                                    # mov eax, [rdi+08]
        c += b'\x51'                                            # push rcx
        c += b'\x48\xB9' + struct.pack('<Q', md)                # mov rcx, mapDestData
        c += b'\xF2\x0F\x11\x01'                                # movsd [rcx], xmm0
        c += b'\x89\x41\x08'                                    # mov [rcx+08], eax
        c += b'\xC7\x41\x0C\x01\x00\x00\x00'                   # mov dword [rcx+0C], 1
        c += b'\x59'                                            # pop rcx
        c += self._abs_jmp(ret)
        return bytes(c)

    def _install_hooks(self):
        handle = self.pm.process_handle

        # Initialize teleportData: height boost at +30
        init = bytearray(64)
        struct.pack_into('<f', init, 0x30, HEIGHT_BOOST)
        self.pm.write_bytes(self.td, bytes(init), len(init))

        # Initialize invulnFlag = 0
        self.pm.write_bytes(self.inv, b'\x00', 1)

        # Initialize mapDestData = zeros
        self.pm.write_bytes(self.md, bytes(16), 16)

        # Write caves
        caves = [
            (self.OFF_CA, self._build_cave_a()),
            (self.OFF_CB, self._build_cave_b()),
            (self.OFF_CC, self._build_cave_c()),
            (self.OFF_CD, self._build_cave_d()),
        ]
        for off, code in caves:
            addr = self.block + off
            self.pm.write_bytes(addr, code, len(code))

        # Save original bytes and install JMP patches
        hooks = [
            (self.hook_a, self.block + self.OFF_CA),
            (self.hook_b, self.block + self.OFF_CB),
            (self.hook_c, self.block + self.OFF_CC),
            (self.hook_d, self.block + self.OFF_CD),
        ]

        if not self._far_mode:
            # Near mode: direct rel32 jumps from hook -> cave
            for hook_addr, cave_addr in hooks:
                self.orig_bytes[hook_addr] = self.pm.read_bytes(hook_addr, 7)
                patch = self._jmp_patch(hook_addr, cave_addr)
                self.pm.write_bytes(hook_addr, patch, 7)
        else:
            # Far mode: hook -> trampoline (near, 14-byte abs jmp) -> cave
            # Each trampoline is a small allocation near its hook that
            # contains an absolute jump to the far cave.
            for hook_addr, cave_addr in hooks:
                tramp = self._alloc_near(handle, hook_addr, 64)
                if not tramp:
                    raise RuntimeError(
                        f"Could not allocate memory. Close other programs that hook into the game.\n")
                self._trampolines.append(tramp)
                # Write abs jmp to cave in trampoline
                abs_jmp = self._abs_jmp(cave_addr)
                self.pm.write_bytes(tramp, abs_jmp, len(abs_jmp))
                # Patch hook site with rel32 jmp to trampoline
                self.orig_bytes[hook_addr] = self.pm.read_bytes(hook_addr, 7)
                patch = self._jmp_patch(hook_addr, tramp)
                self.pm.write_bytes(hook_addr, patch, 7)

        self.hooks_installed = True

    def uninstall_hooks(self):
        if not self.hooks_installed:
            return
        for addr, orig in self.orig_bytes.items():
            try:
                self.pm.write_bytes(addr, orig, len(orig))
            except Exception:
                pass
        self.hooks_installed = False

    # ── read helpers ─────────────────────────────────────────────────

    def get_player_pos(self):
        """Return (x, y, z) local position, or None."""
        if not self.td:
            return None
        try:
            raw = self.pm.read_bytes(self.td + 0x20, 12)
            x, y, z = struct.unpack('<fff', raw)
            if x == 0.0 and y == 0.0 and z == 0.0:
                return None
            return x, y, z
        except Exception:
            return None

    def get_entity_base(self):
        if not self.td:
            return 0
        try:
            return self.pm.read_ulonglong(self.td + 0x18)
        except Exception:
            return 0

    def get_world_offsets(self):
        """Return (ox, oy, oz, ow) or None."""
        if not self.world_offset_addr:
            return None
        try:
            raw = self.pm.read_bytes(self.world_offset_addr, 16)
            return struct.unpack('<ffff', raw)
        except Exception:
            return None

    def get_player_abs(self):
        """Return absolute world position (x, y, z), or None."""
        pos = self.get_player_pos()
        if not pos:
            return None
        off = self.get_world_offsets()
        if off:
            return pos[0] + off[0], pos[1], pos[2] + off[2]
        return pos

    def get_map_dest(self):
        """Return (x, y, z) map destination, or None if not set."""
        if not self.md:
            return None
        try:
            raw = self.pm.read_bytes(self.md, 16)
            x, y, z, flag = struct.unpack('<fffI', raw)
            if flag != 1:
                return None
            return x, y, z
        except Exception:
            return None

    # ── teleport ─────────────────────────────────────────────────────

    def teleport_to_abs(self, abs_x, abs_y, abs_z):
        """Teleport player to absolute world coordinates."""
        entity = self.get_entity_base()
        if not entity:
            return False, "Player entity not captured.\nMove around in-game and try again."

        off = self.get_world_offsets()
        if off:
            lx = abs_x - off[0]
            ly = abs_y
            lz = abs_z - off[2]
        else:
            lx, ly, lz = abs_x, abs_y, abs_z

        try:
            # Write to position field (+0x90)
            for base_off in [0x90, 0x1A0]:
                self.pm.write_float(entity + base_off, lx)
                self.pm.write_float(entity + base_off + 4, ly)
                self.pm.write_float(entity + base_off + 8, lz)
            return True, ""
        except Exception as e:
            return False, str(e)

    def set_invuln(self, on):
        if self.inv:
            try:
                self.pm.write_bytes(self.inv, b'\x01' if on else b'\x00', 1)
            except Exception:
                pass


# ── WaypointStore ────────────────────────────────────────────────────

class WaypointStore:
    def __init__(self):
        self.local: list[dict] = []
        self.shared: list[dict] = []
        os.makedirs(SAVE_DIR, exist_ok=True)

    def load(self):
        if not os.path.exists(SAVE_FILE):
            return
        try:
            with open(SAVE_FILE, 'r', encoding='utf-8') as f:
                self.local = json.load(f)
        except Exception:
            self.local = []

    def save(self):
        with open(SAVE_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.local, f, indent=2, ensure_ascii=False)

    def add(self, name, x, y, z):
        self.local.append({"name": name, "absX": x, "absY": y, "absZ": z})
        self.save()

    def delete(self, index):
        if 0 <= index < len(self.local):
            self.local.pop(index)
            self.save()

    def rename(self, index, new_name):
        if 0 <= index < len(self.local):
            self.local[index]["name"] = new_name
            self.save()

    def update_coords(self, index, x, y, z):
        if 0 <= index < len(self.local):
            self.local[index].update(absX=x, absY=y, absZ=z)
            self.save()

    def swap(self, i, j):
        if 0 <= i < len(self.local) and 0 <= j < len(self.local):
            self.local[i], self.local[j] = self.local[j], self.local[i]
            self.save()

    def fetch_shared(self):
        try:
            # Cache-bust to ensure fresh data from Google Sheets
            url = f"{SHARED_CSV_URL}&_t={int(time.time())}"
            req = __import__('urllib.request', fromlist=['Request']).Request(
                url, headers={'Cache-Control': 'no-cache', 'Pragma': 'no-cache'})
            resp = urlopen(req, timeout=10)
            content = resp.read().decode('utf-8', errors='replace')
        except Exception as e:
            return False, f"Could not fetch shared waypoints.\n\n{e}"

        self.shared = []
        first = True
        for line in content.splitlines():
            if first:
                first = False
                continue
            # Try quoted name: timestamp,"name",x,y,z
            import re
            m = re.match(r'^[^,]*,"([^"]*)",([^,]+),([^,]+),([^,]+)', line)
            if not m:
                m = re.match(r'^[^,]*,([^,]+),([^,]+),([^,]+),([^,]+)', line)
            if m:
                try:
                    self.shared.append({
                        "name": m.group(1),
                        "absX": float(m.group(2)),
                        "absY": float(m.group(3)),
                        "absZ": float(m.group(4)),
                    })
                except ValueError:
                    pass

        if not self.shared:
            return False, "No waypoints found in the spreadsheet."
        return True, f"Loaded {len(self.shared)} community waypoints."

    def submit(self, name, x, y, z):
        try:
            url = (
                f"{FORM_SUBMIT_URL}"
                f"?{FORM_FIELDS['name']}={quote_plus(name)}"
                f"&{FORM_FIELDS['x']}={x:.6f}"
                f"&{FORM_FIELDS['y']}={y:.6f}"
                f"&{FORM_FIELDS['z']}={z:.6f}"
                f"&submit=Submit"
            )
            urlopen(url, timeout=10)
            return True
        except Exception:
            return False


# ── GUI ──────────────────────────────────────────────────────────────

class TeleporterApp(tk.Tk):
    # Dark theme palette
    BG       = '#1e1e2e'
    BG_ALT   = '#252536'
    BG_INPUT = '#313244'
    BG_CARD  = '#2a2a3c'
    FG       = '#cdd6f4'
    FG_DIM   = '#6c7086'
    ACCENT   = '#89b4fa'
    OK_CLR   = '#a6e3a1'
    ERR_CLR  = '#f38ba8'
    WARN_CLR = '#f9e2af'
    BORDER   = '#45475a'
    SURFACE  = '#363649'
    SEL_BG   = '#5b7bb5'

    HOTKEY_DEFS = [
        ("teleport", "Teleport to Map Marker"),
        ("save",     "Save Map Marker as Waypoint"),
        ("abort",    "Return to Pre-Teleport Position"),
    ]
    HOTKEY_ACTIONS = {
        "teleport": "_teleport_to_map",
        "save":     "_save_map_marker",
        "abort":    "_abort_teleport",
    }

    def __init__(self):
        super().__init__()
        self.title("Crimson Desert Teleporter Tool")
        self.geometry("760x820")
        self.minsize(700, 720)
        self.configure(bg=self.BG)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Window icon + dark title bar
        self._set_icon()
        self._set_dark_titlebar()

        # Center on screen
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x = (sw - 760) // 2
        y = (sh - 820) // 2
        self.geometry(f"760x820+{x}+{y}")

        self.engine = TeleportEngine()
        self.waypoints = WaypointStore()
        self.waypoints.load()

        self.recovery_pos = None
        self.invuln_end = 0
        self._key_state = {}
        self._local_filter_map = []
        self._shared_filter_map = []
        self._rebinding = None  # which hotkey id is being rebound

        # Load hotkey settings
        self._hotkeys = {}
        settings = _load_settings()
        for hk_id, _desc in self.HOTKEY_DEFS:
            saved = settings.get("hotkeys", {}).get(hk_id)
            if saved and "vk" in saved:
                self._hotkeys[hk_id] = {"vk": saved["vk"],
                                        "mod": saved.get("mod", 0),
                                        "enabled": saved.get("enabled", True)}
            else:
                self._hotkeys[hk_id] = dict(DEFAULT_HOTKEYS[hk_id])

        self._adv_height = 1200.0  # default, overwritten by _load_adv_settings

        # Map window state
        self._map_dest = None        # (abs_x, abs_y, abs_z) clicked destination
        self._map_height = 1200.0
        self._map_visible = False
        self._map_webview = None
        self._map_bridge = None

        self._apply_style()
        self._build_ui()

        atexit.register(self._cleanup)
        self.after(100, self._auto_attach)
        self.after(50, self._poll)
        self.after(500, self._fetch_shared_silent)

    def _save_hotkey_settings(self):
        settings = _load_settings()
        settings["hotkeys"] = self._hotkeys
        _save_settings(settings)

    # ── Icon ─────────────────────────────────────────────────────────

    def _set_icon(self):
        """Generate a 32x32 teleport-arrow icon in-memory."""
        sz = 32
        img = tk.PhotoImage(width=sz, height=sz)
        # Colors
        bg = self.BG
        accent = self.ACCENT
        sel = self.SEL_BG
        dim = '#3a3a5c'

        # Draw pixel by pixel — upward arrow with ring
        # Background fill
        img.put(bg, to=(0, 0, sz, sz))

        # Outer ring (circle outline r=14, center 16,16)
        cx, cy, r_out, r_in = 16, 16, 14, 11
        for y in range(sz):
            for x in range(sz):
                d2 = (x - cx) ** 2 + (y - cy) ** 2
                if r_in * r_in <= d2 <= r_out * r_out:
                    img.put(dim, to=(x, y, x + 1, y + 1))

        # Arrow body (vertical bar) — center column, narrowed
        for y in range(8, 25):
            for x in range(14, 19):
                img.put(sel, to=(x, y, x + 1, y + 1))

        # Arrow head (triangle pointing up)
        tip_y = 5
        for row in range(8):
            y = tip_y + row
            half = row
            for x in range(cx - half, cx + half + 1):
                if 0 <= x < sz:
                    img.put(accent, to=(x, y, x + 1, y + 1))

        # Arrow tail flare
        for row in range(3):
            y = 24 + row
            half = 3 + row
            for x in range(cx - half, cx + half + 1):
                if 0 <= x < sz:
                    img.put(sel, to=(x, y, x + 1, y + 1))

        self._icon_img = img  # prevent GC
        self.iconphoto(True, img)

    def _set_dark_titlebar(self):
        """Use Windows DWM API to set dark title bar + custom caption color."""
        # Must be called after the window is mapped so the HWND exists
        self.update_idletasks()
        try:
            # Get the actual top-level HWND via wm frame
            hwnd = int(self.wm_frame(), 16)
            dwm = ctypes.windll.dwmapi
            dwm.DwmSetWindowAttribute.argtypes = [
                ctypes.c_void_p, ctypes.c_ulong,
                ctypes.c_void_p, ctypes.c_ulong,
            ]
            # DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            val = ctypes.c_int(1)
            dwm.DwmSetWindowAttribute(
                hwnd, 20, ctypes.byref(val), ctypes.sizeof(val))
            # DWMWA_CAPTION_COLOR = 35 (Windows 11)
            r = int(self.BG[1:3], 16)
            g = int(self.BG[3:5], 16)
            b = int(self.BG[5:7], 16)
            color = ctypes.c_uint(r | (g << 8) | (b << 16))
            dwm.DwmSetWindowAttribute(
                hwnd, 35, ctypes.byref(color), ctypes.sizeof(color))
        except Exception:
            pass

    # ── Styling ──────────────────────────────────────────────────────

    def _apply_style(self):
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('.', background=self.BG, foreground=self.FG,
                        borderwidth=0, focuscolor=self.ACCENT)
        style.configure('TFrame', background=self.BG)
        style.configure('Card.TFrame', background=self.BG_CARD)
        style.configure('TLabel', background=self.BG, foreground=self.FG)
        style.configure('Card.TLabel', background=self.BG_CARD, foreground=self.FG)
        style.configure('TLabelframe', background=self.BG, foreground=self.FG,
                        bordercolor=self.BORDER)
        style.configure('TLabelframe.Label', background=self.BG, foreground=self.FG)
        style.configure('TButton', background=self.BG_INPUT, foreground=self.FG,
                        borderwidth=1, bordercolor=self.BORDER,
                        lightcolor=self.BORDER, darkcolor=self.BORDER,
                        padding=(10, 5), font=('Segoe UI', 9))
        style.map('TButton',
                  background=[('disabled', self.BG), ('active', self.BORDER),
                              ('pressed', self.ACCENT)],
                  foreground=[('disabled', self.FG_DIM)],
                  lightcolor=[('active', self.BORDER), ('pressed', self.ACCENT)],
                  darkcolor=[('active', self.BORDER), ('pressed', self.ACCENT)])
        style.configure('Action.TButton', background=self.SURFACE,
                        foreground=self.FG, padding=(12, 5),
                        bordercolor=self.BORDER, lightcolor=self.BORDER,
                        darkcolor=self.BORDER, font=('Segoe UI', 9))
        style.map('Action.TButton',
                  background=[('disabled', self.BG), ('active', self.ACCENT),
                              ('pressed', self.SEL_BG)],
                  foreground=[('disabled', self.FG_DIM), ('active', self.BG)],
                  lightcolor=[('active', self.ACCENT)],
                  darkcolor=[('active', self.ACCENT)])
        style.configure('Small.TButton', background=self.BG_INPUT,
                        foreground=self.FG, padding=(6, 2),
                        bordercolor=self.BORDER, lightcolor=self.BORDER,
                        darkcolor=self.BORDER, font=('Segoe UI', 8))
        style.map('Small.TButton',
                  background=[('active', self.ACCENT), ('pressed', self.SEL_BG)],
                  foreground=[('active', self.BG)],
                  lightcolor=[('active', self.ACCENT)],
                  darkcolor=[('active', self.ACCENT)])
        style.configure('SmallDim.TButton', background=self.BG_INPUT,
                        foreground=self.FG_DIM, padding=(6, 2),
                        bordercolor=self.BORDER, lightcolor=self.BORDER,
                        darkcolor=self.BORDER, font=('Segoe UI', 8))
        style.map('SmallDim.TButton',
                  background=[('active', self.BG_INPUT), ('pressed', self.BG_INPUT)],
                  foreground=[('active', self.FG_DIM)],
                  lightcolor=[('active', self.BORDER)],
                  darkcolor=[('active', self.BORDER)])
        style.configure('Arrow.TButton', padding=(2, 6), font=('Segoe UI', 11),
                        bordercolor=self.BORDER, lightcolor=self.BORDER,
                        darkcolor=self.BORDER)
        style.configure('TEntry', fieldbackground=self.BG_INPUT, foreground=self.FG,
                        insertcolor=self.FG, borderwidth=1,
                        bordercolor=self.BORDER, lightcolor=self.BORDER,
                        darkcolor=self.BORDER)
        style.map('TEntry',
                  lightcolor=[('focus', self.ACCENT)],
                  darkcolor=[('focus', self.ACCENT)],
                  bordercolor=[('focus', self.ACCENT)])
        style.configure('TCheckbutton', background=self.BG_CARD, foreground=self.FG,
                        indicatorbackground=self.BG_INPUT,
                        indicatorcolor=self.BG_INPUT,
                        upperbordercolor=self.BORDER,
                        lowerbordercolor=self.BORDER,
                        font=('Segoe UI', 9))
        style.map('TCheckbutton',
                  background=[('active', self.BG_CARD)],
                  foreground=[('disabled', self.FG_DIM)],
                  indicatorbackground=[('selected', self.ACCENT),
                                       ('pressed', self.ACCENT)],
                  indicatorcolor=[('selected', self.ACCENT),
                                  ('!selected', self.BG_INPUT)],
                  upperbordercolor=[('active', self.ACCENT)],
                  lowerbordercolor=[('active', self.ACCENT)])
        style.configure('TNotebook', background=self.BG, borderwidth=0,
                        tabmargins=(2, 5, 2, 0),
                        bordercolor=self.BORDER, lightcolor=self.BORDER,
                        darkcolor=self.BORDER)
        style.configure('TNotebook.Tab', background=self.BG_INPUT, foreground=self.FG_DIM,
                        padding=(16, 5), font=('Segoe UI', 9),
                        bordercolor=self.BORDER, lightcolor=self.BORDER,
                        darkcolor=self.BORDER)
        style.map('TNotebook.Tab',
                  background=[('selected', self.ACCENT), ('active', self.BORDER)],
                  foreground=[('selected', self.BG)],
                  padding=[('selected', (16, 6))],
                  font=[('selected', ('Segoe UI', 9, 'bold'))],
                  lightcolor=[('selected', self.ACCENT)],
                  darkcolor=[('selected', self.ACCENT)],
                  bordercolor=[('selected', self.ACCENT)])
        style.configure('Treeview', background=self.BG_ALT, foreground=self.FG,
                        fieldbackground=self.BG_ALT, borderwidth=1,
                        bordercolor=self.BORDER, lightcolor=self.BORDER,
                        darkcolor=self.BORDER,
                        rowheight=26, font=('Consolas', 9))
        style.configure('Treeview.Heading', background=self.BG_INPUT,
                        foreground=self.FG, borderwidth=1, relief='flat',
                        bordercolor=self.BORDER, lightcolor=self.BORDER,
                        darkcolor=self.BORDER, font=('Segoe UI', 9, 'bold'))
        style.map('Treeview.Heading',
                  background=[('active', self.SURFACE)],
                  lightcolor=[('active', self.BORDER)],
                  darkcolor=[('active', self.BORDER)])
        style.map('Treeview',
                  background=[('selected', self.SEL_BG)],
                  foreground=[('selected', '#ffffff')])
        style.configure('Vertical.TScrollbar', background=self.BG_INPUT,
                        troughcolor=self.BG_ALT, borderwidth=0, arrowsize=14,
                        bordercolor=self.BORDER, lightcolor=self.BORDER,
                        darkcolor=self.BORDER, arrowcolor=self.FG_DIM)
        style.map('Vertical.TScrollbar',
                  background=[('active', self.SURFACE), ('pressed', self.ACCENT)],
                  arrowcolor=[('active', self.FG), ('pressed', self.FG)])
        style.configure('Green.TLabel', foreground=self.OK_CLR, background=self.BG)
        style.configure('Red.TLabel', foreground=self.ERR_CLR, background=self.BG)
        style.configure('Dim.TLabel', foreground=self.FG_DIM, background=self.BG)
        style.configure('Title.TLabel', foreground=self.ACCENT, background=self.BG,
                        font=('Segoe UI', 13, 'bold'))
        style.configure('Mono.TLabel', background=self.BG_CARD,
                        foreground=self.FG, font=('Consolas', 9))
        style.configure('Warn.TLabel', foreground=self.WARN_CLR,
                        background=self.BG_CARD, font=('Segoe UI', 9, 'bold'))
        style.configure('KeyBadge.TLabel', background=self.SURFACE,
                        foreground=self.ACCENT, font=('Consolas', 10, 'bold'),
                        padding=(6, 2), borderwidth=1, relief='groove',
                        bordercolor=self.ACCENT, lightcolor=self.SURFACE,
                        darkcolor=self.SURFACE)
        style.configure('KeyBadge.Rebind.TLabel', background=self.ERR_CLR,
                        foreground=self.BG, font=('Consolas', 10, 'bold'),
                        padding=(6, 2), borderwidth=1, relief='groove',
                        bordercolor=self.ERR_CLR, lightcolor=self.ERR_CLR,
                        darkcolor=self.ERR_CLR)
        style.configure('TSeparator', background=self.BORDER)

    # ── Placeholder helper ────────────────────────────────────────────

    def _setup_placeholder(self, entry, var, text):
        """Add greyed-out placeholder text that clears on focus."""
        entry._placeholder_active = False
        entry._placeholder_suppress = False

        def _show():
            if not var.get() and not entry._placeholder_active:
                entry._placeholder_suppress = True
                entry.insert(0, text)
                entry._placeholder_suppress = False
                entry.configure(foreground=self.FG_DIM)
                entry._placeholder_active = True

        def _on_focus_in(_e):
            if entry._placeholder_active:
                entry._placeholder_suppress = True
                entry.delete(0, tk.END)
                entry._placeholder_suppress = False
                entry.configure(foreground=self.FG)
                entry._placeholder_active = False

        def _on_focus_out(_e):
            if not var.get():
                _show()

        entry.bind('<FocusIn>', _on_focus_in)
        entry.bind('<FocusOut>', _on_focus_out)
        _show()

    # ── Build UI ─────────────────────────────────────────────────────

    def _build_ui(self):
        # Main scrollable area via a frame
        main = ttk.Frame(self)
        main.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        # ── Header ───────────────────────────────────────────────────
        hdr = ttk.Frame(main)
        hdr.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(hdr, text="Crimson Desert Teleporter Tool",
                  style='Title.TLabel').pack(side=tk.LEFT)

        # World Map button (only if pywebview available)
        if _HAS_WEBVIEW:
            self._map_btn = ttk.Button(hdr, text="World Map \u25B6",
                                       style='Action.TButton',
                                       command=self._toggle_map)
            self._map_btn.pack(side=tk.RIGHT, padx=(8, 0))

        # Status indicator (dot + text + retry button)
        self._status_frame = ttk.Frame(hdr)
        self._status_frame.pack(side=tk.RIGHT)
        self._status_dot = tk.Canvas(self._status_frame, width=10, height=10,
                                     bg=self.BG, highlightthickness=0)
        self._status_dot.pack(side=tk.LEFT, padx=(0, 5), pady=3)
        self._status_dot.create_oval(1, 1, 9, 9, fill=self.ERR_CLR,
                                     outline='', tags='dot')
        self.status_var = tk.StringVar(value="Connecting...")
        self._status_lbl = ttk.Label(self._status_frame,
                                      textvariable=self.status_var,
                                      font=('Segoe UI', 9))
        self._status_lbl.pack(side=tk.LEFT)
        self._retry_btn = ttk.Button(self._status_frame, text="Retry",
                                      style='Small.TButton',
                                      command=self._retry_attach)
        # Retry button is hidden initially; shown on error

        ttk.Separator(main, orient='horizontal').pack(fill=tk.X, pady=(0, 8))

        # ── Position card ────────────────────────────────────────────
        pos_card = ttk.Frame(main, style='Card.TFrame')
        pos_card.pack(fill=tk.X, pady=(0, 8), ipady=6)

        pos_inner = ttk.Frame(pos_card, style='Card.TFrame')
        pos_inner.pack(fill=tk.X, padx=10, pady=(6, 2))

        # Local/World coords row with Save Position button on the right
        coords_row = ttk.Frame(pos_inner, style='Card.TFrame')
        coords_row.pack(fill=tk.X, pady=(0, 1))

        coords_labels = ttk.Frame(coords_row, style='Card.TFrame')
        coords_labels.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.pos_var = tk.StringVar(value="Local     ---.--   ---.--   ---.--")
        ttk.Label(coords_labels, textvariable=self.pos_var,
                  style='Mono.TLabel').pack(anchor='w', pady=(0, 1))
        self.abs_var = tk.StringVar(value="World     ---.--   ---.--   ---.--")
        ttk.Label(coords_labels, textvariable=self.abs_var,
                  style='Mono.TLabel').pack(anchor='w', pady=(0, 1))

        ttk.Button(coords_row, text="Save Position", style='Small.TButton',
                   command=self._save_current_pos).pack(side=tk.RIGHT, padx=(6, 0))

        # Map dest row with action buttons
        map_row = ttk.Frame(pos_inner, style='Card.TFrame')
        map_row.pack(fill=tk.X, pady=(1, 0))
        self.map_dest_var = tk.StringVar(value="Map Dest  (no marker)")
        ttk.Label(map_row, textvariable=self.map_dest_var,
                  style='Mono.TLabel').pack(side=tk.LEFT)

        # Map marker action buttons
        map_btn_frame = ttk.Frame(map_row, style='Card.TFrame')
        map_btn_frame.pack(side=tk.RIGHT)
        ttk.Button(map_btn_frame, text="Teleport", style='Small.TButton',
                   command=self._teleport_to_map).pack(side=tk.LEFT, padx=(0, 3))
        self._save_map_btn = ttk.Button(map_btn_frame, text="Save",
                                         style='Small.TButton',
                                         command=self._save_map_marker)
        self._save_map_btn.pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(map_btn_frame, text="Return", style='Small.TButton',
                   command=self._abort_teleport).pack(side=tk.LEFT)

        self.invuln_var = tk.StringVar(value="")
        self._invuln_lbl = ttk.Label(pos_card, textvariable=self.invuln_var,
                                     style='Warn.TLabel')

        # ── Hotkeys card ─────────────────────────────────────────────
        hk_card = ttk.Frame(main, style='Card.TFrame')
        hk_card.pack(fill=tk.X, pady=(0, 8), ipady=4)

        hk_title_row = ttk.Frame(hk_card, style='Card.TFrame')
        hk_title_row.pack(fill=tk.X, padx=10, pady=(6, 4))
        ttk.Label(hk_title_row, text="Hotkeys",
                  style='Card.TLabel',
                  font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT)

        self._hk_widgets = {}
        for hk_id, desc in self.HOTKEY_DEFS:
            row = ttk.Frame(hk_card, style='Card.TFrame')
            row.pack(fill=tk.X, padx=10, pady=2)

            # Toggle checkbox
            var = tk.BooleanVar(value=self._hotkeys[hk_id]["enabled"])
            cb = ttk.Checkbutton(row, variable=var,
                                 command=lambda hid=hk_id: self._toggle_hotkey(hid))
            cb.pack(side=tk.LEFT, padx=(0, 6))

            # Key badge
            vk = self._hotkeys[hk_id]["vk"]
            mod = self._hotkeys[hk_id].get("mod", 0)
            key_name = _hotkey_display(vk, mod)
            badge_var = tk.StringVar(value=f"  {key_name}  ")
            badge = ttk.Label(row, textvariable=badge_var, style='KeyBadge.TLabel')
            badge.pack(side=tk.LEFT, padx=(0, 8))
            badge.bind('<Button-1>', lambda e, hid=hk_id: (
                self._cancel_rebind() if self._rebinding == hid
                else self._start_rebind(hid)))

            # Description
            ttk.Label(row, text=desc, style='Card.TLabel',
                      font=('Segoe UI', 9)).pack(side=tk.LEFT, fill=tk.X)

            self._hk_widgets[hk_id] = {
                "var": var, "badge_var": badge_var, "badge": badge, "cb": cb,
            }

        # ── Notebook (tabs) ──────────────────────────────────────────
        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 4))

        # ── LOCAL tab ────────────────────────────────────────────────
        self.pan_local = ttk.Frame(self.notebook)
        self.notebook.add(self.pan_local, text="  Local Waypoints  ")

        # List + arrows (container for search + tree + arrows)
        list_row = ttk.Frame(self.pan_local)
        list_row.pack(fill=tk.BOTH, expand=True, padx=8, pady=(8, 0))

        # Left side: search + tree
        lf = ttk.Frame(list_row)
        lf.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Search bar inside lf so it matches tree width
        self.local_search_var = tk.StringVar()
        self._local_search = ttk.Entry(lf, textvariable=self.local_search_var)
        self._local_search.pack(fill=tk.X, pady=(0, 4))
        self._setup_placeholder(self._local_search, self.local_search_var, "Search...")
        self.local_search_var.trace_add('write', lambda *_:
            None if self._local_search._placeholder_suppress
            else self._refresh_local_list())
        cols = ("idx", "name", "coords")
        self.local_tree = ttk.Treeview(lf, columns=cols, show='headings',
                                       selectmode='extended')
        self.local_tree.heading("idx", text="#")
        self.local_tree.heading("name", text="Name")
        self.local_tree.heading("coords", text="Coordinates")
        self.local_tree.column("idx", width=35, anchor='e', stretch=False)
        self.local_tree.column("name", width=240)
        self.local_tree.column("coords", width=300)
        lscroll = ttk.Scrollbar(lf, orient=tk.VERTICAL,
                                command=self.local_tree.yview)
        self.local_tree.configure(yscrollcommand=lscroll.set)
        self.local_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        lscroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Arrow column
        arrow_frame = ttk.Frame(list_row)
        arrow_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(4, 0))
        # Spacer to center arrows vertically
        ttk.Frame(arrow_frame).pack(fill=tk.Y, expand=True)
        ttk.Button(arrow_frame, text="\u25B2", width=3, style='Arrow.TButton',
                   command=lambda: self._move(-1)).pack(pady=(0, 4))
        ttk.Button(arrow_frame, text="\u25BC", width=3, style='Arrow.TButton',
                   command=lambda: self._move(1)).pack()
        ttk.Frame(arrow_frame).pack(fill=tk.Y, expand=True)

        # Bindings
        self.local_tree.bind('<Delete>', lambda _: self._delete_selected())
        self.local_tree.bind('<Double-1>', lambda _: self._teleport_selected())
        self.local_tree.bind('<Button-3>', self._local_context_menu)
        self.local_tree.bind('<ButtonPress-1>', self._drag_start)
        self.local_tree.bind('<B1-Motion>', self._drag_motion)
        self.local_tree.bind('<ButtonRelease-1>', self._drag_end)

        # Action buttons
        btn_bar = ttk.Frame(self.pan_local)
        btn_bar.pack(fill=tk.X, padx=8, pady=(6, 8))
        ttk.Button(btn_bar, text="Teleport", style='Action.TButton',
                   command=self._teleport_selected).pack(side=tk.LEFT, padx=(0, 4))
        # Visual separator
        ttk.Frame(btn_bar, width=8).pack(side=tk.LEFT)
        for text, cmd in [
            ("Rename",         self._rename_selected),
            ("Update Coords",  self._update_coords),
            ("Delete",         self._delete_selected),
            ("Contribute",     self._contribute),
        ]:
            ttk.Button(btn_bar, text=text, style='Action.TButton',
                       command=cmd).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(btn_bar, text="Open Folder", style='Small.TButton',
                   command=lambda: os.startfile(SAVE_DIR)).pack(side=tk.RIGHT)

        self._refresh_local_list()

        # ── SHARED tab ───────────────────────────────────────────────
        self.pan_shared = ttk.Frame(self.notebook)
        self.notebook.add(self.pan_shared, text="  Community Waypoints  ")

        shared_list_row = ttk.Frame(self.pan_shared)
        shared_list_row.pack(fill=tk.BOTH, expand=True, padx=8, pady=(8, 0))

        slf = ttk.Frame(shared_list_row)
        slf.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Search bar inside slf so it matches tree width
        self.shared_search_var = tk.StringVar()
        self._shared_search = ttk.Entry(slf, textvariable=self.shared_search_var)
        self._shared_search.pack(fill=tk.X, pady=(0, 4))
        self._setup_placeholder(self._shared_search, self.shared_search_var, "Search...")
        self.shared_search_var.trace_add('write', lambda *_:
            None if self._shared_search._placeholder_suppress
            else self._refresh_shared_list())
        self.shared_tree = ttk.Treeview(slf, columns=cols, show='headings',
                                        selectmode='extended')
        self.shared_tree.heading("idx", text="#")
        self.shared_tree.heading("name", text="Name")
        self.shared_tree.heading("coords", text="Coordinates")
        self.shared_tree.column("idx", width=35, anchor='e', stretch=False)
        self.shared_tree.column("name", width=240)
        self.shared_tree.column("coords", width=300)
        sscroll = ttk.Scrollbar(slf, orient=tk.VERTICAL,
                                command=self.shared_tree.yview)
        self.shared_tree.configure(yscrollcommand=sscroll.set)
        self.shared_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.shared_tree.bind('<Double-1>', lambda _: self._teleport_shared())
        self.shared_tree.bind('<Button-3>', self._shared_context_menu)

        # Arrow column for shared list
        shared_arrow = ttk.Frame(shared_list_row)
        shared_arrow.pack(side=tk.LEFT, fill=tk.Y, padx=(4, 0))
        ttk.Frame(shared_arrow).pack(fill=tk.Y, expand=True)
        ttk.Button(shared_arrow, text="\u25B2", width=3, style='Arrow.TButton',
                   command=lambda: self._shared_move(-1)).pack(pady=(0, 4))
        ttk.Button(shared_arrow, text="\u25BC", width=3, style='Arrow.TButton',
                   command=lambda: self._shared_move(1)).pack()
        ttk.Frame(shared_arrow).pack(fill=tk.Y, expand=True)

        self.shared_status_var = tk.StringVar(
            value="Click 'Refresh' to load community waypoints.")
        ttk.Label(self.pan_shared, textvariable=self.shared_status_var,
                  style='Dim.TLabel').pack(anchor='w', padx=8, pady=4)

        sb = ttk.Frame(self.pan_shared)
        sb.pack(fill=tk.X, padx=8, pady=(2, 8))
        ttk.Button(sb, text="Teleport", style='Action.TButton',
                   command=self._teleport_shared).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Frame(sb, width=8).pack(side=tk.LEFT)
        for text, cmd in [
            ("Refresh",       self._fetch_shared),
            ("Copy to Local", self._copy_to_local),
            ("Copy All",      self._copy_all),
        ]:
            ttk.Button(sb, text=text, style='Action.TButton',
                       command=cmd).pack(side=tk.LEFT, padx=(0, 4))

        # ── ADVANCED tab ─────────────────────────────────────────────
        self.pan_advanced = ttk.Frame(self.notebook)
        self.notebook.add(self.pan_advanced, text="  Advanced  ")

        adv_inner = ttk.Frame(self.pan_advanced, style='Card.TFrame')
        adv_inner.pack(fill=tk.X, padx=8, pady=8, ipady=8)

        # Enable Advanced Mode toggle
        self._adv_enabled_var = tk.BooleanVar(value=False)
        enable_row = ttk.Frame(adv_inner, style='Card.TFrame')
        enable_row.pack(fill=tk.X, padx=12, pady=(8, 4))
        ttk.Checkbutton(enable_row, text="Enable Advanced Options",
                         variable=self._adv_enabled_var,
                         command=self._on_adv_toggle).pack(side=tk.LEFT)
        self._adv_status_var = tk.StringVar(value="  Disabled")
        self._adv_status_lbl = ttk.Label(enable_row,
                                          textvariable=self._adv_status_var,
                                          style='Red.TLabel',
                                          font=('Segoe UI', 9, 'bold'))
        self._adv_status_lbl.pack(side=tk.LEFT, padx=(8, 0))

        ttk.Separator(adv_inner, orient='horizontal').pack(
            fill=tk.X, padx=12, pady=8)

        # Height Settings
        height_title = ttk.Label(adv_inner, text="Teleport Height",
                                  style='Card.TLabel',
                                  font=('Segoe UI', 10, 'bold'))
        height_title.pack(anchor='w', padx=12, pady=(0, 2))
        ttk.Label(adv_inner,
                  text="Height used when map destination Y is 0 (no height data).",
                  style='Card.TLabel',
                  font=('Segoe UI', 8)).pack(anchor='w', padx=12, pady=(0, 8))

        # Preset buttons row
        preset_row = ttk.Frame(adv_inner, style='Card.TFrame')
        preset_row.pack(fill=tk.X, padx=12, pady=(0, 6))

        self._height_var = tk.StringVar(value="1200.0")
        ttk.Button(preset_row, text="Ground (1200)", style='Action.TButton',
                   command=lambda: self._set_adv_height(1200.0)).pack(
            side=tk.LEFT, padx=(0, 6))
        ttk.Button(preset_row, text="Abyss (2400)", style='Action.TButton',
                   command=lambda: self._set_adv_height(2400.0)).pack(
            side=tk.LEFT, padx=(0, 6))

        # Custom height input
        custom_row = ttk.Frame(adv_inner, style='Card.TFrame')
        custom_row.pack(fill=tk.X, padx=12, pady=(0, 8))
        ttk.Label(custom_row, text="Custom Height:",
                  style='Card.TLabel',
                  font=('Segoe UI', 9)).pack(side=tk.LEFT, padx=(0, 6))
        self._height_entry = ttk.Entry(custom_row,
                                        textvariable=self._height_var,
                                        width=12)
        self._height_entry.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(custom_row, text="Apply", style='Small.TButton',
                   command=self._apply_custom_height).pack(side=tk.LEFT)

        # Current height display
        self._adv_height_display = tk.StringVar(value="Current: 1200.0")
        ttk.Label(adv_inner, textvariable=self._adv_height_display,
                  style='Card.TLabel',
                  font=('Consolas', 9, 'bold')).pack(
            anchor='w', padx=12, pady=(0, 8))

        ttk.Separator(adv_inner, orient='horizontal').pack(
            fill=tk.X, padx=12, pady=4)

        # Override height toggle
        self._height_override_var = tk.BooleanVar(value=False)
        override_row = ttk.Frame(adv_inner, style='Card.TFrame')
        override_row.pack(fill=tk.X, padx=12, pady=(8, 4))
        ttk.Checkbutton(override_row,
                         text="Always override map destination height",
                         variable=self._height_override_var,
                         command=self._save_adv_settings).pack(side=tk.LEFT)
        ttk.Label(adv_inner,
                  text="When enabled, ALL teleports to map marker use the "
                       "height above\ninstead of the marker's actual Y coordinate.",
                  style='Card.TLabel',
                  font=('Segoe UI', 8)).pack(anchor='w', padx=12, pady=(0, 8))

        ttk.Separator(adv_inner, orient='horizontal').pack(
            fill=tk.X, padx=12, pady=4)

        # Always start in advanced mode
        self._adv_autostart_var = tk.BooleanVar(value=False)
        auto_row = ttk.Frame(adv_inner, style='Card.TFrame')
        auto_row.pack(fill=tk.X, padx=12, pady=(8, 4))
        self._adv_autostart_cb = ttk.Checkbutton(
            auto_row, text="Always start in Advanced mode",
            variable=self._adv_autostart_var,
            command=self._on_autostart_toggle)
        self._adv_autostart_cb.pack(side=tk.LEFT)

        # Load saved advanced settings and apply
        self._load_adv_settings()


        # ── Bottom status bar ────────────────────────────────────────
        status_bar = ttk.Frame(self, style='Card.TFrame')
        status_bar.pack(fill=tk.X, side=tk.BOTTOM, ipady=4)
        self.bottom_var = tk.StringVar(value="Ready")
        ttk.Label(status_bar, textvariable=self.bottom_var,
                  style='Card.TLabel', font=('Segoe UI', 8)).pack(
            side=tk.LEFT, padx=12)
        ttk.Label(status_bar, text=f"v{VERSION}",
                  style='Dim.TLabel', font=('Segoe UI', 8)).pack(
            side=tk.RIGHT, padx=12)

    # ── Hotkey management ────────────────────────────────────────────

    def _toggle_hotkey(self, hk_id):
        w = self._hk_widgets[hk_id]
        self._hotkeys[hk_id]["enabled"] = w["var"].get()
        self._save_hotkey_settings()

    def _start_rebind(self, hk_id):
        if self._rebinding:
            self._cancel_rebind()
        self._rebinding = hk_id
        w = self._hk_widgets[hk_id]
        w["badge_var"].set(" ... ")
        w["badge"].configure(style='KeyBadge.Rebind.TLabel')
        self.bottom_var.set(
            f"Press a key to bind for: {dict(self.HOTKEY_DEFS)[hk_id]}  "
            f"(click badge again to cancel)")

    def _cancel_rebind(self):
        if not self._rebinding:
            return
        hk_id = self._rebinding
        self._rebinding = None
        w = self._hk_widgets[hk_id]
        vk = self._hotkeys[hk_id]["vk"]
        mod = self._hotkeys[hk_id].get("mod", 0)
        key_name = _hotkey_display(vk, mod)
        w["badge_var"].set(f"  {key_name}  ")
        w["badge"].configure(style='KeyBadge.TLabel')
        self.bottom_var.set("Rebind cancelled.")

    def _poll_rebind(self):
        """Check if a valid key was pressed during rebind mode.
        Supports modifier+key combos (Ctrl/Alt/Shift + any key).
        If only a modifier is held, waits for a primary key.
        If a primary key is pressed alone, binds without modifier.
        """
        if not self._rebinding:
            return
        get_key = ctypes.windll.user32.GetAsyncKeyState

        # Detect which modifier (if any) is held
        active_mod = 0
        for mvk in MOD_VKS:
            if get_key(mvk) & 0x8000:
                active_mod = mvk
                break  # only one modifier supported at a time

        # Check if a primary (non-modifier) key is pressed
        for vk in VK_NAMES:
            if get_key(vk) & 0x8000:
                mod = active_mod
                display = _hotkey_display(vk, mod)

                # Check for conflicts (same vk AND same mod)
                conflict = None
                for other_id, cfg in self._hotkeys.items():
                    if (other_id != self._rebinding
                            and cfg["vk"] == vk
                            and cfg.get("mod", 0) == mod):
                        conflict = other_id
                        break
                if conflict:
                    desc = dict(self.HOTKEY_DEFS).get(conflict, conflict)
                    self.bottom_var.set(
                        f"{display} is already bound to '{desc}'. Pick another key.")
                    return

                hk_id = self._rebinding
                self._rebinding = None
                self._hotkeys[hk_id]["vk"] = vk
                self._hotkeys[hk_id]["mod"] = mod
                self._save_hotkey_settings()

                w = self._hk_widgets[hk_id]
                w["badge_var"].set(f"  {display}  ")
                w["badge"].configure(style='KeyBadge.TLabel')
                self.bottom_var.set(
                    f"Bound {display} to: {dict(self.HOTKEY_DEFS)[hk_id]}")
                self._key_state[hk_id] = True
                return

    # ── Context menu ─────────────────────────────────────────────────

    def _make_menu(self):
        return tk.Menu(self, tearoff=0,
                       bg=self.BG_INPUT, fg=self.FG,
                       activebackground=self.ACCENT, activeforeground=self.BG,
                       disabledforeground=self.FG_DIM,
                       borderwidth=1, relief='solid',
                       font=('Segoe UI', 9))

    def _local_context_menu(self, event):
        iid = self.local_tree.identify_row(event.y)

        # Right-clicked on blank space — show "Add Waypoint" only
        if not iid:
            self.local_tree.selection_remove(*self.local_tree.selection())
            menu = self._make_menu()
            menu.add_command(label="Add Waypoint Manually",
                             command=self._add_manual_waypoint)
            menu.tk_popup(event.x_root, event.y_root)
            return

        # Clicked on an item — select it if needed
        sel = self.local_tree.selection()
        if iid not in sel:
            self.local_tree.selection_set(iid)

        menu = self._make_menu()
        indices = self._selected_local_indices()
        if len(indices) > 1:
            menu.add_command(
                label=f"Delete {len(indices)} Selected",
                command=self._delete_selected)
        else:
            menu.add_command(label="Teleport",      command=self._teleport_selected)
            menu.add_separator()
            menu.add_command(label="Rename",         command=self._rename_selected)
            menu.add_command(label="Update Coords",  command=self._update_coords)
            menu.add_command(label="Contribute",     command=self._contribute)
            menu.add_separator()
            menu.add_command(label="Delete",         command=self._delete_selected)
            menu.add_separator()
            menu.add_command(label="Add Waypoint Manually",
                             command=self._add_manual_waypoint)

        menu.tk_popup(event.x_root, event.y_root)

    def _shared_context_menu(self, event):
        iid = self.shared_tree.identify_row(event.y)
        if not iid:
            return

        sel = self.shared_tree.selection()
        if iid not in sel:
            self.shared_tree.selection_set(iid)

        menu = self._make_menu()
        sel = self.shared_tree.selection()
        if len(sel) > 1:
            menu.add_command(label="Copy to Local", command=self._copy_to_local)
        else:
            menu.add_command(label="Teleport", command=self._teleport_shared)
            menu.add_separator()
            menu.add_command(label="Copy to Local", command=self._copy_to_local)

        menu.tk_popup(event.x_root, event.y_root)

    def _add_manual_waypoint(self):
        """Dialog with separate X, Y, Z fields and a name field."""
        dlg = tk.Toplevel(self)
        dlg.title("Add Waypoint")
        dlg.configure(bg=self.BG)
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()

        dlg.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 340) // 2
        py = self.winfo_y() + (self.winfo_height() - 220) // 2
        dlg.geometry(f"340x220+{px}+{py}")

        result: list[dict | None] = [None]

        # Coordinate fields
        coord_frame = tk.Frame(dlg, bg=self.BG)
        coord_frame.pack(fill=tk.X, padx=16, pady=(16, 8))

        entries = {}
        for label in ("X:", "Y:", "Z:"):
            row = tk.Frame(coord_frame, bg=self.BG)
            row.pack(fill=tk.X, pady=2)
            tk.Label(row, text=label, bg=self.BG, fg=self.FG,
                     font=('Consolas', 10, 'bold'), width=3,
                     anchor='e').pack(side=tk.LEFT)
            e = tk.Entry(row, bg=self.BG_INPUT, fg=self.FG,
                         insertbackground=self.FG, relief='flat',
                         font=('Consolas', 10),
                         highlightbackground=self.BORDER,
                         highlightcolor=self.ACCENT, highlightthickness=1)
            e.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))
            e.insert(0, "0.0")
            entries[label[0].lower()] = e

        # Name field
        name_row = tk.Frame(dlg, bg=self.BG)
        name_row.pack(fill=tk.X, padx=16, pady=(0, 8))
        tk.Label(name_row, text="Name:", bg=self.BG, fg=self.FG,
                 font=('Segoe UI', 10)).pack(side=tk.LEFT)
        name_entry = tk.Entry(name_row, bg=self.BG_INPUT, fg=self.FG,
                              insertbackground=self.FG, relief='flat',
                              font=('Segoe UI', 10),
                              highlightbackground=self.BORDER,
                              highlightcolor=self.ACCENT, highlightthickness=1)
        name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

        entries['x'].focus_set()
        entries['x'].select_range(0, tk.END)

        # Buttons
        btn_frame = tk.Frame(dlg, bg=self.BG)
        btn_frame.pack(fill=tk.X, padx=16, pady=(0, 12))

        def ok():
            try:
                x = float(entries['x'].get())
                y = float(entries['y'].get())
                z = float(entries['z'].get())
            except ValueError:
                self.bottom_var.set("Invalid coordinate values.")
                return
            if x == 0 or y == 0 or z == 0:
                self.bottom_var.set("Coordinates cannot be 0.")
                return
            name = name_entry.get().strip()
            if not name:
                name = f"Manual ({x:.0f}, {z:.0f})"
            result[0] = {"name": name, "x": x, "y": y, "z": z}
            dlg.destroy()

        def cancel():
            dlg.destroy()

        self._make_primary_btn(btn_frame, "Add", ok).pack(
            side=tk.RIGHT, padx=(4, 0))
        self._make_secondary_btn(btn_frame, "Cancel", cancel).pack(
            side=tk.RIGHT)

        # Tab between fields, Enter to confirm
        for e in [entries['x'], entries['y'], entries['z'], name_entry]:
            e.bind('<Return>', lambda _: ok())
        dlg.bind('<Escape>', lambda _: cancel())

        dlg.wait_window()
        if result[0]:
            r = result[0]
            self.waypoints.add(r["name"], r["x"], r["y"], r["z"])
            self._refresh_local_list()
            self.bottom_var.set(f"Added: {r['name']}")

    # ── Dialog button helper ────────────────────────────────────────

    def _make_primary_btn(self, parent, text, command, width=10):
        """Accent-colored button with hover effect."""
        btn = tk.Button(parent, text=text, width=width, command=command,
                        bg=self.ACCENT, fg=self.BG,
                        activebackground=self.SEL_BG, activeforeground=self.FG,
                        relief='flat', font=('Segoe UI', 9, 'bold'),
                        cursor='hand2')
        btn.bind('<Enter>', lambda _: btn.configure(bg=self.SEL_BG))
        btn.bind('<Leave>', lambda _: btn.configure(bg=self.ACCENT))
        return btn

    def _make_secondary_btn(self, parent, text, command, width=10):
        """Subdued button with hover effect."""
        btn = tk.Button(parent, text=text, width=width, command=command,
                        bg=self.BG_INPUT, fg=self.FG,
                        activebackground=self.BORDER, activeforeground=self.FG,
                        relief='flat', font=('Segoe UI', 9),
                        cursor='hand2')
        btn.bind('<Enter>', lambda _: btn.configure(bg=self.BORDER))
        btn.bind('<Leave>', lambda _: btn.configure(bg=self.BG_INPUT))
        return btn

    # ── Themed dialogs ───────────────────────────────────────────────

    def _themed_askstring(self, title, prompt, initialvalue=""):
        """Dark-themed replacement for simpledialog.askstring."""
        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.configure(bg=self.BG)
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()

        # Center on parent
        dlg.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 360) // 2
        py = self.winfo_y() + (self.winfo_height() - 150) // 2
        dlg.geometry(f"360x150+{px}+{py}")

        result: list[str | None] = [None]

        tk.Label(dlg, text=prompt, bg=self.BG, fg=self.FG,
                 font=('Segoe UI', 10)).pack(padx=16, pady=(16, 6), anchor='w')

        entry = tk.Entry(dlg, bg=self.BG_INPUT, fg=self.FG,
                         insertbackground=self.FG, relief='solid',
                         borderwidth=1, font=('Segoe UI', 10),
                         highlightbackground=self.BORDER,
                         highlightcolor=self.ACCENT, highlightthickness=1)
        entry.pack(fill=tk.X, padx=16, pady=(0, 12))
        entry.insert(0, initialvalue)
        entry.select_range(0, tk.END)
        entry.focus_set()

        btn_frame = tk.Frame(dlg, bg=self.BG)
        btn_frame.pack(fill=tk.X, padx=16, pady=(0, 12))

        def ok():
            result[0] = entry.get()
            dlg.destroy()
        def cancel():
            dlg.destroy()

        self._make_primary_btn(btn_frame, "OK", ok).pack(
            side=tk.RIGHT, padx=(4, 0))
        self._make_secondary_btn(btn_frame, "Cancel", cancel).pack(
            side=tk.RIGHT)

        entry.bind('<Return>', lambda _: ok())
        entry.bind('<Escape>', lambda _: cancel())

        dlg.wait_window()
        return result[0]

    def _themed_msgbox(self, title, message):
        """Dark-themed info message box with a single OK button."""
        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.configure(bg=self.BG)
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()

        dlg.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 420) // 2
        py = self.winfo_y() + (self.winfo_height() - 160) // 2
        dlg.geometry(f"420x160+{px}+{py}")

        tk.Label(dlg, text=message, bg=self.BG, fg=self.FG,
                 font=('Segoe UI', 10), wraplength=380,
                 justify='left').pack(padx=20, pady=(20, 16), anchor='w')

        btn_frame = tk.Frame(dlg, bg=self.BG)
        btn_frame.pack(fill=tk.X, padx=20, pady=(0, 16))

        self._make_primary_btn(btn_frame, "OK", dlg.destroy).pack(
            side=tk.RIGHT)

        dlg.bind('<Return>', lambda _: dlg.destroy())
        dlg.bind('<Escape>', lambda _: dlg.destroy())
        dlg.focus_set()
        dlg.wait_window()

    def _themed_askyesno(self, title, message):
        """Dark-themed replacement for messagebox.askyesno."""
        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.configure(bg=self.BG)
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()

        dlg.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 400) // 2
        py = self.winfo_y() + (self.winfo_height() - 140) // 2
        dlg.geometry(f"400x140+{px}+{py}")

        result = [False]

        tk.Label(dlg, text=message, bg=self.BG, fg=self.FG,
                 font=('Segoe UI', 10), wraplength=360,
                 justify='left').pack(padx=20, pady=(20, 16), anchor='w')

        btn_frame = tk.Frame(dlg, bg=self.BG)
        btn_frame.pack(fill=tk.X, padx=20, pady=(0, 16))

        def yes():
            result[0] = True
            dlg.destroy()
        def no():
            dlg.destroy()

        self._make_primary_btn(btn_frame, "Yes", yes).pack(
            side=tk.RIGHT, padx=(4, 0))
        self._make_secondary_btn(btn_frame, "No", no).pack(
            side=tk.RIGHT)

        dlg.bind('<Return>', lambda _: yes())
        dlg.bind('<Escape>', lambda _: no())
        dlg.focus_set()

        dlg.wait_window()
        return result[0]

    # ── Drag-and-drop reordering ─────────────────────────────────────

    def _drag_start(self, event):
        self._drag_item = self.local_tree.identify_row(event.y) or None

    def _drag_motion(self, event):
        drag = getattr(self, '_drag_item', None)
        if not drag:
            return
        target = self.local_tree.identify_row(event.y)
        if target and target != drag:
            src_tree_idx = self.local_tree.index(drag)
            dst_tree_idx = self.local_tree.index(target)
            if src_tree_idx < len(self._local_filter_map) and \
               dst_tree_idx < len(self._local_filter_map):
                src_real = self._local_filter_map[src_tree_idx]
                dst_real = self._local_filter_map[dst_tree_idx]
                self.waypoints.swap(src_real, dst_real)
                self._local_filter_map[src_tree_idx], \
                    self._local_filter_map[dst_tree_idx] = \
                    self._local_filter_map[dst_tree_idx], \
                    self._local_filter_map[src_tree_idx]
                self.local_tree.move(drag, '', dst_tree_idx)
                self._update_tree_indices()

    def _drag_end(self, _event):
        self._drag_item = None

    def _update_tree_indices(self):
        for tree_idx, iid in enumerate(self.local_tree.get_children()):
            if tree_idx < len(self._local_filter_map):
                real_idx = self._local_filter_map[tree_idx]
            else:
                real_idx = 0
            vals = list(self.local_tree.item(iid, 'values'))
            vals[0] = str(real_idx + 1)
            self.local_tree.item(iid, values=vals)

    # ── Auto-attach ──────────────────────────────────────────────────

    def _set_status(self, text, color, show_retry=False):
        self.status_var.set(text)
        self._status_dot.itemconfigure('dot', fill=color)
        if show_retry:
            self._retry_btn.pack(side=tk.LEFT, padx=(6, 0))
        else:
            self._retry_btn.pack_forget()

    def _is_game_running(self):
        """Check if the attached game process is still alive."""
        try:
            handle = self.engine.pm.process_handle
            exit_code = ctypes.c_ulong()
            ctypes.windll.kernel32.GetExitCodeProcess(
                handle, ctypes.byref(exit_code))
            return exit_code.value == 259  # STILL_ACTIVE
        except Exception:
            return False

    def _handle_disconnect(self):
        """Clean up after game process dies and start auto-reconnect."""
        try:
            self.engine.attached = False
            self.engine.hooks_installed = False
            self.engine.block = 0
            self.engine.orig_bytes.clear()
            if self.engine.pm:
                try:
                    self.engine.pm.close_process()
                except Exception:
                    pass
                self.engine.pm = None
        except Exception:
            pass
        self.invuln_end = 0
        self.invuln_var.set("")
        self._invuln_lbl.pack_forget()
        self.pos_var.set("Local     ---.--   ---.--   ---.--")
        self.abs_var.set("World     ---.--   ---.--   ---.--")
        self.map_dest_var.set("Map Dest  (no marker)")
        self._set_status("Game closed — waiting to reconnect...", self.ERR_CLR)
        self.bottom_var.set("Game process ended. Will reconnect automatically.")
        self.after(3000, self._auto_attach)

    def _auto_attach(self):
        try:
            self.engine.attach()
            self._set_status("Scanning...", self.WARN_CLR)
            self.update()
            self.engine.scan_and_hook()
            self._set_status("Connected", self.OK_CLR)
            self.bottom_var.set(
                "Hooks installed. Move around in-game to capture entity.")
        except pymem.exception.ProcessNotFound:
            self._set_status("Waiting for game...", self.ERR_CLR)
            self.after(3000, self._auto_attach)
        except Exception as e:
            self._set_status("Error", self.ERR_CLR, show_retry=True)
            self.bottom_var.set(
                f"{e}  —  Make sure you are loaded into the game world "
                f"before launching the program.")

    def _retry_attach(self):
        """Clean up failed state and try attaching again."""
        # Reset engine state from any partial attach
        try:
            self.engine.detach()
        except Exception:
            pass
        self.engine = TeleportEngine()
        self._set_status("Retrying...", self.WARN_CLR)
        self.bottom_var.set("Retrying connection...")
        self.update()
        self.after(500, self._auto_attach)

    # ── Polling (hotkeys + position) ─────────────────────────────────

    def _poll(self):
        get_key = ctypes.windll.user32.GetAsyncKeyState

        # Rebind mode takes priority
        if self._rebinding:
            self._poll_rebind()
            self.after(50, self._poll)
            return

        if self.engine.attached and self.engine.hooks_installed:
            # Check if game process is still alive
            if not self._is_game_running():
                self._handle_disconnect()
                self.after(50, self._poll)
                return

            # Update position display
            pos = self.engine.get_player_pos()
            if pos:
                self.pos_var.set(
                    f"Local   {pos[0]:>10.2f}  {pos[1]:>10.2f}  {pos[2]:>10.2f}")
                apos = self.engine.get_player_abs()
                if apos:
                    self.abs_var.set(
                        f"World   {apos[0]:>10.2f}  {apos[1]:>10.2f}  {apos[2]:>10.2f}")
            else:
                self.pos_var.set("Local     ---.--   ---.--   ---.--")
                self.abs_var.set("World     (move in-game to capture)")

            # Update map dest display
            mdest = self.engine.get_map_dest()
            if mdest:
                self.map_dest_var.set(
                    f"Map Dest{mdest[0]:>10.2f}  {mdest[1]:>10.2f}  {mdest[2]:>10.2f}")
            else:
                self.map_dest_var.set("Map Dest  (no marker)")

            # Push player position to webview map (throttled to ~every 500ms)
            if _HAS_WEBVIEW and self._map_visible and self._map_webview:
                now = time.time()
                if now - getattr(self, '_map_last_refresh', 0) > 0.5:
                    self._map_last_refresh = now
                    self._push_map_player()

            # Invulnerability timer
            if self.invuln_end and time.time() >= self.invuln_end:
                self.engine.set_invuln(False)
                self.invuln_end = 0
                self.invuln_var.set("")
                self._invuln_lbl.pack_forget()
            elif self.invuln_end:
                remaining = max(0, self.invuln_end - time.time())
                self.invuln_var.set(f"  Invulnerable  {remaining:.0f}s")
                self._invuln_lbl.pack(anchor='w', padx=10, pady=(2, 4))

            # Hotkeys (edge-triggered, with modifier support)
            for hk_id, cfg in self._hotkeys.items():
                if not cfg["enabled"]:
                    self._key_state[hk_id] = False
                    continue
                key_down = bool(get_key(cfg["vk"]) & 0x8000)
                mod = cfg.get("mod", 0)
                if mod:
                    mod_down = bool(get_key(mod) & 0x8000)
                else:
                    # No modifier required — make sure no modifier is held
                    mod_down = not any(
                        get_key(m) & 0x8000 for m in MOD_VKS)
                pressed = key_down and mod_down
                was_pressed = self._key_state.get(hk_id, False)
                if pressed and not was_pressed:
                    method = getattr(self, self.HOTKEY_ACTIONS[hk_id])
                    method()
                self._key_state[hk_id] = pressed

        self.after(50, self._poll)

    # ── Teleport actions ─────────────────────────────────────────────

    def _trigger_invuln(self):
        self.engine.set_invuln(True)
        self.invuln_end = time.time() + INVULN_SECONDS

    def _teleport_to_map(self):
        dest = self.engine.get_map_dest()
        if not dest:
            self.bottom_var.set(
                "No map marker set. Open the map and place a marker first.")
            return
        x, y, z = dest
        effective_y = self._get_effective_height(y)
        if effective_y is None:
            self.bottom_var.set(
                "Map marker has no height data (Y=0). Enable Advanced "
                "options to set a custom height.")
            return
        apos = self.engine.get_player_abs()
        if apos:
            self.recovery_pos = apos
        effective_y += HEIGHT_BOOST
        ok, err = self.engine.teleport_to_abs(x, effective_y, z)
        if ok:
            self._trigger_invuln()
            self.bottom_var.set(
                f"Teleported to map marker ({x:.0f}, {effective_y:.0f}, {z:.0f})")
        else:
            self.bottom_var.set(err)

    def _save_map_marker(self):
        if self._adv_enabled_var.get():
            self._themed_msgbox(
                "Save Disabled",
                "Map destination saving is disabled in Advanced mode.\n\n"
                "Instead, teleport to the location you want to save and "
                "once you are on the ground, use the Save Position button.")
            return
        dest = self.engine.get_map_dest()
        if not dest:
            self.bottom_var.set("No map marker set.")
            return
        x, y, z = dest
        if y == 0.0:
            self.bottom_var.set(
                "Cannot save marker with no height data (Y=0). "
                "Move there first, then use Save Position, or advanced settings.")
            return
        if z == 0.0:
            self.bottom_var.set("Invalid map marker (Z=0). Place marker in a different spot.")
            return
        default_name = f"Map Marker ({x:.0f}, {z:.0f})"
        result = self._themed_askstring(
            "Save Waypoint", "Waypoint name:", initialvalue=default_name)
        if not result:
            return
        self.waypoints.add(result, x, y, z)
        self._refresh_local_list()
        self.bottom_var.set(f"Saved: {result}")

    def _abort_teleport(self):
        if not self.recovery_pos:
            self.bottom_var.set("No recovery position. Teleport somewhere first.")
            return
        rx, ry, rz = self.recovery_pos
        ok, err = self.engine.teleport_to_abs(rx, ry, rz)
        if ok:
            self.recovery_pos = None
            self.bottom_var.set("Returned to pre-teleport position.")
        else:
            self.bottom_var.set(err)

    # ── Local waypoint actions ───────────────────────────────────────

    def _selected_local_idx(self):
        sel = self.local_tree.selection()
        if not sel:
            return -1
        tree_idx = self.local_tree.index(sel[0])
        if tree_idx < len(self._local_filter_map):
            return self._local_filter_map[tree_idx]
        return -1

    def _selected_local_indices(self):
        indices = []
        for iid in self.local_tree.selection():
            tree_idx = self.local_tree.index(iid)
            if tree_idx < len(self._local_filter_map):
                indices.append(self._local_filter_map[tree_idx])
        return sorted(indices)

    def _refresh_local_list(self, select_real_idx=None):
        self.local_tree.delete(*self.local_tree.get_children())
        self._local_filter_map = []
        filt = "" if getattr(self._local_search, '_placeholder_active', False) \
            else self.local_search_var.get().lower()
        select_iid = None
        for i, loc in enumerate(self.waypoints.local):
            if filt and filt not in loc["name"].lower():
                continue
            self._local_filter_map.append(i)
            iid = self.local_tree.insert('', tk.END, values=(
                i + 1,
                loc["name"],
                f"({loc['absX']:.1f}, {loc['absY']:.1f}, {loc['absZ']:.1f})",
            ))
            if select_real_idx is not None and i == select_real_idx:
                select_iid = iid
        if select_iid:
            self.local_tree.selection_set(select_iid)
            self.local_tree.see(select_iid)
        # Update webview map waypoints
        if _HAS_WEBVIEW and self._map_visible:
            self._push_waypoints()

    def _save_current_pos(self):
        apos = self.engine.get_player_abs()
        if not apos:
            self.bottom_var.set("Cannot read position. Move in-game first.")
            return
        default = f"Position ({apos[0]:.0f}, {apos[2]:.0f})"
        name = self._themed_askstring(
            "Save Current Position", "Waypoint name:", initialvalue=default)
        if not name:
            return
        self.waypoints.add(name, *apos)
        self._refresh_local_list()
        self.bottom_var.set(f"Saved: {name}")

    def _teleport_selected(self):
        idx = self._selected_local_idx()
        if idx < 0:
            self.bottom_var.set("Select a waypoint first.")
            return
        loc = self.waypoints.local[idx]
        apos = self.engine.get_player_abs()
        if apos:
            self.recovery_pos = apos
        ok, err = self.engine.teleport_to_abs(
            loc["absX"], loc["absY"], loc["absZ"])
        if ok:
            self._trigger_invuln()
            self.bottom_var.set(f"Teleported to: {loc['name']}")
        else:
            self.bottom_var.set(err)

    def _delete_selected(self):
        indices = self._selected_local_indices()
        if not indices:
            return
        if len(indices) == 1:
            name = self.waypoints.local[indices[0]]["name"]
            msg = f"Delete '{name}'?"
        else:
            msg = f"Delete {len(indices)} selected waypoints?"
        if self._themed_askyesno("Delete", msg):
            for idx in reversed(indices):
                self.waypoints.delete(idx)
            self._refresh_local_list()

    def _rename_selected(self):
        idx = self._selected_local_idx()
        if idx < 0:
            self.bottom_var.set("Select a waypoint first.")
            return
        old_name = self.waypoints.local[idx]["name"]
        new_name = self._themed_askstring(
            "Rename Waypoint", "New name:", initialvalue=old_name)
        if not new_name or new_name == old_name:
            return
        self.waypoints.rename(idx, new_name)
        self._refresh_local_list()
        self.bottom_var.set(f"Renamed to: {new_name}")

    def _move(self, direction):
        idx = self._selected_local_idx()
        if idx < 0:
            return
        target = idx + direction
        if 0 <= target < len(self.waypoints.local):
            self.waypoints.swap(idx, target)
            self.local_search_var.set("")
            self._refresh_local_list(select_real_idx=target)

    def _update_coords(self):
        idx = self._selected_local_idx()
        if idx < 0:
            return
        apos = self.engine.get_player_abs()
        if not apos:
            self.bottom_var.set("Cannot read position.")
            return
        self.waypoints.update_coords(idx, *apos)
        self._refresh_local_list()
        self.bottom_var.set("Coordinates updated.")

    def _contribute(self):
        idx = self._selected_local_idx()
        if idx < 0:
            self.bottom_var.set("Select a waypoint to contribute.")
            return
        loc = self.waypoints.local[idx]
        if not self._themed_askyesno(
                "Contribute",
                f"Submit '{loc['name']}' to the community list?\n\n"
                f"Coords: ({loc['absX']:.1f}, {loc['absY']:.1f}, "
                f"{loc['absZ']:.1f})"):
            return
        ok = self.waypoints.submit(
            loc["name"], loc["absX"], loc["absY"], loc["absZ"])
        if ok:
            self.bottom_var.set(
                f"'{loc['name']}' submitted to community list!")
        else:
            self.bottom_var.set("Failed to submit. Check internet connection.")

    # ── Shared waypoint actions ──────────────────────────────────────

    def _selected_shared_idx(self):
        sel = self.shared_tree.selection()
        if not sel:
            return -1
        tree_idx = self.shared_tree.index(sel[0])
        if tree_idx < len(self._shared_filter_map):
            return self._shared_filter_map[tree_idx]
        return -1

    def _refresh_shared_list(self):
        self.shared_tree.delete(*self.shared_tree.get_children())
        self._shared_filter_map = []
        filt = "" if getattr(self._shared_search, '_placeholder_active', False) \
            else self.shared_search_var.get().lower()
        for i, loc in enumerate(self.waypoints.shared):
            if filt and filt not in loc["name"].lower():
                continue
            self._shared_filter_map.append(i)
            self.shared_tree.insert('', tk.END, values=(
                i + 1,
                loc["name"],
                f"({loc['absX']:.1f}, {loc['absY']:.1f}, {loc['absZ']:.1f})",
            ))

    def _shared_move(self, direction):
        """Move the selection up/down in the shared list."""
        sel = self.shared_tree.selection()
        children = self.shared_tree.get_children()
        if not children:
            return
        if not sel:
            # Nothing selected — select first or last
            target = children[0] if direction > 0 else children[-1]
        else:
            cur_idx = self.shared_tree.index(sel[0])
            target_idx = cur_idx + direction
            if target_idx < 0 or target_idx >= len(children):
                return
            target = children[target_idx]
        self.shared_tree.selection_set(target)
        self.shared_tree.see(target)

    def _fetch_shared_silent(self):
        """Auto-fetch community waypoints on startup without blocking UI."""
        ok, msg = self.waypoints.fetch_shared()
        if ok:
            self.shared_status_var.set(msg)
            self._refresh_shared_list()

    def _fetch_shared(self):
        self.shared_status_var.set("Fetching...")
        self.update()
        ok, msg = self.waypoints.fetch_shared()
        self.shared_status_var.set(msg)
        if ok:
            self.shared_search_var.set("")
            self._refresh_shared_list()

    def _teleport_shared(self):
        idx = self._selected_shared_idx()
        if idx < 0:
            self.bottom_var.set("Select a shared waypoint first.")
            return
        loc = self.waypoints.shared[idx]
        apos = self.engine.get_player_abs()
        if apos:
            self.recovery_pos = apos
        ok, err = self.engine.teleport_to_abs(
            loc["absX"], loc["absY"], loc["absZ"])
        if ok:
            self._trigger_invuln()
            self.bottom_var.set(f"Teleported to: {loc['name']}")
        else:
            self.bottom_var.set(err)

    def _copy_to_local(self):
        idx = self._selected_shared_idx()
        if idx < 0:
            return
        loc = self.waypoints.shared[idx]
        self.waypoints.add(loc["name"], loc["absX"], loc["absY"], loc["absZ"])
        self._refresh_local_list()
        self.bottom_var.set(f"Copied '{loc['name']}' to Local.")

    def _copy_all(self):
        if not self.waypoints.shared:
            self.bottom_var.set("No shared waypoints loaded.")
            return
        if not self._themed_askyesno(
                "Copy All",
                f"Copy all {len(self.waypoints.shared)} shared waypoints "
                f"to Local?"):
            return
        for loc in self.waypoints.shared:
            self.waypoints.add(
                loc["name"], loc["absX"], loc["absY"], loc["absZ"])
        self._refresh_local_list()
        self.bottom_var.set(f"Copied {len(self.waypoints.shared)} waypoints.")

    # ── World Map (webview overlay on crimsondesert.gg) ─────────────────

    def _toggle_map(self):
        """Open or close the interactive web map window."""
        if self._map_visible:
            if self._map_webview:
                try:
                    self._map_webview.destroy()
                except Exception:
                    pass
            self._map_webview = None
            self._map_visible = False
            self._map_btn.configure(text="World Map \u25B6")
        else:
            self._start_map_window()

    def _start_map_window(self):
        """Launch the pywebview window loading the map site."""
        settings = _load_settings()
        self._map_height = settings.get("map_height", 1200.0)
        map_url = MAP_URL

        bridge = _MapBridge(self)
        self._map_bridge = bridge

        win = webview.create_window(
            "World Map \u2014 Crimson Desert Teleporter",
            url=map_url,
            js_api=bridge,
            width=1350,
            height=1125,
            min_size=(700, 500),
            background_color="#1e1e2e",
        )
        self._map_webview = win
        self._map_visible = True
        self._map_btn.configure(text="World Map \u25C0")

        def _on_closed():
            self._map_webview = None
            self.after(0, self._on_map_closed)

        def _on_loaded():
            try:
                win.evaluate_js(_MAP_INJECT_JS)
                win.evaluate_js(f"setHeight({self._map_height})")
                # Load saved calibration if available
                cal = _load_settings().get("map_cal")
                if cal:
                    win.evaluate_js(
                        f"setCalibration({cal['lz']},{cal['oz']},"
                        f"{cal['lx']},{cal['ox']})")
                self._push_waypoints()
            except Exception:
                pass

        win.events.closed += _on_closed
        win.events.loaded += _on_loaded


    def _on_map_closed(self):
        """Handle the map window being closed by the user."""
        self._map_visible = False
        self._map_webview = None
        try:
            self._map_btn.configure(text="World Map \u25B6")
        except Exception:
            pass

    def _push_map_player(self):
        """Push the current player position to the webview map."""
        if not self._map_webview:
            return
        try:
            if self.engine.attached and self.engine.hooks_installed:
                apos = self.engine.get_player_abs()
                if apos:
                    self._map_webview.evaluate_js(
                        f"updatePlayer({apos[0]},{apos[2]})")
        except Exception:
            pass

    def _push_waypoints(self):
        """Push current waypoints to the webview map."""
        if not self._map_webview:
            return
        try:
            wps = []
            for wp in self.waypoints.local:
                wps.append({
                    "x": wp.get("absX", 0),
                    "z": wp.get("absZ", 0),
                    "name": wp.get("name", ""),
                })
            import json as _json
            self._map_webview.evaluate_js(
                f"updateWaypoints({_json.dumps(wps)})")
        except Exception:
            pass

    def _map_teleport(self):
        """Teleport to the destination set via the web map."""
        if self._map_dest is None:
            self.bottom_var.set("Set a destination on the map first.")
            if self._map_webview:
                try:
                    self._map_webview.evaluate_js(
                        "setStatus('Set a destination on the map first.')")
                except Exception:
                    pass
            return
        if not self.engine.attached or not self.engine.hooks_installed:
            self.bottom_var.set("Not connected to game.")
            return
        x, _, z = self._map_dest
        y = self._map_height
        apos = self.engine.get_player_abs()
        if apos:
            self.recovery_pos = apos
        ok, err = self.engine.teleport_to_abs(x, y + HEIGHT_BOOST, z)
        if ok:
            self._trigger_invuln()
            msg = f"Teleported to ({x:.1f}, {z:.1f}) Height: {y:.0f}"
            self.bottom_var.set(f"Map teleport to ({x:.1f}, {z:.1f})")
            if self._map_webview:
                try:
                    self._map_webview.evaluate_js(f"setStatus('{msg}')")
                except Exception:
                    pass
        else:
            self.bottom_var.set(err)


    # ── Advanced settings ────────────────────────────────────────

    def _load_adv_settings(self):
        """Load advanced settings from JSON and apply to UI."""
        settings = _load_settings()
        adv = settings.get("advanced", {})
        self._adv_height = adv.get("height", 1200.0)
        self._height_var.set(str(self._adv_height))
        self._adv_height_display.set(f"Current: {self._adv_height}")
        self._height_override_var.set(adv.get("height_override", False))
        self._adv_autostart_var.set(adv.get("autostart", False))
        # If autostart is enabled, activate advanced mode
        if adv.get("autostart", False):
            self._adv_enabled_var.set(True)
            self._adv_status_var.set("  Enabled")
            self._adv_status_lbl.configure(style='Green.TLabel')
            self._save_map_btn.configure(style='SmallDim.TButton')

    def _save_adv_settings(self):
        """Persist advanced settings to JSON."""
        settings = _load_settings()
        settings["advanced"] = {
            "height": self._adv_height,
            "height_override": self._height_override_var.get(),
            "autostart": self._adv_autostart_var.get(),
        }
        _save_settings(settings)

    def _on_adv_toggle(self):
        """Toggle advanced mode on/off."""
        enabled = self._adv_enabled_var.get()
        if enabled:
            self._adv_status_var.set("  Enabled")
            self._adv_status_lbl.configure(style='Green.TLabel')
            self._save_map_btn.configure(style='SmallDim.TButton')
            self.bottom_var.set("Advanced options enabled.")
        else:
            self._adv_status_var.set("  Disabled")
            self._adv_status_lbl.configure(style='Red.TLabel')
            self._save_map_btn.configure(style='Small.TButton')
            # Uncheck and disable autostart when advanced is off
            self._adv_autostart_var.set(False)
            self._save_adv_settings()
            self.bottom_var.set("Advanced options disabled.")
        self._adv_autostart_cb.state(
            ['!disabled'] if enabled else ['disabled'])

    def _on_autostart_toggle(self):
        """When autostart is toggled on, also enable advanced mode."""
        if self._adv_autostart_var.get() and not self._adv_enabled_var.get():
            self._adv_enabled_var.set(True)
            self._on_adv_toggle()
        self._save_adv_settings()

    def _set_adv_height(self, value):
        """Set height from a preset button."""
        self._adv_height = value
        self._height_var.set(str(value))
        self._adv_height_display.set(f"Current: {value}")
        self._save_adv_settings()
        self.bottom_var.set(f"Height set to {value}")

    def _apply_custom_height(self):
        """Apply the custom height value from the entry field."""
        try:
            value = float(self._height_var.get())
        except ValueError:
            self.bottom_var.set("Invalid height value. Enter a number.")
            return
        if value <= 0:
            self.bottom_var.set("Height must be greater than 0.")
            return
        self._adv_height = value
        self._adv_height_display.set(f"Current: {value}")
        self._save_adv_settings()
        self.bottom_var.set(f"Custom height set to {value}")

    def _get_effective_height(self, map_y):
        """Determine the Y coordinate to use for teleportation.

        If advanced mode is disabled, returns map_y as-is (or None if 0).
        If advanced mode is enabled:
          - height_override ON: always use advanced height
          - height_override OFF: use advanced height only when map_y is 0
        """
        if not self._adv_enabled_var.get():
            return map_y if map_y != 0.0 else None
        if self._height_override_var.get():
            return self._adv_height
        if map_y == 0.0:
            return self._adv_height
        return map_y


    # ── Cleanup ──────────────────────────────────────────────────────

    def _cleanup(self):
        try:
            self.engine.detach()
        except Exception:
            pass

    def _on_close(self):
        self._cleanup()
        self.destroy()


def main():
    if _HAS_WEBVIEW:
        # pywebview requires the main thread for COM/WebView2 on Windows.
        # Create a hidden sentinel window to keep the webview event loop alive,
        # then run tkinter in a background thread via the func= callback.
        _sentinel = webview.create_window("", hidden=True)

        def _run_tk():
            app = TeleporterApp()
            app.mainloop()
            try:
                _sentinel.destroy()
            except Exception:
                import os
                os._exit(0)

        storage = os.path.join(SAVE_DIR, "webview_data")
        os.makedirs(storage, exist_ok=True)
        webview.start(func=_run_tk, private_mode=True, storage_path=storage)
    else:
        app = TeleporterApp()
        app.mainloop()


if __name__ == "__main__":
    main()
