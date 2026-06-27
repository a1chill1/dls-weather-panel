/* eslint-disable */
// ============================================================================
// DLS Weather Panel - SPFx client-side web part (no framework)
// ----------------------------------------------------------------------------
// Scheduling-focused weather panel for the DossSurveying home page. 100% NOAA /
// National Weather Service data: animated radar loop (RIDGE GIF), active weather
// alerts, next-12h precip-chance timeline, and a 7-day forecast.
//
// v1.0.0.7 - SEARCHABLE LOCATION. A search box in the header (top-right) geocodes
//   any US address OR "City, ST" (keyless OpenStreetMap/Nominatim) and switches the
//   whole panel - radar station, current conditions, 12h timeline, 7-day, alerts and
//   the location label - to that point. The nearest NWS radar station is resolved
//   per-location from api.weather.gov /points (radarStation), so the radar follows
//   the search. A home button resets to the office, and EVERY page reload defaults
//   back to 107 Scottsville Rd, Lafayette, TN 37083 (36.52146, -86.026315).
//   Built on v1.0.0.6 layout (full-bleed 98vw, crisp ~660px radar, full-width 7-day).
//
// Durability: api.weather.gov + radar.weather.gov (US Gov, free, no key) +
// nominatim.openstreetmap.org (free, keyless, low-volume use; CORS-ok from the tenant).
// Bundled, no external script/CDN, no iframe. Every call wrapped -> fails safe, never
// breaks the page. Config (property pane): title, home label, home lat/long, fallback radar.
// ============================================================================

import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

export interface IWeatherPanelWebPartProps {
  title: string;
  locationLabel: string;
  latitude: string;
  longitude: string;
  radarSite: string;
}

const REFRESH_MS = 20 * 60 * 1000; // 20 minutes

export default class WeatherPanelWebPart extends BaseClientSideWebPart<IWeatherPanelWebPartProps> {
  private _timer: any = undefined;
  private _cssInjected = false;
  // current view (resets to home on every render / page load)
  private _lat = '';
  private _lon = '';
  private _label = '';

  protected onInit(): Promise<void> { return Promise.resolve(); }

  // ---- home / defaults -----------------------------------------------------
  private _homeLat(): string { return (this.properties.latitude || '36.52146').trim(); }
  private _homeLon(): string { return (this.properties.longitude || '-86.026315').trim(); }
  private _homeLabel(): string { return (this.properties.locationLabel || 'Lafayette, TN').trim(); }
  private _fallbackSite(): string { return (this.properties.radarSite || 'KOHX').trim().toUpperCase(); }

  public render(): void {
    this._injectCss();
    const title = (this.properties.title || 'Area Weather').trim();
    // reset to home on every (re)render -> page reload always shows the office
    this._lat = this._homeLat();
    this._lon = this._homeLon();
    this._label = this._homeLabel();

    this.domElement.innerHTML = `
      <div class="dlswx">
        <div class="dlswx-head">
          <div class="dlswx-head-title">
            <span class="dlswx-title">${esc(title)}</span>
            <span class="dlswx-loc" id="dlswx-loc">${esc(this._label)}</span>
          </div>
          <form class="dlswx-search" id="dlswx-form" autocomplete="off">
            <input id="dlswx-q" class="dlswx-q" type="text" placeholder="Search address or city, state" aria-label="Search a location" />
            <button type="submit" class="dlswx-btn" id="dlswx-go" title="Search this location" aria-label="Search">
              <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M15.5 14h-.79l-.28-.27a6.5 6.5 0 1 0-.7.7l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0A4.5 4.5 0 1 1 14 9.5 4.5 4.5 0 0 1 9.5 14z"/></svg>
            </button>
            <button type="button" class="dlswx-btn" id="dlswx-home" title="Back to Lafayette (107 Scottsville Rd)" aria-label="Reset to Lafayette">
              <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3 4 9v12h6v-7h4v7h6V9z"/></svg>
            </button>
            <span class="dlswx-msg" id="dlswx-msg"></span>
          </form>
        </div>
        <div id="dlswx-alerts" class="dlswx-alerts" style="display:none"></div>
        <div class="dlswx-grid">
          <div class="dlswx-radar">
            <div class="dlswx-sub">Live radar</div>
            <img id="dlswx-radar-img" alt="NWS radar loop" />
            <div class="dlswx-radar-cap" id="dlswx-radar-cap">NOAA / NWS RIDGE radar &middot; loops ~1 hr</div>
          </div>
          <div class="dlswx-now"><div id="dlswx-current" class="dlswx-current"></div></div>
          <div class="dlswx-hourwrap"><div class="dlswx-sub">Next 12 hours</div><div id="dlswx-hourly" class="dlswx-hourly"></div></div>
          <div class="dlswx-weekwrap"><div class="dlswx-sub">7-day outlook</div><div id="dlswx-7day" class="dlswx-7day"></div></div>
        </div>
        <div class="dlswx-foot">
          Data: NOAA / National Weather Service (api.weather.gov), geocoding by OpenStreetMap &middot;
          <span id="dlswx-updated">updating&hellip;</span> &middot; <a id="dlswx-refresh" href="#">refresh</a>
        </div>
      </div>`;

    const form = this.domElement.querySelector('#dlswx-form') as HTMLFormElement;
    if (form) { form.onsubmit = (e) => { e.preventDefault(); this._search(); }; }
    const home = this.domElement.querySelector('#dlswx-home') as HTMLElement;
    if (home) { home.onclick = () => { this._goHome(); }; }
    const refreshLink = this.domElement.querySelector('#dlswx-refresh') as HTMLElement;
    if (refreshLink) { refreshLink.onclick = (e) => { e.preventDefault(); this._loadAll(); }; }

    this._loadAll();
    if (this._timer) { clearInterval(this._timer); }
    this._timer = setInterval(() => this._loadAll(), REFRESH_MS);
  }

  protected onDispose(): void {
    if (this._timer) { clearInterval(this._timer); this._timer = undefined; }
  }

  // ---- search / location ---------------------------------------------------

  private _search(): void {
    const input = this.domElement.querySelector('#dlswx-q') as HTMLInputElement;
    const q = ((input && input.value) || '').trim();
    if (!q) { return; }
    this._msg('Searching…');
    this._geocode(q)
      .then(res => {
        if (!res) { this._msg('No match — try “City, ST”.'); return; }
        this._lat = res.lat; this._lon = res.lon; this._label = res.label;
        this._msg('');
        this._loadAll();
      })
      .catch(() => { this._msg('Search unavailable — try again.'); });
  }

  private _goHome(): void {
    const input = this.domElement.querySelector('#dlswx-q') as HTMLInputElement;
    if (input) { input.value = ''; }
    this._lat = this._homeLat(); this._lon = this._homeLon(); this._label = this._homeLabel();
    this._msg('');
    this._loadAll();
  }

  // Keyless US geocoder (OpenStreetMap / Nominatim). Only called on an explicit
  // search, so volume stays well within acceptable use.
  private _geocode(q: string): Promise<{ lat: string; lon: string; label: string } | null> {
    const url = 'https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=us&q=' + encodeURIComponent(q);
    return fetch(url, { headers: { 'Accept': 'application/json' } })
      .then(r => r.json())
      .then((j: any) => {
        if (j && j.length) { return { lat: String(j[0].lat), lon: String(j[0].lon), label: q }; }
        return null;
      });
  }

  // ---- data ----------------------------------------------------------------

  private _loadAll(): void {
    this._setLoc(this._label);
    this._loadPoint();
    this._loadAlerts();
    const u = this.domElement.querySelector('#dlswx-updated');
    if (u) { u.textContent = 'Updated ' + fmtTime(new Date()); }
  }

  // One /points call resolves: nearest radar station, the forecast URLs, and the
  // city/state label - so radar + forecast + label all follow the searched point.
  private _loadPoint(): void {
    fetch('https://api.weather.gov/points/' + this._lat + ',' + this._lon, { headers: { 'Accept': 'application/geo+json' } })
      .then(r => r.json())
      .then(pj => {
        const props = (pj && pj.properties) || {};
        this._setRadar((props.radarStation || this._fallbackSite()));
        const rl = props.relativeLocation && props.relativeLocation.properties;
        if (rl && rl.city && rl.state) { this._setLoc(rl.city + ', ' + rl.state); }
        if (props.forecastHourly) { this._renderHourly(props.forecastHourly); }
        if (props.forecast) { this._render7day(props.forecast); }
        if (!props.forecast && !props.forecastHourly) {
          this._fail('#dlswx-current', 'Forecast unavailable for this location.');
          this._fail('#dlswx-hourly', ''); this._fail('#dlswx-7day', '');
        }
      })
      .catch(() => {
        this._setRadar(this._fallbackSite());
        this._fail('#dlswx-current', 'Forecast unavailable for this location.');
        this._fail('#dlswx-hourly', ''); this._fail('#dlswx-7day', '');
      });
  }

  private _setRadar(station: string): void {
    const st = String(station || this._fallbackSite()).toUpperCase();
    const img = this.domElement.querySelector('#dlswx-radar-img') as HTMLImageElement;
    if (img) { img.src = 'https://radar.weather.gov/ridge/standard/' + st + '_loop.gif?cb=' + Date.now(); }
    const cap = this.domElement.querySelector('#dlswx-radar-cap');
    if (cap) { cap.textContent = 'NOAA / NWS ' + st + ' radar · loops ~1 hr'; }
  }

  private _loadAlerts(): void {
    const box = this.domElement.querySelector('#dlswx-alerts') as HTMLElement;
    if (!box) { return; }
    fetch('https://api.weather.gov/alerts/active?point=' + this._lat + ',' + this._lon,
          { headers: { 'Accept': 'application/geo+json' } })
      .then(r => r.json())
      .then(j => {
        const feats = (j && j.features) || [];
        if (!feats.length) { box.style.display = 'none'; box.innerHTML = ''; return; }
        box.innerHTML = feats.slice(0, 4).map((f: any) => {
          const p = f.properties || {};
          const sev = (p.severity || '').toLowerCase();
          const cls = (sev === 'extreme' || sev === 'severe') ? 'dlswx-alert-sev' : 'dlswx-alert-mod';
          return '<div class="dlswx-alert ' + cls + '">&#9888; ' + esc(p.event || 'Weather alert') +
                 (p.headline ? '<span class="dlswx-alert-h">' + esc(p.headline) + '</span>' : '') + '</div>';
        }).join('');
        box.style.display = 'block';
      })
      .catch(() => { box.style.display = 'none'; });
  }

  private _renderHourly(url: string): void {
    fetch(url, { headers: { 'Accept': 'application/geo+json' } })
      .then(r => r.json())
      .then(j => {
        const periods = ((j.properties && j.properties.periods) || []).slice(0, 12);
        if (!periods.length) { throw new Error('empty hourly'); }
        const c = periods[0];
        const cur = this.domElement.querySelector('#dlswx-current') as HTMLElement;
        if (cur) {
          cur.innerHTML =
            (c.icon ? '<img class="dlswx-cur-icon" src="' + esc(c.icon) + '" alt="" />' : '') +
            '<div class="dlswx-cur-temp">' + Number(c.temperature) + '&deg;' + esc(c.temperatureUnit || 'F') + '</div>' +
            '<div class="dlswx-cur-desc">' + esc(c.shortForecast || '') +
              '<div class="dlswx-cur-wind">Wind ' + esc(c.windSpeed || '') + ' ' + esc(c.windDirection || '') + '</div>' +
            '</div>';
        }
        const wrap = this.domElement.querySelector('#dlswx-hourly') as HTMLElement;
        if (wrap) {
          wrap.innerHTML = periods.map((p: any) => {
            const pop = (p.probabilityOfPrecipitation && p.probabilityOfPrecipitation.value) || 0;
            const hot = pop >= 30 ? ' dlswx-pop-hi' : '';
            return '<div class="dlswx-hr">' +
                   '<div class="dlswx-hr-t">' + fmtHour(p.startTime) + '</div>' +
                   '<div class="dlswx-hr-bar"><div class="dlswx-hr-fill' + hot + '" style="height:' + Math.max(4, pop) + '%"></div></div>' +
                   '<div class="dlswx-hr-p">' + pop + '%</div>' +
                   '<div class="dlswx-hr-temp">' + Number(p.temperature) + '&deg;</div>' +
                   '</div>';
          }).join('');
        }
      })
      .catch(() => {
        this._fail('#dlswx-current', 'Hourly data temporarily unavailable.');
        this._fail('#dlswx-hourly', '');
      });
  }

  private _render7day(url: string): void {
    fetch(url, { headers: { 'Accept': 'application/geo+json' } })
      .then(r => r.json())
      .then(j => {
        const periods = (j.properties && j.properties.periods) || [];
        if (!periods.length) { throw new Error('empty forecast'); }
        const rows: string[] = [];
        for (let i = 0; i < periods.length && rows.length < 7; i++) {
          const p = periods[i];
          if (!p.isDaytime && rows.length > 0) { continue; }
          const next = periods[i + 1];
          const lo = (next && !next.isDaytime) ? next.temperature : null;
          const hi = p.isDaytime ? p.temperature : null;
          const pop = (p.probabilityOfPrecipitation && p.probabilityOfPrecipitation.value) || 0;
          rows.push(
            '<div class="dlswx-d">' +
              '<span class="dlswx-d-n">' + esc(shortName(p.name)) + '</span>' +
              (p.icon ? '<img class="dlswx-d-i" src="' + esc(p.icon) + '" alt="" title="' + esc(p.shortForecast || '') + '"/>' : '<span class="dlswx-d-i"></span>') +
              '<span class="dlswx-d-s">' + esc(p.shortForecast || '') + '</span>' +
              '<span class="dlswx-d-pop' + (pop >= 30 ? ' dlswx-pop-hi' : '') + '">&#128167;' + pop + '%</span>' +
              '<span class="dlswx-d-t">' +
                (hi !== null ? '<span class="dlswx-d-hi">' + hi + '&deg;</span>' : '') +
                (lo !== null ? '<span class="dlswx-d-lo">' + lo + '&deg;</span>' : '') +
              '</span>' +
            '</div>');
        }
        const wrap = this.domElement.querySelector('#dlswx-7day') as HTMLElement;
        if (wrap) { wrap.innerHTML = rows.join(''); }
      })
      .catch(() => { this._fail('#dlswx-7day', '7-day outlook temporarily unavailable.'); });
  }

  private _fail(sel: string, msg: string): void {
    const el = this.domElement.querySelector(sel) as HTMLElement;
    if (el) { el.innerHTML = msg ? '<div class="dlswx-err">' + esc(msg) + '</div>' : ''; }
  }

  private _msg(t: string): void {
    const e = this.domElement.querySelector('#dlswx-msg'); if (e) { e.textContent = t; }
  }
  private _setLoc(t: string): void {
    const e = this.domElement.querySelector('#dlswx-loc'); if (e) { e.textContent = t; }
  }

  // ---- styling -------------------------------------------------------------

  private _injectCss(): void {
    if (this._cssInjected) { return; }
    const s = document.createElement('style');
    s.textContent = DLSWX_CSS;
    document.head.appendChild(s);
    this._cssInjected = true;
  }

  // ---- property pane -------------------------------------------------------

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'Default ("home") weather location. The search box on the panel can switch to any US address or city; every reload returns to this home point.' },
        groups: [{
          groupName: 'Home location',
          groupFields: [
            PropertyPaneTextField('title', { label: 'Panel title' }),
            PropertyPaneTextField('locationLabel', { label: 'Home label (e.g. Lafayette, TN)' }),
            PropertyPaneTextField('latitude', { label: 'Home latitude' }),
            PropertyPaneTextField('longitude', { label: 'Home longitude' }),
            PropertyPaneTextField('radarSite', { label: 'Fallback radar code (used only if NWS lookup fails)' })
          ]
        }]
      }]
    };
  }
}

// ---- helpers ----------------------------------------------------------------

function esc(s: any): string {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function fmtTime(d: Date): string {
  let h = d.getHours(); const m = d.getMinutes(); const ap = h >= 12 ? 'PM' : 'AM';
  h = h % 12; if (h === 0) { h = 12; }
  return h + ':' + (m < 10 ? '0' + m : m) + ' ' + ap;
}
function fmtHour(iso: string): string {
  try {
    const d = new Date(iso); let h = d.getHours(); const ap = h >= 12 ? 'p' : 'a';
    h = h % 12; if (h === 0) { h = 12; } return h + ap;
  } catch (e) { return ''; }
}
function shortName(n: string): string {
  if (!n) { return ''; }
  return n.replace('This Afternoon', 'Today').replace('This Morning', 'Today');
}

// DARK theme, v1.0.0.7. Full-bleed (98vw) to line up with the Deed Map; crisp ~660px
// radar; full-width 7-day strip; header search box (top-right). Scoped to .dlswx.
const DLSWX_CSS = `
.dlswx{font-family:'Segoe UI',system-ui,sans-serif;color:#e6e4e2;border:1px solid #3b3a39;border-radius:8px;overflow:hidden;background:#1b1a19;box-shadow:0 1.6px 3.6px rgba(0,0,0,.45);width:100%;box-sizing:border-box}
@media (min-width:1300px) and (orientation:landscape){.dlswx{width:98vw;position:relative;left:50%;margin-left:-49vw;border-radius:10px}}
.dlswx *{box-sizing:border-box}
.dlswx-head{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px 14px;padding:10px 16px;background:#0a2c49;color:#fff;border-bottom:1px solid #16334f}
.dlswx-head-title{display:flex;align-items:baseline;gap:8px;min-width:0;flex:1 1 auto}
.dlswx-title{font-size:16px;font-weight:600;white-space:nowrap}
.dlswx-loc{font-size:12px;opacity:.85;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.dlswx-search{display:flex;align-items:center;gap:6px;flex:0 1 auto}
.dlswx-q{width:240px;max-width:52vw;height:30px;padding:0 10px;border-radius:6px;border:1px solid #29547d;background:#0f3c63;color:#fff;font-size:13px;font-family:inherit;outline:none}
.dlswx-q:focus{border-color:#4a90d9;background:#114a72}
.dlswx-q::placeholder{color:#9fc0db}
.dlswx-btn{display:inline-flex;align-items:center;justify-content:center;width:30px;height:30px;flex:0 0 auto;border-radius:6px;border:1px solid #29547d;background:#0f3c63;color:#dceaf7;cursor:pointer;padding:0}
.dlswx-btn:hover{background:#16527f}
.dlswx-btn svg{width:16px;height:16px;fill:currentColor;display:block}
.dlswx-msg{font-size:11px;color:#f3d9a0;min-height:14px;white-space:nowrap}
.dlswx-alerts{padding:8px 16px 0}
.dlswx-alert{padding:7px 11px;border-radius:6px;margin-bottom:6px;font-size:13px;font-weight:600}
.dlswx-alert-sev{background:#3b1416;color:#f5a3a6;border:1px solid #7a2a2e}
.dlswx-alert-mod{background:#3a3209;color:#f3e0a0;border:1px solid #7a6a1e}
.dlswx-alert-h{display:block;font-weight:400;font-size:12px;margin-top:2px;opacity:.9}
.dlswx-grid{display:grid;grid-template-columns:1fr;gap:14px;padding:12px 16px;grid-template-areas:"rad" "now" "hour" "week"}
.dlswx-radar{grid-area:rad;display:flex;flex-direction:column;min-width:0}
.dlswx-now{grid-area:now;min-width:0}
.dlswx-hourwrap{grid-area:hour;min-width:0}
.dlswx-weekwrap{grid-area:week;min-width:0}
.dlswx-sub{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;color:#a19f9d;margin:0 0 6px}
.dlswx-radar img{width:100%;max-width:560px;aspect-ratio:16/9;object-fit:cover;object-position:center 40%;border:1px solid #3b3a39;border-radius:6px;display:block;background:#0f0f0f}
.dlswx-radar-cap{font-size:11px;color:#7a7775;margin-top:4px}
.dlswx-current{display:flex;align-items:center;gap:12px;margin:0}
.dlswx-cur-icon{width:50px;height:50px;border-radius:6px;background:#2d2c2b}
.dlswx-cur-temp{font-size:30px;font-weight:600;line-height:1;color:#fff}
.dlswx-cur-desc{font-size:13px;color:#e1dfdd}
.dlswx-cur-wind{font-size:12px;color:#a19f9d;margin-top:2px}
.dlswx-hourly{display:flex;gap:4px;overflow-x:auto;padding-bottom:2px}
.dlswx-hr{flex:0 0 auto;width:32px;text-align:center}
.dlswx-hr-t{font-size:10px;color:#a19f9d}
.dlswx-hr-bar{height:44px;width:13px;margin:2px auto;background:#2d2c2b;border-radius:3px;display:flex;align-items:flex-end;overflow:hidden}
.dlswx-hr-fill{width:100%;background:#4a90d9}
.dlswx-hr-fill.dlswx-pop-hi{background:#2b88e0}
.dlswx-hr-p{font-size:10px;color:#7fb1de;font-weight:600}
.dlswx-hr-temp{font-size:10px;color:#a19f9d}
.dlswx-7day{display:flex;flex-direction:column}
.dlswx-d{display:flex;align-items:center;gap:8px;padding:5px 0;border-top:1px solid #2d2c2b;font-size:12px}
.dlswx-d:first-child{border-top:none}
.dlswx-d-n{flex:0 0 58px;font-weight:600;color:#f3f2f1}
.dlswx-d-i{flex:0 0 auto;width:26px;height:26px;background:transparent}
.dlswx-d-s{flex:1 1 auto;color:#a19f9d;font-size:11px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.dlswx-d-pop{flex:0 0 46px;text-align:right;color:#7fb1de;font-size:11px}
.dlswx-d-pop.dlswx-pop-hi{color:#4aa3e8;font-weight:700}
.dlswx-d-t{flex:0 0 58px;text-align:right}
.dlswx-d-hi{font-weight:700;color:#fff}
.dlswx-d-lo{color:#8a8886;margin-left:4px}
.dlswx-foot{font-size:11px;color:#7a7775;padding:7px 16px 10px;border-top:1px solid #2d2c2b}
.dlswx-foot a{color:#7fb1de}
.dlswx-err{font-size:12px;color:#a19f9d;font-style:italic;padding:6px 0}

/* MEDIUM (normal section width): radar left, info column right, 7-day list */
@media (min-width:621px){
.dlswx-grid{grid-template-columns:minmax(300px,1fr) minmax(320px,1.1fr);grid-template-areas:"rad now" "rad hour" "rad week";gap:14px 20px;align-items:start}
.dlswx-radar{align-self:center}
}

/* WIDE (full-bleed, matches the Deed Map): crisp radar + current + tall 12h on top; big full-width 7-day strip below */
@media (min-width:1300px) and (orientation:landscape){
.dlswx-grid{grid-template-columns:auto minmax(200px,0.85fr) minmax(440px,1.6fr);grid-template-rows:auto auto;grid-template-areas:"rad now hour" "week week week";gap:18px 26px;padding:16px 22px;align-items:center}
.dlswx-radar{align-self:start}
.dlswx-radar img{width:660px;max-width:42vw;aspect-ratio:16/9}
.dlswx-now{align-self:center}
.dlswx-hourwrap{align-self:center}
.dlswx-current{flex-direction:column;align-items:flex-start;gap:10px}
.dlswx-cur-icon{width:76px;height:76px}
.dlswx-cur-temp{font-size:64px}
.dlswx-cur-desc{font-size:17px}
.dlswx-cur-wind{font-size:14px;margin-top:3px}
.dlswx-hourly{gap:8px;overflow-x:visible}
.dlswx-hr{flex:1 1 0;width:auto;max-width:60px}
.dlswx-hr-t{font-size:13px}
.dlswx-hr-bar{height:200px;width:100%;max-width:30px;margin:3px auto}
.dlswx-hr-p{font-size:13px}
.dlswx-hr-temp{font-size:13px}
.dlswx-weekwrap .dlswx-sub{font-size:12px}
.dlswx-7day{flex-direction:row;gap:12px}
.dlswx-d{flex:1 1 0;flex-direction:column;align-items:center;justify-content:flex-start;text-align:center;gap:7px;border-top:none;background:#232221;border:1px solid #2d2c2b;border-radius:10px;padding:16px 12px}
.dlswx-d-n{flex:none;font-size:17px;color:#f3f2f1}
.dlswx-d-i{width:66px;height:66px}
.dlswx-d-s{flex:none;white-space:normal;font-size:13.5px;line-height:1.3;color:#c8c6c4}
.dlswx-d-pop{flex:none;font-size:15px}
.dlswx-d-t{flex:none;font-size:20px;margin-top:2px}
.dlswx-d-hi{font-size:21px}
.dlswx-d-lo{font-size:17px;margin-left:8px}
}
`;
