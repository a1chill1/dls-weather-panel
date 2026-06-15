/* eslint-disable */
// ============================================================================
// DLS Weather Panel - SPFx client-side web part (no framework)
// ----------------------------------------------------------------------------
// A scheduling-focused weather panel for the DossSurveying home page. 100% NOAA /
// National Weather Service data: animated radar loop (RIDGE GIF), active weather
// alerts, next-12h precip-chance timeline, and a 7-day forecast.
//
// WHY THIS DESIGN (durability, ~5-yr target):
//   * Data source = api.weather.gov + radar.weather.gov  -> US Government, free,
//     NO API KEY, no quota, no per-seat license, no third-party host that can
//     disappear or change its terms. About as future-proof as a live feed gets.
//   * Everything is BUNDLED in this one file. No external <script> / CDN (this
//     tenant blocks those) and no iframe (so no embed allow-list dependency).
//     Radar + weather icons load as <img>, which the tenant permits.
//   * FAILS GRACEFULLY. Every network call is wrapped; if NWS is unreachable the
//     affected section shows a quiet "temporarily unavailable" note and the rest
//     of the panel (and the whole home page) keeps working. It can never throw
//     up and break the page.
//
// CONFIG (property pane): title, location label, latitude, longitude, radar site.
// Defaults are set for Lafayette / Macon Co. TN (NWS grid OHX 76,75, radar KOHX).
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

// Refresh cadence for the live data (radar GIF is also cache-busted each cycle).
const REFRESH_MS = 20 * 60 * 1000; // 20 minutes

export default class WeatherPanelWebPart extends BaseClientSideWebPart<IWeatherPanelWebPartProps> {
  private _timer: any = undefined;
  private _cssInjected = false;

  protected onInit(): Promise<void> {
    return Promise.resolve();
  }

  public render(): void {
    this._injectCss();

    const title = (this.properties.title || 'Area Weather').trim();
    const loc = (this.properties.locationLabel || 'Lafayette, TN').trim();

    this.domElement.innerHTML = `
      <div class="dlswx">
        <div class="dlswx-head">
          <div>
            <span class="dlswx-title">${esc(title)}</span>
            <span class="dlswx-loc">${esc(loc)}</span>
          </div>
          <div class="dlswx-updated" id="dlswx-updated">loading&hellip;</div>
        </div>
        <div id="dlswx-alerts" class="dlswx-alerts" style="display:none"></div>
        <div class="dlswx-grid">
          <div class="dlswx-radar">
            <div class="dlswx-sub">Live radar</div>
            <img id="dlswx-radar-img" alt="NWS radar loop" />
            <div class="dlswx-radar-cap">NOAA / NWS RIDGE radar &middot; loops ~1 hr</div>
          </div>
          <div class="dlswx-now">
            <div class="dlswx-sub">Now &amp; next 12 hours</div>
            <div id="dlswx-current" class="dlswx-current"></div>
            <div id="dlswx-hourly" class="dlswx-hourly"></div>
          </div>
        </div>
        <div class="dlswx-sub dlswx-sub-7">7-day outlook</div>
        <div id="dlswx-7day" class="dlswx-7day"></div>
        <div class="dlswx-foot">
          Data: NOAA / National Weather Service (api.weather.gov) &middot;
          <a id="dlswx-refresh" href="#">refresh</a>
        </div>
      </div>`;

    const refreshLink = this.domElement.querySelector('#dlswx-refresh') as HTMLElement;
    if (refreshLink) {
      refreshLink.onclick = (e) => { e.preventDefault(); this._loadAll(); };
    }

    this._loadAll();

    // (re)arm the auto-refresh timer
    if (this._timer) { clearInterval(this._timer); }
    this._timer = setInterval(() => this._loadAll(), REFRESH_MS);
  }

  protected onDispose(): void {
    if (this._timer) { clearInterval(this._timer); this._timer = undefined; }
  }

  // ---- data loading ---------------------------------------------------------

  private _lat(): string { return (this.properties.latitude || '36.524').trim(); }
  private _lon(): string { return (this.properties.longitude || '-86.026').trim(); }
  private _site(): string { return (this.properties.radarSite || 'KOHX').trim().toUpperCase(); }

  private _loadAll(): void {
    this._loadRadar();
    this._loadAlerts();
    this._loadForecast();
    const u = this.domElement.querySelector('#dlswx-updated');
    if (u) { u.textContent = 'Updated ' + fmtTime(new Date()); }
  }

  private _loadRadar(): void {
    const img = this.domElement.querySelector('#dlswx-radar-img') as HTMLImageElement;
    if (!img) { return; }
    img.src = 'https://radar.weather.gov/ridge/standard/' + this._site() + '_loop.gif?cb=' + Date.now();
  }

  private _loadAlerts(): void {
    const box = this.domElement.querySelector('#dlswx-alerts') as HTMLElement;
    if (!box) { return; }
    fetch('https://api.weather.gov/alerts/active?point=' + this._lat() + ',' + this._lon(),
          { headers: { 'Accept': 'application/geo+json' } })
      .then(r => r.json())
      .then(j => {
        const feats = (j && j.features) || [];
        if (!feats.length) { box.style.display = 'none'; box.innerHTML = ''; return; }
        const items = feats.slice(0, 4).map((f: any) => {
          const p = f.properties || {};
          const sev = (p.severity || '').toLowerCase();
          const cls = (sev === 'extreme' || sev === 'severe') ? 'dlswx-alert-sev' : 'dlswx-alert-mod';
          return '<div class="dlswx-alert ' + cls + '">&#9888; ' + esc(p.event || 'Weather alert') +
                 (p.headline ? '<span class="dlswx-alert-h">' + esc(p.headline) + '</span>' : '') + '</div>';
        }).join('');
        box.innerHTML = items;
        box.style.display = 'block';
      })
      .catch(() => { box.style.display = 'none'; });
  }

  private _loadForecast(): void {
    const lat = this._lat(), lon = this._lon();
    // points -> the canonical forecast + hourly URLs for this coordinate
    fetch('https://api.weather.gov/points/' + lat + ',' + lon, { headers: { 'Accept': 'application/geo+json' } })
      .then(r => r.json())
      .then(pj => {
        const props = (pj && pj.properties) || {};
        const fUrl = props.forecast;
        const hUrl = props.forecastHourly;
        if (hUrl) { this._renderHourly(hUrl); }
        if (fUrl) { this._render7day(fUrl); }
        if (!fUrl && !hUrl) { throw new Error('no forecast urls'); }
      })
      .catch(() => {
        this._fail('#dlswx-current', 'Forecast temporarily unavailable.');
        this._fail('#dlswx-hourly', '');
        this._fail('#dlswx-7day', '');
      });
  }

  private _renderHourly(url: string): void {
    fetch(url, { headers: { 'Accept': 'application/geo+json' } })
      .then(r => r.json())
      .then(j => {
        const periods = ((j.properties && j.properties.periods) || []).slice(0, 12);
        if (!periods.length) { throw new Error('empty hourly'); }
        // current conditions = first hourly period
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
        // precip-chance timeline
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
        // Pair day/night periods into <=7 day cards.
        const cards: string[] = [];
        for (let i = 0; i < periods.length && cards.length < 7; i++) {
          const p = periods[i];
          if (!p.isDaytime && cards.length === 0) {
            // starts on a "Tonight" -> render a single night card first
          } else if (!p.isDaytime) {
            continue; // night handled alongside its day below
          }
          const next = periods[i + 1];
          const lo = (next && !next.isDaytime) ? next.temperature : null;
          const hi = p.isDaytime ? p.temperature : null;
          const pop = (p.probabilityOfPrecipitation && p.probabilityOfPrecipitation.value) || 0;
          cards.push(
            '<div class="dlswx-day">' +
              '<div class="dlswx-day-n">' + esc(shortName(p.name)) + '</div>' +
              (p.icon ? '<img class="dlswx-day-icon" src="' + esc(p.icon) + '" alt="" title="' + esc(p.shortForecast || '') + '"/>' : '') +
              '<div class="dlswx-day-temps">' +
                (hi !== null ? '<span class="dlswx-hi">' + hi + '&deg;</span>' : '') +
                (lo !== null ? '<span class="dlswx-lo">' + lo + '&deg;</span>' : '') +
              '</div>' +
              '<div class="dlswx-day-pop' + (pop >= 30 ? ' dlswx-pop-hi' : '') + '">&#128167; ' + pop + '%</div>' +
              '<div class="dlswx-day-s">' + esc(p.shortForecast || '') + '</div>' +
            '</div>');
        }
        const wrap = this.domElement.querySelector('#dlswx-7day') as HTMLElement;
        if (wrap) { wrap.innerHTML = cards.join(''); }
      })
      .catch(() => { this._fail('#dlswx-7day', '7-day outlook temporarily unavailable.'); });
  }

  private _fail(sel: string, msg: string): void {
    const el = this.domElement.querySelector(sel) as HTMLElement;
    if (el) { el.innerHTML = msg ? '<div class="dlswx-err">' + esc(msg) + '</div>' : ''; }
  }

  // ---- styling --------------------------------------------------------------

  private _injectCss(): void {
    if (this._cssInjected) { return; }
    const s = document.createElement('style');
    s.textContent = DLSWX_CSS;
    document.head.appendChild(s);
    this._cssInjected = true;
  }

  // ---- property pane --------------------------------------------------------

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'Weather panel settings (NOAA/NWS). Find a radar site code at radar.weather.gov.' },
        groups: [{
          groupName: 'Location',
          groupFields: [
            PropertyPaneTextField('title', { label: 'Panel title' }),
            PropertyPaneTextField('locationLabel', { label: 'Location label (display only)' }),
            PropertyPaneTextField('latitude', { label: 'Latitude' }),
            PropertyPaneTextField('longitude', { label: 'Longitude' }),
            PropertyPaneTextField('radarSite', { label: 'NWS radar site code (e.g. KOHX = Nashville)' })
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

// Self-contained styles, scoped under .dlswx. Uses SharePoint theme neutrals where
// reasonable; lays out responsively and wraps on narrow screens.
// DARK THEME (v1.0.0.2) - matches the dark DossSurveying home page. Scoped under .dlswx.
const DLSWX_CSS = `
.dlswx{font-family:'Segoe UI',system-ui,sans-serif;color:#e6e4e2;border:1px solid #3b3a39;border-radius:8px;overflow:hidden;background:#1b1a19;box-shadow:0 1.6px 3.6px rgba(0,0,0,.45)}
.dlswx-head{display:flex;justify-content:space-between;align-items:center;padding:12px 16px;background:#0a2c49;color:#fff;border-bottom:1px solid #16334f}
.dlswx-title{font-size:18px;font-weight:600}
.dlswx-loc{font-size:13px;opacity:.85;margin-left:8px}
.dlswx-updated{font-size:12px;opacity:.8}
.dlswx-alerts{padding:8px 16px 0}
.dlswx-alert{padding:8px 12px;border-radius:6px;margin-bottom:6px;font-size:13px;font-weight:600}
.dlswx-alert-sev{background:#3b1416;color:#f5a3a6;border:1px solid #7a2a2e}
.dlswx-alert-mod{background:#3a3209;color:#f3e0a0;border:1px solid #7a6a1e}
.dlswx-alert-h{display:block;font-weight:400;font-size:12px;margin-top:2px;opacity:.9}
.dlswx-grid{display:flex;gap:16px;padding:12px 16px}
.dlswx-radar{flex:1 1 360px;min-width:280px}
.dlswx-now{flex:1 1 320px;min-width:260px}
.dlswx-sub{font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;color:#a19f9d;margin:0 0 6px}
.dlswx-sub-7{padding:0 16px}
.dlswx-radar img{width:100%;height:auto;border:1px solid #3b3a39;border-radius:6px;display:block;background:#0f0f0f;min-height:120px}
.dlswx-radar-cap{font-size:11px;color:#7a7775;margin-top:4px}
.dlswx-current{display:flex;align-items:center;gap:10px;margin-bottom:10px}
.dlswx-cur-icon{width:54px;height:54px;border-radius:6px;background:#2d2c2b}
.dlswx-cur-temp{font-size:30px;font-weight:600;line-height:1;color:#fff}
.dlswx-cur-desc{font-size:13px;color:#e1dfdd}
.dlswx-cur-wind{font-size:12px;color:#a19f9d;margin-top:2px}
.dlswx-hourly{display:flex;gap:5px;overflow-x:auto;padding-bottom:4px}
.dlswx-hr{flex:0 0 auto;width:34px;text-align:center}
.dlswx-hr-t{font-size:11px;color:#a19f9d}
.dlswx-hr-bar{height:48px;width:14px;margin:3px auto;background:#2d2c2b;border-radius:3px;display:flex;align-items:flex-end;overflow:hidden}
.dlswx-hr-fill{width:100%;background:#4a90d9}
.dlswx-hr-fill.dlswx-pop-hi{background:#2b88e0}
.dlswx-hr-p{font-size:11px;color:#7fb1de;font-weight:600}
.dlswx-hr-temp{font-size:11px;color:#a19f9d}
.dlswx-7day{display:flex;gap:8px;padding:6px 16px 14px;overflow-x:auto}
.dlswx-day{flex:0 0 auto;width:96px;text-align:center;border:1px solid #3b3a39;border-radius:6px;padding:8px 4px;background:#252423}
.dlswx-day-n{font-size:12px;font-weight:600;color:#f3f2f1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.dlswx-day-icon{width:42px;height:42px;margin:4px 0}
.dlswx-day-temps{font-size:14px}
.dlswx-hi{font-weight:700;color:#fff}
.dlswx-lo{color:#8a8886;margin-left:5px}
.dlswx-day-pop{font-size:11px;color:#7fb1de}
.dlswx-day-pop.dlswx-pop-hi{color:#4aa3e8;font-weight:700}
.dlswx-day-s{font-size:10px;color:#a19f9d;line-height:1.25;margin-top:3px;height:38px;overflow:hidden}
.dlswx-foot{font-size:11px;color:#7a7775;padding:8px 16px 12px;border-top:1px solid #2d2c2b}
.dlswx-foot a{color:#7fb1de}
.dlswx-err{font-size:12px;color:#a19f9d;font-style:italic;padding:6px 0}
@media (max-width:640px){.dlswx-grid{flex-direction:column}}
`;
