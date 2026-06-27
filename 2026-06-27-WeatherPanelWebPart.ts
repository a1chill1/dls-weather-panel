/* eslint-disable */
// ============================================================================
// DLS Weather Panel - SPFx client-side web part (no framework)
// ----------------------------------------------------------------------------
// Scheduling-focused weather panel for the DossSurveying home page. 100% NOAA /
// National Weather Service data: animated radar loop (RIDGE GIF), active weather
// alerts, next-12h precip-chance timeline, and a 7-day forecast.
//
// v1.0.0.5 - READABILITY REBALANCE. The NWS loop GIF is only 600px native, so the
//   v1.0.0.4 hero (~1355px wide) UPSCALED it ~2.3x -> pixelated. Fix:
//     * Radar capped near its native width (<=560px) so it stays CRISP (no upscale),
//       framed as a 16:9 landscape (keeps central TN + southern KY).
//     * The freed space goes to a BIG, full-width 7-day strip (large cards, larger
//       type + icons) so the forecast is easy to read.
//   WIDE layout = top row [ crisp radar | current conditions | tall 12h timeline ]
//   then a full-width 7-day card strip beneath. Still full-bleed 98vw to match the
//   Deed Map. Medium = radar left + info right (7-day list). Mobile = stacked.
//   Same NWS fetch logic / pinned IDs (upgrade in place).
//
// Durability: api.weather.gov + radar.weather.gov (US Gov, free, no key). Bundled,
// no external script/CDN, no iframe. Every call wrapped -> fails safe, never breaks
// the page. Config (property pane): title, location label, latitude, longitude, radar.
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

  protected onInit(): Promise<void> { return Promise.resolve(); }

  public render(): void {
    this._injectCss();
    const title = (this.properties.title || 'Area Weather').trim();
    const loc = (this.properties.locationLabel || 'Lafayette, TN').trim();

    this.domElement.innerHTML = `
      <div class="dlswx">
        <div class="dlswx-head">
          <div><span class="dlswx-title">${esc(title)}</span><span class="dlswx-loc">${esc(loc)}</span></div>
          <div class="dlswx-updated" id="dlswx-updated">loading&hellip;</div>
        </div>
        <div id="dlswx-alerts" class="dlswx-alerts" style="display:none"></div>
        <div class="dlswx-grid">
          <div class="dlswx-radar">
            <div class="dlswx-sub">Live radar</div>
            <img id="dlswx-radar-img" alt="NWS radar loop" />
            <div class="dlswx-radar-cap">NOAA / NWS RIDGE radar &middot; loops ~1 hr</div>
          </div>
          <div class="dlswx-now"><div id="dlswx-current" class="dlswx-current"></div></div>
          <div class="dlswx-hourwrap"><div class="dlswx-sub">Next 12 hours</div><div id="dlswx-hourly" class="dlswx-hourly"></div></div>
          <div class="dlswx-weekwrap"><div class="dlswx-sub">7-day outlook</div><div id="dlswx-7day" class="dlswx-7day"></div></div>
        </div>
        <div class="dlswx-foot">
          Data: NOAA / National Weather Service (api.weather.gov) &middot;
          <a id="dlswx-refresh" href="#">refresh</a>
        </div>
      </div>`;

    const refreshLink = this.domElement.querySelector('#dlswx-refresh') as HTMLElement;
    if (refreshLink) { refreshLink.onclick = (e) => { e.preventDefault(); this._loadAll(); }; }

    this._loadAll();
    if (this._timer) { clearInterval(this._timer); }
    this._timer = setInterval(() => this._loadAll(), REFRESH_MS);
  }

  protected onDispose(): void {
    if (this._timer) { clearInterval(this._timer); this._timer = undefined; }
  }

  // ---- data ----------------------------------------------------------------

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

  private _loadForecast(): void {
    const lat = this._lat(), lon = this._lon();
    fetch('https://api.weather.gov/points/' + lat + ',' + lon, { headers: { 'Accept': 'application/geo+json' } })
      .then(r => r.json())
      .then(pj => {
        const props = (pj && pj.properties) || {};
        if (props.forecastHourly) { this._renderHourly(props.forecastHourly); }
        if (props.forecast) { this._render7day(props.forecast); }
        if (!props.forecast && !props.forecastHourly) { throw new Error('no forecast urls'); }
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
        header: { description: 'Weather panel settings (NOAA/NWS). Radar site code KOHX = local (Nashville/Old Hickory); a regional code like SOUTHEAST or SOUTHMISSVLY shows a wider multi-state view.' },
        groups: [{
          groupName: 'Location',
          groupFields: [
            PropertyPaneTextField('title', { label: 'Panel title' }),
            PropertyPaneTextField('locationLabel', { label: 'Location label (display only)' }),
            PropertyPaneTextField('latitude', { label: 'Latitude' }),
            PropertyPaneTextField('longitude', { label: 'Longitude' }),
            PropertyPaneTextField('radarSite', { label: 'NWS radar code (KOHX local; SOUTHEAST/SOUTHMISSVLY = regional)' })
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

// DARK theme, v1.0.0.5. Full-bleed (98vw) at wide-landscape to line up with the Deed
// Map. Radar capped near native res (crisp). Wide = top row (radar | current | 12h)
// + full-width 7-day card strip. Scoped to .dlswx.
const DLSWX_CSS = `
.dlswx{font-family:'Segoe UI',system-ui,sans-serif;color:#e6e4e2;border:1px solid #3b3a39;border-radius:8px;overflow:hidden;background:#1b1a19;box-shadow:0 1.6px 3.6px rgba(0,0,0,.45);width:100%;box-sizing:border-box}
@media (min-width:1300px) and (orientation:landscape){.dlswx{width:98vw;position:relative;left:50%;margin-left:-49vw;border-radius:10px}}
.dlswx *{box-sizing:border-box}
.dlswx-head{display:flex;justify-content:space-between;align-items:center;padding:10px 16px;background:#0a2c49;color:#fff;border-bottom:1px solid #16334f}
.dlswx-title{font-size:16px;font-weight:600}
.dlswx-loc{font-size:12px;opacity:.85;margin-left:8px}
.dlswx-updated{font-size:11px;opacity:.8}
.dlswx-alerts{padding:8px 16px 0}
.dlswx-alert{padding:7px 11px;border-radius:6px;margin-bottom:6px;font-size:13px;font-weight:600}
.dlswx-alert-sev{background:#3b1416;color:#f5a3a6;border:1px solid #7a2a2e}
.dlswx-alert-mod{background:#3a3209;color:#f3e0a0;border:1px solid #7a6a1e}
.dlswx-alert-h{display:block;font-weight:400;font-size:12px;margin-top:2px;opacity:.9}

/* GRID base = mobile stack */
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
