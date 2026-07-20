/* eslint-disable */
// ============================================================================
// DLS Crew Clock — SPFx client-side web part (framework: none)
//
// GPS-assisted one-tap clock-in/out for the two field crews. Lives on
// Crew-Clock.aspx. Reads today's Field Schedule + a slim WIP cache, pre-selects
// the nearest job, and on DONE creates ONE COMPLETE row in Crew Time Log.
//
// WHY ROWS ARE BORN COMPLETE: Flow 12 triggers on item CREATE ONLY. A
// create-then-patch would roll up the wrong hours (or none). Never split the
// write. -- see DLS-Automation-3-Crew-Time-Log.md
//
// WHY COORDINATES ARE PRECISION-GATED: Flow 9 geocodes some WIP rows from city
// name only and adds +-0.005deg jitter. Those coordinates are ~500 m of fiction;
// a distance claim off them would send a crew to the wrong job. Only
// GeocodePrecision = Parcel/address rows get a distance. City rows still appear
// in the tap list, just with no distance shown.
//
// Conventions follow the Property & Deed Map web part (framework none, plain
// DOM, SPHttpClient with odata=nometadata, digest handled by SPHttpClient).
//
// SEPARATE PACKAGE: this is dls-crew-clock.sppkg. It must never be merged into
// the Deed Map bundle (shared-bundle edits have silently overwritten work here).
// ============================================================================
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

// ============================================================================
// Internal field names — ALL VERIFIED LIVE 2026-07-15 via REST GET on the
// fields endpoints of Crew Time Log, Field Schedule, and WIP Tracking.
// Nothing else in the file hardcodes a field name; if a column is ever renamed,
// fix it here only. A wrong name fails the POST with a 400 and the row is
// queued on the device, not lost.
// ============================================================================
const F = {
  endTime:      'WorkEndTime',                      // DateTime
  totalHours:   'Total_x0020_Crew_x0020_Hours',     // Calculated, read-only
  projectId:    'ProjectId',                        // 'Project' Lookup -> WIP; REST writes the Id field
  crewOnSite:   'CrewOnSite',                       // MultiChoice (no _x0020_ — created without spaces)
  crewChief:    'CrewChief',                        // Choice
  startTime:    'Work_x0020_Date',                  // DateTime — display "Work Date & Start Time"
  fieldComplete:'Field_x0020_Complete',             // Boolean

  // NEW columns added by build step 2 (names chosen at creation — keep in sync):
  entrySource:  'EntrySource',
  distanceMi:   'ClockInDistanceMi',
  autoFlagged:  'AutoFlagged'
};

// Field Schedule — verified live 2026-07-15.
const FS = {
  date:         'Target_x0020_Field_x0020_Date',    // DateTime — display "Target Field Date"
  crew:         'CrewAssignment',                   // MultiChoice
  projectId:    'ProjectId'                         // 'Project' Lookup -> WIP (field id c6ce9502-...)
};

// WIP Tracking — verified live 2026-07-15.
const WIP = {
  label:        'MapLabel',                         // Calculated. Was 'JobLabel', changed 2026-07-16:
                                                    // JobLabel is null on 46 of 209 rows (22%), and
                                                    // they're the newest jobs (260601+) — exactly what
                                                    // crews log against. MapLabel is null on zero rows.
  lat:          'Lat',                              // Number
  lng:          'Lng',                              // Number
  precision:    'GeocodePrecision'                  // Text. LIVE VALUES: 'Parcel' (22), 'Address' (8),
                                                    // null (177 — includes the city-geocoded rows).
                                                    // null fails the PARCEL_PRECISION regex, so the
                                                    // gate excludes them without a special case.
};

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

// Office anchor: 107 Scottsville Road, Lafayette, TN 37083.
// Geocoded once at build time (2026-07-15) via the Azure Maps key in the cost
// estimator workbook — Point Address match, score 0.99998. Hardcoded so the
// runtime never needs the key and the key never enters source control.
const OFFICE_LAT = 36.521489;
const OFFICE_LNG = -86.026388;

const NEAR_MI          = 0.5;   // "at the job" / "at the office" radius
const DOMINANCE        = 2.0;   // nearest must be this many x closer to auto-pick
const AWAY_FROM_JOB_MI = 1.0;   // "you've left the job" -> offer to end
const STALE_HOURS      = 14;    // older than this -> prior-day close-out flow
const MIN_HOURS        = 0.5;   // list validation: end > start; enforce half hour

const LS_SESSION = 'dls.crewclock.session.v1';
const LS_QUEUE   = 'dls.crewclock.queue.v1';
const LS_FS      = 'dls.crewclock.fs.v1';
const LS_WIP     = 'dls.crewclock.wip.v1';
// Set once per device (v1.0.0.6). All four devices share the Fielding@ login, so
// SharePoint 'Created By' reads "Field Crew" on every row and carries no attribution.
// This is the only thing that records WHO tapped, as opposed to who the schedule
// SAID would be there.
const LS_CHIEF   = 'dls.crewclock.chief.v1';

const PARCEL_PRECISION = /^(parcel|address|point|rooftop)/i; // anything but City

// ---------------------------------------------------------------------------
// Pure helpers
// ---------------------------------------------------------------------------

function haversineMi(aLat: number, aLng: number, bLat: number, bLng: number): number {
  const R = 3958.7613; // mean earth radius, miles
  const dLat = (bLat - aLat) * Math.PI / 180;
  const dLng = (bLng - aLng) * Math.PI / 180;
  const s1 = Math.sin(dLat / 2), s2 = Math.sin(dLng / 2);
  const h = s1 * s1 + Math.cos(aLat * Math.PI / 180) * Math.cos(bLat * Math.PI / 180) * s2 * s2;
  return 2 * R * Math.asin(Math.min(1, Math.sqrt(h)));
}

// Round to the nearest :00/:30. The list validates MINUTE = 0 or 30 on BOTH
// time columns; a :15 write is rejected outright. Central is a whole-hour UTC
// offset, so rounding local minutes also lands UTC on 0/30.
function roundHalfHour(d: Date): Date {
  const r = new Date(d.getTime());
  r.setSeconds(0, 0);
  const m = r.getMinutes();
  const snapped = Math.round(m / 30) * 30; // 0, 30, or 60
  r.setMinutes(0);
  return new Date(r.getTime() + snapped * 60000);
}

function hoursBetween(a: Date, b: Date): number {
  return (b.getTime() - a.getTime()) / 3600000;
}

// Force a legal row: end strictly after start, at least MIN_HOURS.
function enforceMinEnd(start: Date, end: Date): Date {
  if (hoursBetween(start, end) >= MIN_HOURS) return end;
  return new Date(start.getTime() + MIN_HOURS * 3600000);
}

function fmtTime(d: Date): string {
  let h = d.getHours(); const m = d.getMinutes();
  const ap = h >= 12 ? 'PM' : 'AM';
  h = h % 12; if (h === 0) h = 12;
  return h + ':' + (m < 10 ? '0' + m : m) + ' ' + ap;
}

function fmtElapsed(ms: number): string {
  const total = Math.max(0, Math.floor(ms / 60000));
  const h = Math.floor(total / 60), m = total % 60;
  return h + 'h ' + (m < 10 ? '0' + m : m) + 'm';
}

function fmtMi(mi: number): string {
  return (mi < 10 ? mi.toFixed(1) : Math.round(mi).toString()) + ' mi';
}

function dayName(d: Date): string {
  return ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][d.getDay()];
}

function isoLocalDayRange(d: Date): { start: string; end: string } {
  const s = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
  const e = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59, 999);
  return { start: s.toISOString(), end: e.toISOString() };
}

function sameLocalDay(a: Date, b: Date): boolean {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function lsGet(key: string): any {
  try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : null; } catch (e) { return null; }
}
function lsSet(key: string, val: any): void {
  try { localStorage.setItem(key, JSON.stringify(val)); } catch (e) { /* quota/private mode — non-fatal */ }
}
function lsDel(key: string): void {
  try { localStorage.removeItem(key); } catch (e) { /* no-op */ }
}

function esc(s: any): string {
  return ('' + (s == null ? '' : s)).replace(/[&<>"']/g, function (c) {
    return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' } as any)[c];
  });
}

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

interface IJob {
  // The FS ROW id, not the project id. Jobs are keyed on this because the office
  // puts two crews on one big job as TWO Field Schedule rows against the same
  // project — keying on wipId collapsed them to the first row, so the second
  // crew's tap silently wrote the first crew's CrewAssignment. 0 = no FS row
  // (the "job not on today's list" search path).
  fsId: number;
  wipId: number;
  label: string;
  lat: number | null;
  lng: number | null;
  precise: boolean;      // true only for Parcel/address precision
  crew: string[];        // from the FS row's CrewAssignment
  distMi: number | null; // null when no fix or not precise
}

interface ISession {
  fsId: number;
  wipId: number;
  label: string;
  crew: string[];
  chief: string;         // snapshot at START — see startedBy()
  startIso: string;
  distMi: number | null;
}

interface IPending {
  body: any;
  label: string;
  queuedIso: string;
}

export interface IDlsCrewClockWebPartProps {
  title: string;
  timeLogListGuid: string;
  fieldScheduleListTitle: string;
  wipListGuid: string;
  // Comma-separated. MUST match the Crew Time Log 'CrewChief' Choice values
  // exactly — a mismatch fails the POST with a 400 and queues the row on the
  // device. Property rather than a constant so the roster tracks the column
  // without a rebuild. Live value 2026-07-20: "Cody, Desmond".
  crewChiefs: string;
}

// ============================================================================
export default class DlsCrewClockWebPart extends BaseClientSideWebPart<IDlsCrewClockWebPartProps> {

  private jobsToday: IJob[] = [];
  private wipAll: any[] = [];
  private session: ISession | null = null;
  private chief: string = '';
  private fix: { lat: number; lng: number } | null = null;
  private geoError: string = '';
  private status: string = '';
  private busy: boolean = false;
  private showAllJobs: boolean = false;
  private jobFilter: string = '';
  private tick: any = null;
  private watchId: any = null;

  // ---- lifecycle ----------------------------------------------------------

  protected onInit(): Promise<void> {
    this.session = lsGet(LS_SESSION);
    this.chief = lsGet(LS_CHIEF) || '';
    this.jobsToday = lsGet(LS_FS) || [];
    this.wipAll = lsGet(LS_WIP) || [];

    window.addEventListener('online', this.onOnline);
    // Crews open this page from the 6:45 AM Teams link onto a tab that may have
    // been loaded days ago. refresh() otherwise runs only at init and on
    // 'online', so a foregrounded page could show a stale schedule.
    document.addEventListener('visibilitychange', this.onVisibility);
    this.startGeo();

    // Live elapsed clock on the active-session view. Skipped while the DONE
    // confirm is up — found live 2026-07-15: the tick re-render wiped the
    // confirm dialog out from under the user mid-decision.
    this.tick = setInterval(() => { if (this.session && !this._confirming) this.render(); }, 30000);

    // No render() here — SPFx calls render() itself once onInit resolves, and
    // domElement is not guaranteed before that. refresh() re-renders on arrival.
    this.refresh();
    return Promise.resolve();
  }

  protected onDispose(): void {
    window.removeEventListener('online', this.onOnline);
    document.removeEventListener('visibilitychange', this.onVisibility);
    if (this.tick) clearInterval(this.tick);
    if (this.watchId != null && navigator.geolocation) navigator.geolocation.clearWatch(this.watchId);
  }

  private onOnline = (): void => { this.flushQueue(); this.refresh(); };

  // refresh() also flushes the offline queue, so a foregrounded page both
  // refetches today's schedule and sends anything saved while out of signal.
  private onVisibility = (): void => { if (document.visibilityState === 'visible') this.refresh(); };

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  // ---- SharePoint plumbing (mirrors the Deed Map web part) ----------------

  private cfg(): any { return (SPHttpClient as any).configurations.v1; }
  private web(): string { return this.context.pageContext.web.absoluteUrl; }
  private timeLogApi(): string { return this.web() + "/_api/web/lists(guid'" + this.properties.timeLogListGuid + "')"; }
  private wipApi(): string { return this.web() + "/_api/web/lists(guid'" + this.properties.wipListGuid + "')"; }
  private fsApi(): string {
    return this.web() + "/_api/web/lists/getbytitle('" + (this.properties.fieldScheduleListTitle || 'Field Schedule').replace(/'/g, "''") + "')";
  }

  private spGet(url: string): Promise<any> {
    return this.context.spHttpClient
      .get(url, this.cfg(), { headers: { Accept: 'application/json;odata=nometadata' } })
      .then((r: any) => { if (r.status >= 200 && r.status < 300) return r.json(); throw new Error('HTTP ' + r.status); });
  }

  private spPost(url: string, body: any, extra?: any): Promise<any> {
    const headers: any = {
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'odata-version': ''
    };
    if (extra) { for (const k in extra) { headers[k] = extra[k]; } }
    return this.context.spHttpClient.post(url, this.cfg(), { headers: headers, body: body ? JSON.stringify(body) : '{}' });
  }

  // ---- crew identity ------------------------------------------------------

  private chiefList(): string[] {
    const raw = '' + (this.properties.crewChiefs || '');
    const out: string[] = [];
    const parts = raw.split(',');
    for (const p of parts) { const n = p.replace(/^\s+|\s+$/g, ''); if (n) out.push(n); }
    return out;
  }

  // Who to attribute a row to. Prefer the session snapshot so a device handed to
  // the other crew mid-job still credits whoever tapped START. Falls back to the
  // device chief for sessions written by 1.0.0.5, which have no chief key.
  private startedBy(): string {
    return (this.session && this.session.chief) ? this.session.chief : this.chief;
  }

  // ---- geolocation --------------------------------------------------------

  private startGeo(): void {
    if (!navigator.geolocation) { this.geoError = 'This device has no location support.'; return; }
    this.watchId = navigator.geolocation.watchPosition(
      (p: any) => {
        this.fix = { lat: p.coords.latitude, lng: p.coords.longitude };
        this.geoError = '';
        this.recomputeDistances();
        this.render();
      },
      (e: any) => {
        this.geoError = e && e.code === 1
          ? 'Location permission is off — tap your job from the list below.'
          : 'No location fix yet — tap your job from the list below.';
        this.render();
      },
      { enableHighAccuracy: true, maximumAge: 30000, timeout: 20000 }
    );
  }

  private recomputeDistances(): void {
    const f = this.fix;
    for (const j of this.jobsToday) {
      j.distMi = (f && j.precise && j.lat != null && j.lng != null)
        ? haversineMi(f.lat, f.lng, j.lat, j.lng)
        : null;
    }
  }

  private milesFromOffice(): number | null {
    return this.fix ? haversineMi(this.fix.lat, this.fix.lng, OFFICE_LAT, OFFICE_LNG) : null;
  }

  // ---- data ---------------------------------------------------------------

  private refresh(): void {
    this.flushQueue();

    const today = new Date();
    const r = isoLocalDayRange(today);
    const fsUrl = this.fsApi() + '/items?$top=200'
      + '&$select=Id,' + FS.projectId + ',' + FS.crew + ',' + FS.date
      + "&$filter=" + FS.date + " ge datetime'" + r.start + "' and " + FS.date + " le datetime'" + r.end + "'";

    const wipUrl = this.wipApi() + '/items?$top=5000'
      + '&$select=Id,' + WIP.label + ',' + WIP.lat + ',' + WIP.lng + ',' + WIP.precision;

    Promise.all([this.spGet(fsUrl), this.spGet(wipUrl)]).then((res: any[]) => {
      const fsRows = (res[0] && res[0].value) || [];
      const wipRows = (res[1] && res[1].value) || [];

      this.wipAll = wipRows;
      lsSet(LS_WIP, wipRows);

      const byId: any = {};
      for (const w of wipRows) byId[w.Id] = w;

      const jobs: IJob[] = [];
      for (const row of fsRows) {
        const pid = row[FS.projectId];
        const w = pid != null ? byId[pid] : null;
        if (!w) continue; // schedule row with no/dead project link — nothing to clock against
        const lat = typeof w[WIP.lat] === 'number' ? w[WIP.lat] : parseFloat(w[WIP.lat]);
        const lng = typeof w[WIP.lng] === 'number' ? w[WIP.lng] : parseFloat(w[WIP.lng]);
        const hasCoord = isFinite(lat) && isFinite(lng);
        jobs.push({
          fsId: row.Id,
          wipId: pid,
          label: w[WIP.label] || ('WIP ' + pid),
          lat: hasCoord ? lat : null,
          lng: hasCoord ? lng : null,
          precise: hasCoord && PARCEL_PRECISION.test('' + (w[WIP.precision] || '')),
          crew: this.asArray(row[FS.crew]),
          distMi: null
        });
      }

      this.jobsToday = jobs;
      lsSet(LS_FS, jobs);
      this.recomputeDistances();
      this.status = '';
      this.render();
    }).catch(() => {
      // Offline or a bad field name. Cached lists (if any) keep the page usable.
      this.status = this.jobsToday.length
        ? 'Showing your saved schedule — no signal.'
        : 'Cannot reach SharePoint and nothing is saved on this device.';
      this.render();
    });
  }

  private asArray(v: any): string[] {
    if (!v) return [];
    if (Object.prototype.toString.call(v) === '[object Array]') return v;
    if (v.results) return v.results;
    return ['' + v];
  }

  // ---- the write ----------------------------------------------------------

  // CrewChief = who actually tapped (device truth). CrewOnSite = what the schedule
  // said. When they disagree — a crew reassigned that morning — BOTH are written and
  // the office sees the disagreement in the list and in Flow 39's email. Deliberately
  // NOT AutoFlagged: that flag means "the time was guessed, go check it", and
  // overloading it with crew mismatch would dilute a signal the office is starting to
  // rely on (decision: Alex, 2026-07-20).
  private buildBody(job: { wipId: number; label: string; crew: string[] }, start: Date, end: Date,
                    fieldComplete: boolean, distMi: number | null, source: string, flagged: boolean,
                    chief: string): any {
    const s = roundHalfHour(start);
    const e = enforceMinEnd(s, roundHalfHour(end));
    const body: any = {};
    body[F.projectId]     = job.wipId;
    body[F.startTime]     = s.toISOString();
    body[F.endTime]       = e.toISOString();
    body[F.fieldComplete] = !!fieldComplete;
    body[F.entrySource]   = source;
    body[F.autoFlagged]   = !!flagged;
    // Flow 39's approval email renders the crew column from CrewChief only —
    // CrewOnSite is an array of {Id,Value} the expression language can't map over
    // (decision (c), 2026-07-16). Before 1.0.0.6 nothing ever wrote this, so every
    // app row showed a blank crew in the daily email.
    if (chief) body[F.crewChief] = chief;
    // MultiChoice under odata=nometadata takes a PLAIN ARRAY — the {results:[...]}
    // wrapper is odata=verbose only and 400s here (verified live 2026-07-15).
    if (job.crew && job.crew.length) body[F.crewOnSite] = job.crew;
    if (distMi != null) body[F.distanceMi] = Math.round(distMi * 100) / 100;
    return body;
  }

  // One complete row, one POST. Never create-then-patch (Flow 12 is create-only).
  private submit(body: any, label: string): void {
    if (this.busy) return;
    this.busy = true;
    this.render();

    const finish = (msg: string) => { this.busy = false; this.status = msg; this.render(); };

    if (!navigator.onLine) { this.queue(body, label); finish('No signal — saved. It will send by itself.'); return; }

    this.spPost(this.timeLogApi() + '/items', body).then((r: any) => {
      if (r.status >= 200 && r.status < 300) {
        this.clearSession();
        finish('Logged ' + label + '. Thanks.');
      } else {
        // 4xx here usually means a field name in F is wrong. Queue rather than
        // lose the entry; the office can see the backlog on the device.
        this.queue(body, label);
        finish('Could not send (' + r.status + ') — saved and will retry.');
      }
    }).catch(() => {
      this.queue(body, label);
      finish('No signal — saved. It will send by itself.');
    });
  }

  private queue(body: any, label: string): void {
    const q: IPending[] = lsGet(LS_QUEUE) || [];
    q.push({ body: body, label: label, queuedIso: new Date().toISOString() });
    lsSet(LS_QUEUE, q);
    this.clearSession();
  }

  private flushQueue(): void {
    const q: IPending[] = lsGet(LS_QUEUE) || [];
    if (!q.length || !navigator.onLine) return;

    const keep: IPending[] = [];
    let done = 0;
    const step = (i: number): void => {
      if (i >= q.length) {
        lsSet(LS_QUEUE, keep);
        if (done) { this.status = 'Sent ' + done + ' saved ' + (done === 1 ? 'entry' : 'entries') + '.'; }
        this.render();
        return;
      }
      this.spPost(this.timeLogApi() + '/items', q[i].body).then((r: any) => {
        if (r.status >= 200 && r.status < 300) done++; else keep.push(q[i]);
        step(i + 1);
      }).catch(() => { keep.push(q[i]); step(i + 1); });
    };
    step(0);
  }

  private clearSession(): void { this.session = null; lsDel(LS_SESSION); }

  // ---- actions ------------------------------------------------------------

  private start(job: IJob): void {
    this.session = {
      fsId: job.fsId, wipId: job.wipId, label: job.label, crew: job.crew,
      chief: this.chief,
      startIso: roundHalfHour(new Date()).toISOString(),
      distMi: job.distMi
    };
    lsSet(LS_SESSION, this.session);
    this.status = '';
    this.render();
  }

  private done(fieldComplete: boolean): void {
    const s = this.session; if (!s) return;
    const body = this.buildBody(s, new Date(s.startIso), new Date(), fieldComplete, s.distMi, 'CrewClockApp', false, this.startedBy());
    this.submit(body, s.label);
  }

  // End an open session at an explicit time (stale close-out / end-now banner).
  private endAt(end: Date, flagged: boolean, source: string): void {
    const s = this.session; if (!s) return;
    const body = this.buildBody(s, new Date(s.startIso), end, false, s.distMi, source, flagged, this.startedBy());
    this.submit(body, s.label);
  }

  private catchUp(job: IJob, day: Date, startHhmm: string, endHhmm: string): void {
    const mk = (hhmm: string) => {
      const p = hhmm.split(':');
      return new Date(day.getFullYear(), day.getMonth(), day.getDate(), parseInt(p[0], 10), parseInt(p[1], 10), 0, 0);
    };
    // No session here, so the device chief is the only identity available.
    const body = this.buildBody(job, mk(startHhmm), mk(endHhmm), false, null, 'CrewClockApp', true, this.chief);
    this.submit(body, job.label);
  }

  // ---- render -------------------------------------------------------------

  private css(): string {
    return '<style>' +
      '.cc{font-family:"Segoe UI",system-ui,sans-serif;max-width:760px;margin:0 auto;padding:12px;color:#f5f5f5;}' +
      '.cc-h{font-size:15px;color:#b9b9c0;margin:0 0 10px;}' +
      '.cc-btn{display:block;width:100%;border:0;border-radius:14px;background:#c2410c;color:#fff;' +
        'font-size:30px;font-weight:700;padding:30px 16px;margin:12px 0;cursor:pointer;line-height:1.25;}' +
      '.cc-btn:active{background:#9a3412;}' +
      '.cc-btn[disabled]{opacity:.55;}' +
      '.cc-btn.sm{font-size:20px;padding:18px 14px;}' +
      '.cc-done{background:#166534;}.cc-done:active{background:#124e28;}' +
      '.cc-alt{background:#33333a;font-size:17px;padding:14px;font-weight:600;}' +
      '.cc-sub{font-size:15px;font-weight:400;opacity:.85;display:block;margin-top:4px;}' +
      '.cc-card{background:#26262a;border:1px solid #33333a;border-radius:12px;padding:14px;margin:12px 0;}' +
      '.cc-ban{background:#722608;border-radius:12px;padding:14px;margin:12px 0;font-size:17px;}' +
      '.cc-el{font-size:44px;font-weight:700;margin:6px 0;}' +
      '.cc-st{font-size:14px;color:#b9b9c0;min-height:20px;margin-top:10px;}' +
      '.cc-warn{color:#fbbf24;font-size:14px;margin:8px 0;}' +
      '.cc-sel{font-size:18px;padding:10px;border-radius:8px;border:1px solid #444;background:#1b1b1f;color:#f5f5f5;margin:4px 6px 4px 0;}' +
      '.cc-in{font-size:17px;padding:10px;width:100%;box-sizing:border-box;border-radius:8px;border:1px solid #444;background:#1b1b1f;color:#f5f5f5;}' +
      '.cc-lnk{color:#fb923c;cursor:pointer;text-decoration:underline;font-size:15px;}' +
      '.cc-badge{background:#722608;border-radius:10px;padding:2px 8px;font-size:13px;margin-left:6px;}' +
      '</style>';
  }

  private halfHourOptions(sel: string): string {
    let out = '';
    for (let h = 5; h <= 21; h++) {
      for (let m = 0; m < 60; m += 30) {
        const v = (h < 10 ? '0' + h : h) + ':' + (m === 0 ? '00' : '30');
        const d = new Date(2000, 0, 1, h, m);
        out += '<option value="' + v + '"' + (v === sel ? ' selected' : '') + '>' + fmtTime(d) + '</option>';
      }
    }
    return out;
  }

  // MUST be public — BaseClientSideWebPart declares render() and TypeScript will not
  // let a subclass narrow its visibility. SPFx calls this itself after onInit resolves.
  public render(): void {
    if (!this.domElement) return;    // an async refresh/tick can land before the first render
    if (this._confirming) return;    // never repaint over the DONE-confirm mid-decision

    // One-time identity gate. Guarded on the roster being configured so a blank
    // crewChiefs property degrades to pre-1.0.0.6 behaviour rather than locking a
    // crew out of a working app. Deploy outside field hours: this draws OVER an
    // active session, which is recoverable (startedBy() falls back) but confusing.
    if (!this.chief && this.chiefList().length) { this.renderChiefGate(); return; }

    const q: IPending[] = lsGet(LS_QUEUE) || [];
    let h = this.css() + '<div class="cc">';

    h += '<p class="cc-h">' + esc(this.properties.title || 'Crew Clock')
      + (q.length ? '<span class="cc-badge">' + q.length + ' saved to send</span>' : '') + '</p>';

    if (this.session) {
      h += this.renderSession();
    } else {
      h += this.renderPicker();
    }

    if (this.status) h += '<div class="cc-st">' + esc(this.status) + '</div>';
    h += '</div>';

    this.domElement.innerHTML = h;
    this.wire();
  }

  private renderChiefGate(): void {
    let h = this.css() + '<div class="cc">'
      + '<p class="cc-h">' + esc(this.properties.title || 'Crew Clock') + '</p>'
      + '<div class="cc-card"><div style="font-size:22px;font-weight:700;">Who is the crew chief on this device?</div>'
      + '<div class="cc-warn">Asked once. You can change it at the bottom of the screen.</div></div>';
    for (const n of this.chiefList()) {
      h += '<button class="cc-btn" data-act="setchief" data-name="' + esc(n) + '">' + esc(n) + '</button>';
    }
    this.domElement.innerHTML = h + '</div>';
    this.wire();
  }

  // Distance and assigned crew, whichever we have. The crew name matters when the
  // office schedules two crews on one job as two FS rows: without it the picker
  // draws two buttons with identical labels and no way to tell them apart.
  private subLabel(j: IJob): string {
    const bits: string[] = [];
    if (j.distMi != null) bits.push(fmtMi(j.distMi) + ' away');
    if (j.crew && j.crew.length) bits.push(j.crew.join(', '));
    return bits.length ? '<span class="cc-sub">' + esc(bits.join(' · ')) + '</span>' : '';
  }

  private renderSession(): string {
    const s = this.session as ISession;
    const start = new Date(s.startIso);
    const now = new Date();
    const age = hoursBetween(start, now);

    // Prior-day / stale: never close silently, never guess a time without asking.
    if (age > STALE_HOURS || !sameLocalDay(start, now)) {
      return '<div class="cc-card">'
        + '<div style="font-size:20px;font-weight:700;margin-bottom:4px;">When did you finish ' + esc(s.label) + ' on ' + dayName(start) + '?</div>'
        + '<div class="cc-warn">Started ' + fmtTime(start) + '. This will be flagged for the office to check.</div>'
        + '<select class="cc-sel" id="cc-stale">' + this.halfHourOptions('16:00') + '</select>'
        + '<button class="cc-btn sm" id="cc-stale-go"' + (this.busy ? ' disabled' : '') + '>Save it</button>'
        + '</div>';
    }

    // Same day, but they're at the office or well away from the job: offer to end.
    const off = this.milesFromOffice();
    // Match on the FS row. Falls back to wipId for sessions written by 1.0.0.5,
    // which have no fsId — without the fallback their banner would never appear.
    const job = s.fsId
      ? this.jobsToday.filter(j => j.fsId === s.fsId)[0]
      : this.jobsToday.filter(j => j.wipId === s.wipId)[0];
    const awayFromJob = job && job.distMi != null && job.distMi > AWAY_FROM_JOB_MI;
    const atOffice = off != null && off <= NEAR_MI;
    let banner = '';
    if (atOffice || awayFromJob) {
      banner = '<div class="cc-ban">Still clocked in on <b>' + esc(s.label) + '</b>'
        + (atOffice ? ' — you\'re back at the office.' : ' — you\'ve left the job.')
        + '<button class="cc-btn sm" id="cc-endnow"' + (this.busy ? ' disabled' : '') + '>End it now</button></div>';
    }

    return banner
      + '<div class="cc-card">'
      + '<div style="font-size:22px;font-weight:700;">' + esc(s.label) + '</div>'
      + '<div style="color:#b9b9c0;">Started ' + fmtTime(start) + '</div>'
      + '<div class="cc-el">' + fmtElapsed(now.getTime() - start.getTime()) + '</div>'
      + '</div>'
      + '<button class="cc-btn cc-done" id="cc-done"' + (this.busy ? ' disabled' : '') + '>DONE</button>';
  }

  private renderPicker(): string {
    // Ranked by distance; jobs without a trustworthy coordinate sort last.
    const ranked = this.jobsToday.slice().sort((a, b) => {
      if (a.distMi == null && b.distMi == null) return 0;
      if (a.distMi == null) return 1;
      if (b.distMi == null) return -1;
      return a.distMi - b.distMi;
    });
    const withDist = ranked.filter(j => j.distMi != null);

    let h = '';

    if (!this.jobsToday.length) {
      h += '<div class="cc-card">Nothing on today\'s schedule.</div>';
      return h + this.renderAllJobs() + this.renderCatchUp();
    }

    // Clear winner: one job today, or near + dominant.
    const clear =
      (ranked.length === 1 && ranked[0].distMi != null && ranked[0].distMi <= NEAR_MI) ||
      (withDist.length === 1 && withDist[0].distMi as number <= NEAR_MI) ||
      (withDist.length > 1 && (withDist[0].distMi as number) <= NEAR_MI &&
        (withDist[1].distMi as number) >= (withDist[0].distMi as number) * DOMINANCE);

    if (clear && withDist.length) {
      const j = withDist[0];
      h += '<button class="cc-btn" data-act="start" data-id="' + j.fsId + '"' + (this.busy ? ' disabled' : '') + '>'
        + 'START — ' + esc(j.label) + this.subLabel(j) + '</button>';
      h += '<div class="cc-st">Not that one? <span class="cc-lnk" data-act="showall">Pick another job</span></div>';
      return h + this.renderCatchUp() + (this.showAllJobs ? this.renderAllJobs() : '');
    }

    // Ambiguous: a short stack, nearest first.
    const stacked = withDist.length > 1 ? withDist.slice(0, 3) : [];
    if (stacked.length) {
      h += '<p class="cc-h">Which job?</p>';
      for (const j of stacked) {
        h += '<button class="cc-btn sm" data-act="start" data-id="' + j.fsId + '"' + (this.busy ? ' disabled' : '') + '>'
          + 'START — ' + esc(j.label) + this.subLabel(j) + '</button>';
      }
    }

    // Everything the branches above didn't render: no fix, City-precision, a measured
    // job too far to win outright, or the tail of a long stack. A distance is shown
    // only where we have one. Found live 2026-07-17: with exactly ONE scheduled job
    // more than NEAR_MI away, `clear` was false and the stack needed two, so the
    // picker drew no job button at all.
    const rest = ranked.filter(j => stacked.indexOf(j) < 0);
    if (rest.length) {
      if (this.geoError) h += '<div class="cc-warn">' + esc(this.geoError) + '</div>';
      h += '<p class="cc-h">Today\'s schedule</p>';
      for (const j of rest) {
        h += '<button class="cc-btn sm cc-alt" data-act="start" data-id="' + j.fsId + '"' + (this.busy ? ' disabled' : '') + '>'
          + esc(j.label) + this.subLabel(j)
          + '</button>';
      }
    }

    h += '<div class="cc-st"><span class="cc-lnk" data-act="showall">Job not on today\'s list?</span></div>';
    return h + this.renderCatchUp() + (this.showAllJobs ? this.renderAllJobs() : '');
  }

  // Searchable full job list from the WIP cache (footer path).
  private renderAllJobs(): string {
    if (!this.showAllJobs) return '';
    const f = this.jobFilter.toLowerCase();
    const hits = this.wipAll
      .filter((w: any) => !f || ('' + (w[WIP.label] || '')).toLowerCase().indexOf(f) >= 0)
      .slice(0, 25);
    let h = '<div class="cc-card"><input class="cc-in" id="cc-filter" placeholder="Search all jobs" value="' + esc(this.jobFilter) + '">';
    for (const w of hits) {
      h += '<button class="cc-btn sm cc-alt" data-act="startwip" data-id="' + w.Id + '">' + esc(w[WIP.label] || ('WIP ' + w.Id)) + '</button>';
    }
    if (!hits.length) h += '<div class="cc-st">No match.</div>';
    return h + '</div>';
  }

  // Missed-visit catch-up: a scheduled job today with no session and no row yet.
  // Covers the no-signal case where nothing got tapped. Three taps, exception only.
  private renderCatchUp(): string {
    if (this.session) return '';
    let h = '<div class="cc-st" style="margin-top:18px;">'
      + '<span class="cc-lnk" data-act="catchup">Forgot to clock a job? Enter it here</span></div>';
    if (this._catchUpPick) h += this.renderCatchUpPick();
    else if (this._catchUpFor) h += this.renderCatchUpCard();
    if (this.chief) {
      h += '<div class="cc-st" style="margin-top:10px;opacity:.7;">Chief: ' + esc(this.chief)
        + ' — <span class="cc-lnk" data-act="clearchief">not you?</span></div>';
    }
    return h;
  }

  private _catchUpFor: IJob | null = null;
  private _catchUpPick: boolean = false;

  // With two crews out, today's schedule has several jobs and the chief has to say
  // which one they missed. Before 1.0.0.6 this card was hardcoded to jobsToday[0].
  private renderCatchUpPick(): string {
    let h = '<div class="cc-card"><div style="font-size:18px;font-weight:700;margin-bottom:6px;">Which job did you miss?</div>';
    for (const j of this.jobsToday) {
      h += '<button class="cc-btn sm cc-alt" data-act="cu-pick" data-id="' + j.fsId + '">'
        + esc(j.label) + this.subLabel(j) + '</button>';
    }
    return h + '<div class="cc-st"><span class="cc-lnk" data-act="cu-cancel">Cancel</span></div></div>';
  }

  private renderCatchUpCard(): string {
    const j = this._catchUpFor as IJob;
    return '<div class="cc-card">'
      + '<div style="font-size:18px;font-weight:700;margin-bottom:6px;">' + esc(j.label) + '</div>'
      + '<select class="cc-sel" id="cc-cu-day"><option value="0">Today</option><option value="1">Yesterday</option></select>'
      + '<div style="margin-top:8px;">From <select class="cc-sel" id="cc-cu-s">' + this.halfHourOptions('06:00') + '</select>'
      + ' to <select class="cc-sel" id="cc-cu-e">' + this.halfHourOptions('16:00') + '</select></div>'
      + '<button class="cc-btn sm" id="cc-cu-go"' + (this.busy ? ' disabled' : '') + '>Save it</button>'
      + '<div class="cc-st"><span class="cc-lnk" data-act="cu-cancel">Cancel</span></div>'
      + '</div>';
  }

  // Delegated listeners — module scope means inline onclick can't see our methods
  // (same constraint the Deed Map web part hit).
  private wire(): void {
    const el = this.domElement;

    const clicks = el.querySelectorAll('[data-act]');
    for (let i = 0; i < clicks.length; i++) {
      clicks[i].addEventListener('click', (ev: any) => {
        const t = ev.currentTarget as HTMLElement;
        const act = t.getAttribute('data-act');
        const id = parseInt(t.getAttribute('data-id') || '0', 10);

        if (act === 'start') {
          const j = this.jobsToday.filter(x => x.fsId === id)[0];
          if (j) this.start(j);
        } else if (act === 'startwip') {
          const w = this.wipAll.filter((x: any) => x.Id === id)[0];
          // No FS row on this path, so no scheduled crew exists. The device chief is
          // the only truth available — write it as a one-name CrewOnSite rather than
          // leaving the row with no crew attribution at all. The CrewChief Choice and
          // CrewOnSite MultiChoice rosters share names, so the value is valid.
          if (w) this.start({
            fsId: 0, wipId: w.Id, label: w[WIP.label] || ('WIP ' + w.Id),
            lat: null, lng: null, precise: false,
            crew: this.chief ? [this.chief] : [], distMi: null
          });
        } else if (act === 'showall') {
          this.showAllJobs = !this.showAllJobs; this.render();
        } else if (act === 'setchief') {
          this.chief = t.getAttribute('data-name') || '';
          lsSet(LS_CHIEF, this.chief);
          this.render();
        } else if (act === 'clearchief') {
          this.chief = ''; lsDel(LS_CHIEF); this.render();
        } else if (act === 'catchup') {
          this._catchUpPick = false;
          if (this.jobsToday.length === 1) { this._catchUpFor = this.jobsToday[0]; }
          else if (this.jobsToday.length > 1) { this._catchUpFor = null; this._catchUpPick = true; }
          else { this._catchUpFor = null; this.showAllJobs = true; this.status = 'Pick the job from the full list first.'; }
          this.render();
        } else if (act === 'cu-pick') {
          this._catchUpFor = this.jobsToday.filter(x => x.fsId === id)[0] || null;
          this._catchUpPick = false; this.render();
        } else if (act === 'cu-cancel') {
          this._catchUpFor = null; this._catchUpPick = false; this.render();
        }
      });
    }

    const doneBtn = el.querySelector('#cc-done');
    if (doneBtn) doneBtn.addEventListener('click', () => this.confirmDone());

    const endNow = el.querySelector('#cc-endnow');
    if (endNow) endNow.addEventListener('click', () => this.endAt(new Date(), false, 'CrewClockApp'));

    const staleGo = el.querySelector('#cc-stale-go');
    if (staleGo) staleGo.addEventListener('click', () => {
      const sel = el.querySelector('#cc-stale') as HTMLSelectElement;
      const s = this.session as ISession; const st = new Date(s.startIso);
      const p = sel.value.split(':');
      const end = new Date(st.getFullYear(), st.getMonth(), st.getDate(), parseInt(p[0], 10), parseInt(p[1], 10));
      this.endAt(end, true, 'AppAutoClose');
    });

    const filter = el.querySelector('#cc-filter') as HTMLInputElement;
    if (filter) filter.addEventListener('input', () => {
      this.jobFilter = filter.value;
      this.render();
      const f2 = this.domElement.querySelector('#cc-filter') as HTMLInputElement;
      if (f2) { f2.focus(); f2.setSelectionRange(f2.value.length, f2.value.length); }
    });

    const cuGo = el.querySelector('#cc-cu-go');
    if (cuGo) cuGo.addEventListener('click', () => {
      const j = this._catchUpFor as IJob;
      const dsel = el.querySelector('#cc-cu-day') as HTMLSelectElement;
      const s = (el.querySelector('#cc-cu-s') as HTMLSelectElement).value;
      const e = (el.querySelector('#cc-cu-e') as HTMLSelectElement).value;
      const d = new Date(); d.setDate(d.getDate() - parseInt(dsel.value, 10));
      this._catchUpFor = null;
      this.catchUp(j, d, s, e);
    });
  }

  private _confirming: boolean = false;

  // The crew's only decision.
  private confirmDone(): void {
    this._confirming = true;
    const s = this.session as ISession;
    this.domElement.innerHTML = this.css() + '<div class="cc">'
      + '<div class="cc-card"><div style="font-size:22px;font-weight:700;">All field work finished on ' + esc(s.label) + '?</div></div>'
      + '<button class="cc-btn cc-alt" id="cc-back">Done for today — coming back</button>'
      + '<button class="cc-btn cc-done" id="cc-complete">Job complete</button>'
      + '<div class="cc-st"><span class="cc-lnk" id="cc-cancel">Cancel</span></div>'
      + '</div>';
    (this.domElement.querySelector('#cc-back') as HTMLElement).addEventListener('click', () => { this._confirming = false; this.done(false); });
    (this.domElement.querySelector('#cc-complete') as HTMLElement).addEventListener('click', () => { this._confirming = false; this.done(true); });
    (this.domElement.querySelector('#cc-cancel') as HTMLElement).addEventListener('click', () => { this._confirming = false; this.render(); });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'DLS Crew Clock' },
        groups: [{
          groupName: 'Lists',
          groupFields: [
            PropertyPaneTextField('title', { label: 'Title' }),
            PropertyPaneTextField('timeLogListGuid', { label: 'Crew Time Log list GUID' }),
            PropertyPaneTextField('fieldScheduleListTitle', { label: 'Field Schedule list title' }),
            PropertyPaneTextField('wipListGuid', { label: 'WIP Tracking list GUID' }),
            PropertyPaneTextField('crewChiefs', { label: 'Crew chief names (comma-separated)' })
          ]
        }]
      }]
    };
  }
}
