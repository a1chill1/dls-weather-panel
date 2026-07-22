/* eslint-disable */
// ============================================================================
// DLS Section Widener — SPFx client-side web part (framework: none)
//
// Renders NOTHING. Its only job: widen the canvas section it is placed in to
// the same edge-to-edge breakout the Property & Deed Map and Weather Panel use
// (98vw / left:50% / margin-left:-49vw, desktop-landscape only), so the native
// List web parts in that section (Inquiries & Quotes | WIP Tracking on
// CollabHome.aspx) get the full monitor width while staying two columns.
//
// WHY THIS EXISTS: CollabHome is a TEAM site — no built-in full-width section
// layout — and List web parts expose no width setting and no editable CSS.
// The single width cap is max-width:1236px on the section's
// [data-automation-id="CanvasZone-SectionContainer"]; every ancestor above it
// is already full page width with overflow-x:visible. Override that one
// element and the two xl6 columns re-split 50/50 across the widened row.
// Full live-inspected DOM: _handoffs/2026-07-22-homepage-iq-wip-full-width.md
//
// SCOPE GUARANTEE: it reaches UP via closest(), so it can only ever affect the
// one section it is physically placed in. No page-wide selectors, no hashed
// class names (the r_..._ classes regenerate on redeploy/theme change).
//
// ROLLBACK: delete this web part from the page (or restore the page's prior
// version) and the section reverts to native layout instantly — the styles are
// inline on the section element and applied only while this web part runs.
//
// SEPARATE PACKAGE: dls-section-widener.sppkg. Never merge into the Deed Map
// or Crew Clock bundles (shared-bundle edits have silently overwritten work).
// ============================================================================
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

export interface IDlsSectionWidenerWebPartProps {}

export default class DlsSectionWidenerWebPart extends BaseClientSideWebPart<IDlsSectionWidenerWebPartProps> {

  private _observer: MutationObserver | undefined;

  public render(): void {
    // In edit mode show a small label so the (otherwise invisible) part can be
    // found, selected, and deleted. In view mode: nothing at all.
    this.domElement.innerHTML = this.displayMode === 2 /* DisplayMode.Edit */
      ? '<div style="padding:6px 10px;border:1px dashed #c2410c;color:#c2410c;font:12px \'Segoe UI\',sans-serif;">DLS Section Widener — widens this section to full page width on desktop. Renders nothing when published. Delete me to revert.</div>'
      : '';

    const apply = (): void => {
      const sec = this.domElement.closest('[data-automation-id="CanvasZone-SectionContainer"]') as HTMLElement | null;
      if (!sec) { return; }
      // Match the Deed Map breakout exactly so both rows hit the same edges.
      // 98vw (not 100vw) dodges the vertical scrollbar width — a flat 100vw
      // produces a horizontal scrollbar.
      if (window.matchMedia('(min-width:1300px) and (orientation:landscape)').matches) {
        sec.style.setProperty('max-width', '98vw', 'important');
        sec.style.setProperty('width', '98vw', 'important');
        sec.style.setProperty('position', 'relative', 'important');
        sec.style.setProperty('left', '50%', 'important');
        sec.style.setProperty('margin-left', '-49vw', 'important');
      } else {
        // Below 1300px the sm12 columns are meant to stack — hand back the
        // native contained layout untouched.
        ['max-width', 'width', 'position', 'left', 'margin-left'].forEach(p => sec.style.removeProperty(p));
      }
    };

    apply();

    // The modern canvas can re-render and strip inline styles; re-apply if so.
    const sec = this.domElement.closest('[data-automation-id="CanvasZone-SectionContainer"]');
    if (sec && !this._observer) {
      this._observer = new MutationObserver(apply);
      this._observer.observe(sec, { attributes: true, attributeFilter: ['style'] });
    }
    // vw units recompute themselves on resize; only the 1300px media-query
    // boundary needs a JS re-check.
    window.addEventListener('resize', apply);
  }

  protected onDispose(): void {
    if (this._observer) { this._observer.disconnect(); this._observer = undefined; }
    const sec = this.domElement.closest('[data-automation-id="CanvasZone-SectionContainer"]') as HTMLElement | null;
    if (sec) {
      ['max-width', 'width', 'position', 'left', 'margin-left'].forEach(p => sec.style.removeProperty(p));
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
