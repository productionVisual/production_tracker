/**
 * AppSource license gate (v5.4.0).
 *
 * Replaces the previous all-or-nothing model. The visual now always renders a
 * usable free tier; a license unlocks the premium views and lifts the machine
 * cap. See FREE_MACHINE_LIMIT in visual.ts.
 *
 * Two independent signals can unlock the visual:
 *
 *  1. `hasPlan`  — the ACTIVE USER holds an Active/Warning service plan for
 *                  this visual (powerbi.ServicePlanState 1 or 2).
 *  2. `stamp`    — the REPORT carries a license stamp, written into the
 *                  persisted objects bag by a licensed editor.
 *
 * The stamp is what keeps the "viewers are always free" business model intact:
 * a viewer never holds a plan of their own, so without it every consumer of a
 * paying customer's report would be pushed back to the free tier. It also keeps
 * PDF/PowerPoint export correct — REST-API export reports
 * isLicenseUnsupportedEnv, so plan info is unavailable there and only the stamp
 * can tell us the report belongs to a paying customer.
 *
 * The stamp is refreshed, not trusted forever: whenever an editor is in an
 * environment where licensing CAN be evaluated and turns out to have no valid
 * plan, the stamp is cleared. A lapsed subscription therefore falls back to the
 * free tier the next time the report is edited.
 */

import powerbi from "powerbi-visuals-api";
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

/** Partner Center plan IDs that grant the full feature set. */
export const VALID_PLAN_IDS = [
    "tracker_trial",
    "tracker_monthly",
    "tracker_annual",
    "tracker_tenant_annual",
    "tracker_reference_partner"
];

export const APPSOURCE_OFFER_URL =
    "https://appsource.microsoft.com/product/power-bi-visuals/productionvisual.production-tracker?tab=Overview";

/** Objects-bag location of the license stamp. */
const STAMP_OBJECT = "display";
const STAMP_PROPERTY = "licenseStamp";
const STAMP_VALUE = "1";

export class LicenseGate {
    /** The active user holds a usable plan for this visual. */
    private hasPlan = false;
    /** Licensing could actually be evaluated here (signed in, online, supported env). */
    private evaluable = false;
    /** Publish-to-web, embed, national clouds, Report Server, REST-API PDF/PPT export. */
    private unsupportedEnv = false;
    /** The async plan lookup has resolved (guards the first paint). */
    private checked = false;
    /** The report was configured by a licensed editor. */
    private stamp = false;

    private started = false;
    private stampWriteInFlight = false;
    /** Last edit-mode seen by update(), so the async resolver can stamp correctly. */
    private editMode = false;

    constructor(
        private readonly host: IVisualHost,
        private readonly onResolved: () => void
    ) { }

    /** True once the plan lookup has finished — until then, don't advertise the free tier. */
    public get isChecked(): boolean { return this.checked; }

    /** The single question the rest of the visual asks. */
    public get unlocked(): boolean { return this.hasPlan || this.stamp; }

    /** True when this user could buy a license from here (hides dead "Buy" links in export/embed). */
    public get canPurchase(): boolean { return !this.unsupportedEnv; }

    /** Read the stamp out of the report's persisted objects on every update. */
    public syncStamp(dataView: powerbi.DataView | undefined): void {
        const objects = dataView && dataView.metadata && dataView.metadata.objects;
        const bag = objects && objects[STAMP_OBJECT] as Record<string, powerbi.PrimitiveValue> | undefined;
        const raw = bag ? bag[STAMP_PROPERTY] : undefined;
        this.stamp = String(raw || "") === STAMP_VALUE;
    }

    /**
     * Kick off the plan lookup. Runs once per visual lifetime; Power BI caches
     * the result host-side for the session anyway.
     */
    public start(): void {
        if (this.started) { return; }
        this.started = true;

        // tslint:disable-next-line:no-any
        const lm: any = this.host && (this.host as any).licenseManager;
        if (!lm || typeof lm.getAvailableServicePlans !== "function") {
            // API missing entirely (old host / test harness): free tier, no nagging.
            this.checked = true;
            this.evaluable = false;
            return;
        }

        Promise.resolve(lm.getAvailableServicePlans()).then(
            // tslint:disable-next-line:no-any
            (result: any) => {
                const info = result || {};
                this.unsupportedEnv = !!info.isLicenseUnsupportedEnv;
                const infoAvailable = info.isLicenseInfoAvailable !== false;
                this.evaluable = infoAvailable && !this.unsupportedEnv;

                if (this.evaluable) {
                    const plans = info.plans || [];
                    // tslint:disable-next-line:no-any
                    this.hasPlan = plans.some((p: any) => {
                        if (!p) { return false; }
                        // Only Active (1) and Warning (2) are usable licenses. The runtime
                        // hands back the numeric enum; the string forms are belt and braces.
                        const stateOk = p.state === 1 || p.state === 2
                            || p.state === "Active" || p.state === "Warning";
                        // spIdentifier is the Partner Center GENERATED Service ID
                        // ("<publisher>.<offer>.<planId>"), so match on the planId part.
                        const sp = String(p.spIdentifier || "");
                        return stateOk && VALID_PLAN_IDS.some((id) => sp.indexOf(id) >= 0);
                    });
                }
                this.checked = true;
                // Stamp here too, not just from update(): the first update ran
                // before this resolved, so without this a licensed editor's
                // report would stay unstamped until something else redrew it.
                this.reconcileStamp(this.editMode);
                this.onResolved();
            },
            () => {
                // Lookup failed (offline, transient outage): free tier, and never
                // touch the stamp — we can't tell whether the user is licensed.
                this.hasPlan = false;
                this.evaluable = false;
                this.checked = true;
                this.onResolved();
            }
        );
    }

    /**
     * Keep the report's stamp in sync with the editor's actual entitlement.
     * Only ever called for editors, and only where licensing is evaluable —
     * viewers and export renderers must never rewrite the report.
     */
    public reconcileStamp(isEditMode: boolean): void {
        this.editMode = isEditMode;
        if (!isEditMode || !this.checked || !this.evaluable || this.stampWriteInFlight) { return; }
        if (this.hasPlan && !this.stamp) {
            this.writeStamp(STAMP_VALUE);
        } else if (!this.hasPlan && this.stamp) {
            this.writeStamp("");
        }
    }

    private writeStamp(value: string): void {
        this.stampWriteInFlight = true;
        this.stamp = value === STAMP_VALUE;
        this.host.persistProperties({
            merge: [{
                objectName: STAMP_OBJECT,
                selector: undefined,
                properties: { [STAMP_PROPERTY]: value }
            }]
        });
        this.stampWriteInFlight = false;
    }
}
