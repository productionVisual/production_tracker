/**
 * Production Tracker — clean rebuild (cert-safe), styled to match the original
 * GitHub-dark look: top bar with filters + view tabs, a glowing production
 * overview chart grouped by area, Table / Tiles / OEE views, and a detail modal.
 *
 * No innerHTML / eval / fetch — the DOM is built with createElement /
 * createElementNS, styling lives in style/visual.less, and the Rendering Events
 * API, selection manager and tooltip service are used throughout.
 */
import "../style/visual.less";
import powerbi from "powerbi-visuals-api";
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ITooltipService = powerbi.extensibility.ITooltipService;
import IVisualEventService = powerbi.extensibility.IVisualEventService;

import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { TrackerSettingsModel } from "./settings";
import {
    transform, ProcessedData, FilterState, EMPTY_FILTER, fillColor, gaugeColor,
    MachineMetrics, DowntimeEntry
} from "./dataModel";
import { t, Language } from "./i18n";
import {
    SHIFT_TEMPLATES, PLANNED_STOPS, ShiftConfig, templateLabel, stopLabel,
    effectiveShiftMinutes, plannedStopTotal, scheduleParamsFor
} from "./shiftConfig";
import { ScheduleParams } from "./dataModel";

type ViewMode = "table" | "tile" | "oee";
const PLOT_H = 132;   // bar plot height in px
const BAR_W = 26;     // bar width in px
const CARD_W = 54;    // bar column width in px

// v5.3.7: AppSource license manager — strict model (Desktop + Service).
// Edit mode without an active license = editor UI locked + trial banner.
// View mode = always rendered, no license check.
const VALID_PLAN_IDS = [
    "tracker_trial",
    "tracker_monthly",
    "tracker_annual",
    "tracker_tenant_annual",
    "tracker_reference_partner"
];
const APPSOURCE_OFFER_URL =
    "https://appsource.microsoft.com/product/power-bi-visuals/productionvisual.production-tracker?tab=Overview";

export class Visual implements IVisual {
    private readonly host: IVisualHost;
    private readonly root: HTMLElement;
    private readonly selectionManager: ISelectionManager;
    private readonly tooltipService: ITooltipService;
    private readonly events: IVisualEventService;
    private readonly formattingService: FormattingSettingsService;

    private settings: TrackerSettingsModel = new TrackerSettingsModel();
    private filter: FilterState = { ...EMPTY_FILTER };
    private data: ProcessedData | undefined;
    private lastOptions: VisualUpdateOptions | undefined;
    private theme: "dark" | "light" = "dark";
    private collapsed = false;
    private showControls = false;
    private isEditMode = false;  // v5.3.6: only editors (Power BI Edit mode) see the shift-settings panel
    // v5.3.7: license state. licenseChecked stays false until the async API resolves
    // so we don't show a "not licensed" banner during the first paint.
    private hasValidLicense = false;
    private licenseChecked = false;
    private licensePromise: Promise<void> | undefined;
    private view: ViewMode = "table";
    private oeeMachine = "";   // "" = all machines
    private shiftConfig: ShiftConfig | undefined;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.root = options.element;
        this.selectionManager = this.host.createSelectionManager();
        this.tooltipService = this.host.tooltipService;
        this.events = this.host.eventService;
        this.formattingService = new FormattingSettingsService();
        this.root.classList.add("pt-root");
        // v5.3.10: visual-level right-click context menu fallback (AppSource cert 1180.2.5).
        // Machine-level handlers in wireInteractions() also show a per-machine menu via
        // ev.preventDefault(); when that fires, defaultPrevented is true here and we skip.
        // On empty areas / non-machine elements, this fallback shows Power BI's visual-level
        // context menu (Export data, Show as table, Spotlight, etc.).
        this.root.addEventListener("contextmenu", (ev: MouseEvent) => {
            if (ev.defaultPrevented) { return; }
            ev.preventDefault();
            this.selectionManager.showContextMenu({}, { x: ev.clientX, y: ev.clientY });
        });
    }

    public update(options: VisualUpdateOptions): void {
        this.events.renderingStarted(options);
        try {
            // v5.3.6: viewMode 1 = Edit, 2 = InFocusEdit, 0 = View (consumer).
            // Only editors get to open the shift-settings panel and change values.
            this.isEditMode = !!(options && (options.viewMode === 1 || options.viewMode === 2));
            this.startLicenseCheck();
            const dataView = options.dataViews && options.dataViews[0];
            this.settings = this.formattingService.populateFormattingSettingsModel(TrackerSettingsModel, dataView);
            this.syncShiftConfig(dataView);
            this.applyConfigToSettings();
            this.lastOptions = options;
            this.data = transform(dataView, this.settings, this.host, this.filter, this.scheduleParams());
            this.render();
            this.events.renderingFinished(options);
        } catch (error) {
            this.events.renderingFailed(options, error instanceof Error ? error.message : String(error));
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingService.buildFormattingModel(this.settings);
    }

    // ----- rendering ----------------------------------------------------------

    /** Re-run the transform with the current filter state and repaint. */
    private refilter(): void {
        if (!this.lastOptions) { return; }
        const dataView = this.lastOptions.dataViews && this.lastOptions.dataViews[0];
        this.data = transform(dataView, this.settings, this.host, this.filter, this.scheduleParams());
        this.render();
    }

    private scheduleParams(): ScheduleParams | undefined {
        // v5.3.9: In Edit mode without an active license, ignore any saved
        // custom shift config and fall back to defaults. The saved config
        // stays in persistProperties (not wiped) — it just isn't APPLIED
        // until the editor re-licenses. Closes the "trial -> configure ->
        // expire -> free forever" loophole. View mode (viewers) always
        // sees whatever was published, preserving the viewer-free model.
        if (this.isEditMode && !this.hasValidLicense) {
            return undefined;
        }
        return this.shiftConfig ? scheduleParamsFor(this.shiftConfig) : undefined;
    }

    // v5.3.7: AppSource license check. Runs once per session; result cached for
    // the lifetime of the visual. Failures default to "no license" (strict).
    private startLicenseCheck(): void {
        if (this.licensePromise) { return; }
        // tslint:disable-next-line:no-any
        const lm: any = this.host && (this.host as any).licenseManager;
        if (!lm || typeof lm.getAvailableServicePlans !== "function") {
            this.licenseChecked = true;
            this.hasValidLicense = false;
            return;
        }
        this.licensePromise = Promise.resolve(lm.getAvailableServicePlans()).then(
            // tslint:disable-next-line:no-any
            (result: any) => {
                const plans = (result && result.plans) || [];
                this.hasValidLicense = plans.some(
                    // tslint:disable-next-line:no-any
                    (p: any) => {
                        if (!p) { return false; }
                        // state is a numeric enum at runtime (Active=1, Warning=2) — both are usable licenses.
                        const stateOk = p.state === 1 || p.state === 2 || p.state === "Active" || p.state === "Warning";
                        // spIdentifier is the Partner Center GENERATED Service ID
                        // (e.g. "<publisher>.<offer>.<planId>"), not the bare plan ID — match on the planId part.
                        const sp = String(p.spIdentifier || "");
                        return stateOk && VALID_PLAN_IDS.some((id) => sp.indexOf(id) >= 0);
                    }
                );
                this.licenseChecked = true;
                if (this.lastOptions) { this.render(); }
            },
            () => {
                this.hasValidLicense = false;
                this.licenseChecked = true;
                if (this.lastOptions) { this.render(); }
            }
        );
    }

    /** Editors can change settings only if they have an active license. */
    private canEdit(): boolean {
        return this.isEditMode && this.hasValidLicense;
    }

    /** Banner shown in Edit mode when no active license is found. */
    private buildLicenseBanner(lang: Language): HTMLElement {
        const banner = el("div", "pt-license-banner");
        banner.style.cssText =
            "display:flex;align-items:center;gap:12px;padding:10px 14px;" +
            "background:linear-gradient(90deg,#2a1f08,#1a1408);" +
            "border:1px solid #6b4a1a;border-radius:6px;margin:8px;color:#ffd87a;font-size:13px;";
        const icon = el("span");
        icon.style.cssText = "font-size:18px;flex-shrink:0;";
        icon.textContent = "🔒";
        banner.appendChild(icon);
        const text = el("span");
        text.style.cssText = "flex:1;";
        const strong = el("strong");
        strong.style.cssText = "color:#ffe9a8;margin-right:6px;";
        strong.textContent = t(lang, "licenseRequired") + ".";
        text.appendChild(strong);
        text.appendChild(document.createTextNode(" " + t(lang, "licenseDescription")));
        banner.appendChild(text);
        const link = document.createElement("a");
        link.href = APPSOURCE_OFFER_URL;
        link.target = "_blank";
        link.rel = "noopener noreferrer";
        link.textContent = t(lang, "startTrial");
        link.style.cssText =
            "flex-shrink:0;padding:6px 14px;background:#ffb84d;color:#1a1408;" +
            "border-radius:4px;text-decoration:none;font-weight:600;";
        banner.appendChild(link);
        return banner;
    }

    private render(): void {
        clear(this.root);
        this.root.classList.toggle("pt-light", this.theme === "light");
        const lang = this.settings.languageCode;
        const data = this.data;
        if (!data || !data.hasData) {
            const empty = el("div", "pt-empty");
            empty.textContent = t(lang, "noData");
            this.root.appendChild(empty);
            return;
        }
        this.root.appendChild(this.buildTopbar(data, lang));
        if (this.isEditMode && this.licenseChecked && !this.hasValidLicense) {
            this.root.appendChild(this.buildLicenseBanner(lang));
        }
        if (this.canEdit() && this.showControls) {
            this.root.appendChild(this.buildControls(data, lang));
        }
        if (!this.collapsed) {
            this.root.appendChild(this.buildChart(data, lang, 1, false));
        }
        const main = el("div", "pt-main");
        if (this.view === "table") { main.appendChild(this.buildTable(data, lang)); }
        else if (this.view === "tile") { main.appendChild(this.buildTiles(data, lang)); }
        else { main.appendChild(this.buildOee(data, lang)); }
        this.root.appendChild(main);
    }

    // ----- top bar ------------------------------------------------------------

    private buildTopbar(data: ProcessedData, lang: Language): HTMLElement {
        const bar = el("div", "pt-topbar");
        // v5.3.5: brand logo before the title (inline SVG data URI)
        const logo = document.createElement("img");
        logo.src = "data:image/svg+xml;base64,PHN2ZyB2aWV3Qm94PSIwIDAgMjU2IDI1NiIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KICA8ZGVmcz4KICAgIDxsaW5lYXJHcmFkaWVudCBpZD0iZyIgeDE9IjAiIHkxPSIwIiB4Mj0iMSIgeTI9IjEiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiM0YWEzZmYiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIuNSIgc3RvcC1jb2xvcj0iIzM2ZDZlNyIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiM3YzVjZmYiLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImJnIiB4MT0iMCIgeTE9IjAiIHgyPSIwIiB5Mj0iMSI+CiAgICAgIDxzdG9wIG9mZnNldD0iMCIgc3RvcC1jb2xvcj0iIzEwMTYyYiIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwODBiMTIiLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgPC9kZWZzPgogIDxyZWN0IHdpZHRoPSIyNTYiIGhlaWdodD0iMjU2IiByeD0iNTYiIGZpbGw9InVybCgjYmcpIi8+CiAgPCEtLSBPRUUgZ2F1Z2UgcmluZyAoMy80IGNpcmNsZSkgLS0+CiAgPGNpcmNsZSBjeD0iMTI4IiBjeT0iMTI4IiByPSI3MiIgc3Ryb2tlPSIjMWMyNDM4IiBzdHJva2Utd2lkdGg9IjIyIiBmaWxsPSJub25lIgogICAgICAgICAgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtZGFzaGFycmF5PSIzMzkgMTAwMCIgdHJhbnNmb3JtPSJyb3RhdGUoMTM1IDEyOCAxMjgpIi8+CiAgPCEtLSBmaWxsZWQgcG9ydGlvbiAofjc4JSkgLS0+CiAgPGNpcmNsZSBjeD0iMTI4IiBjeT0iMTI4IiByPSI3MiIgc3Ryb2tlPSJ1cmwoI2cpIiBzdHJva2Utd2lkdGg9IjIyIiBmaWxsPSJub25lIgogICAgICAgICAgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtZGFzaGFycmF5PSIyNjUgMTAwMCIgdHJhbnNmb3JtPSJyb3RhdGUoMTM1IDEyOCAxMjgpIi8+CiAgPCEtLSBjZW50ZXIgcHJvZHVjdGlvbiBiYXJzIC0tPgogIDxyZWN0IHg9IjEwNiIgeT0iMTE4IiB3aWR0aD0iMTQiIGhlaWdodD0iMzYiIHJ4PSIzIiBmaWxsPSJ1cmwoI2cpIiBvcGFjaXR5PSIuNyIvPgogIDxyZWN0IHg9IjEyNCIgeT0iMTA0IiB3aWR0aD0iMTQiIGhlaWdodD0iNTAiIHJ4PSIzIiBmaWxsPSJ1cmwoI2cpIi8+CiAgPHJlY3QgeD0iMTQyIiB5PSIxMjQiIHdpZHRoPSIxNCIgaGVpZ2h0PSIzMCIgcng9IjMiIGZpbGw9InVybCgjZykiIG9wYWNpdHk9Ii43Ii8+Cjwvc3ZnPgo=";
        logo.width = 22; logo.height = 22;
        logo.style.cssText = "flex-shrink:0;display:block;margin-right:8px;";
        bar.appendChild(logo);
        const title = el("span", "pt-title");
        title.textContent = t(lang, "title");
        bar.appendChild(title);

        const filters = el("div", "pt-filters");
        filters.appendChild(this.filterSelect(t(lang, "allShifts"), data.filters.shifts, this.filter.shift,
            v => { this.filter.shift = v; this.refilter(); }));
        filters.appendChild(this.filterSelect(t(lang, "allLines"), data.filters.lines, this.filter.line,
            v => { this.filter.line = v; this.refilter(); }));
        filters.appendChild(this.filterSelect(t(lang, "allLocations"), data.filters.locations, this.filter.location,
            v => { this.filter.location = v; this.refilter(); }));
        filters.appendChild(this.filterSelect(t(lang, "allProducts"), data.filters.products, this.filter.product,
            v => { this.filter.product = v; this.refilter(); }));
        filters.appendChild(this.filterSelect(t(lang, "allCategories"), data.filters.categories, this.filter.category,
            v => { this.filter.category = v; this.refilter(); }));
        filters.appendChild(this.dateInput(t(lang, "from"), this.filter.dateFrom,
            v => { this.filter.dateFrom = v; this.refilter(); }));
        filters.appendChild(this.dateInput(t(lang, "to"), this.filter.dateTo,
            v => { this.filter.dateTo = v; this.refilter(); }));
        bar.appendChild(filters);

        const controls = el("div", "pt-controls");
        const tab = (label: string, mode: ViewMode) =>
            this.button(label, this.view === mode, () => { this.view = mode; this.render(); });
        controls.appendChild(tab(t(lang, "tableView"), "table"));
        controls.appendChild(tab(t(lang, "tileView"), "tile"));
        const oeeTab = this.button(t(lang, "oeeView"), this.view === "oee",
            () => { this.view = "oee"; this.render(); });
        oeeTab.classList.add("pt-btn-purple");
        controls.appendChild(oeeTab);

        const reportButton = this.button(t(lang, "report"), false, () => this.showReport(data, lang));
        reportButton.classList.add("pt-btn-purple");
        controls.appendChild(reportButton);

        // v5.3.7: Settings button only for licensed editors. Viewers and
        // unlicensed editors see everything else but can't change settings.
        if (this.canEdit()) {
            controls.appendChild(this.button(t(lang, "settings"), this.showControls, () => {
                this.showControls = !this.showControls; this.render();
            }));
        }
        controls.appendChild(this.button(this.collapsed ? t(lang, "expand") : t(lang, "collapse"), false, () => {
            this.collapsed = !this.collapsed; this.render();
        }));
        controls.appendChild(this.button(this.theme === "light" ? t(lang, "dark") : t(lang, "light"), false, () => {
            this.theme = this.theme === "light" ? "dark" : "light"; this.render();
        }));
        const langButton = this.button(lang === "hu" ? "HU" : "EN", false, () => {
            this.settings.display.language.value =
                lang === "hu" ? { value: "en", displayName: "English" } : { value: "hu", displayName: "Magyar" };
            this.persist("language", this.settings.display.language.value.value);
            this.render();
        });
        langButton.title = "Language";
        controls.appendChild(langButton);
        bar.appendChild(controls);
        return bar;
    }

    private button(label: string, active: boolean, onClick: () => void): HTMLElement {
        const b = el("button", active ? "pt-btn pt-btn-active" : "pt-btn");
        b.setAttribute("type", "button");
        b.textContent = label;
        b.addEventListener("click", onClick);
        return b;
    }

    private dateInput(label: string, value: string, onChange: (value: string) => void): HTMLInputElement {
        const input = document.createElement("input");
        input.type = "date";
        input.className = "pt-date";
        input.value = value || "";
        input.title = label;
        input.addEventListener("change", () => onChange(input.value));
        return input;
    }

    private filterSelect(allLabel: string, options: string[], current: string,
                         onChange: (value: string) => void): HTMLSelectElement {
        const select = document.createElement("select");
        select.className = "pt-select";
        const addOption = (value: string, label: string) => {
            const opt = document.createElement("option");
            opt.value = value; opt.textContent = label;
            if (value === current) { opt.selected = true; }
            select.appendChild(opt);
        };
        addOption("", allLabel);
        for (const o of options) { addOption(o, o); }
        select.addEventListener("change", () => onChange(select.value));
        return select;
    }

    // ----- shift settings panel ----------------------------------------------

    /** Load the persisted shift config, or derive it from the current settings. */
    private syncShiftConfig(dataView: powerbi.DataView | undefined): void {
        const objects = dataView && dataView.metadata && dataView.metadata.objects;
        const raw = objects && objects.shiftSchedule
            ? (objects.shiftSchedule as Record<string, powerbi.PrimitiveValue>).shiftConfig
            : undefined;
        if (typeof raw === "string" && raw) {
            try {
                const parsed = JSON.parse(raw) as ShiftConfig;
                if (parsed && Array.isArray(parsed.stops) && typeof parsed.templateIdx === "number") {
                    this.shiftConfig = {
                        templateIdx: parsed.templateIdx,
                        customMinutes: parsed.customMinutes || 480,
                        stops: PLANNED_STOPS.map((_, i) => Number(parsed.stops[i]) || 0)
                    };
                    return;
                }
            } catch { /* fall through to derive */ }
        }
        if (this.shiftConfig) { return; }
        // derive from the current (persisted) simple settings — no numeric change on upgrade
        const sched = this.settings.shiftSchedule;
        const shiftLen = sched.shiftLengthMinutes.value || 480;
        const perDay = sched.shiftsPerDay.value || 3;
        const stopTotal = sched.plannedStopMinutes.value || 0;
        let idx = SHIFT_TEMPLATES.findIndex(tpl => !tpl.custom && tpl.minutes === shiftLen && tpl.shiftsPerDay === perDay);
        if (idx < 0) { idx = SHIFT_TEMPLATES.length - 1; }   // Custom
        const stops = PLANNED_STOPS.map(() => 0);
        stops[0] = stopTotal;
        this.shiftConfig = { templateIdx: idx, customMinutes: shiftLen, stops };
    }

    /** Push the effective config values into the settings the transform reads. */
    private applyConfigToSettings(): void {
        const cfg = this.shiftConfig;
        if (!cfg) { return; }
        const tpl = SHIFT_TEMPLATES[cfg.templateIdx] || SHIFT_TEMPLATES[0];
        const sched = this.settings.shiftSchedule;
        sched.shiftLengthMinutes.value = effectiveShiftMinutes(cfg);
        sched.plannedStopMinutes.value = plannedStopTotal(cfg);
        sched.shiftsPerDay.value = tpl.shiftsPerDay;
        sched.workingDaysPerWeek.value = tpl.daysOn;
    }

    private commitShiftConfig(): void {
        this.applyConfigToSettings();
        this.host.persistProperties({
            merge: [{
                objectName: "shiftSchedule", selector: undefined,
                properties: { shiftConfig: JSON.stringify(this.shiftConfig) }
            }]
        });
        this.refilter();
    }

    private buildControls(data: ProcessedData, lang: Language): HTMLElement {
        const cfg = this.shiftConfig || { templateIdx: 0, customMinutes: 480, stops: PLANNED_STOPS.map(() => 0) };
        const tpl = SHIFT_TEMPLATES[cfg.templateIdx] || SHIFT_TEMPLATES[0];
        const panel = el("div", "pt-cpanel");

        // ---- shift schedule group (template dropdown + custom minutes) ----
        const g1 = this.controlGroup(t(lang, "shiftSchedule"));
        const tplSelect = document.createElement("select");
        tplSelect.className = "pt-select pt-select-wide";
        SHIFT_TEMPLATES.forEach((template, i) => {
            const opt = document.createElement("option");
            opt.value = String(i); opt.textContent = templateLabel(template, lang);
            if (i === cfg.templateIdx) { opt.selected = true; }
            tplSelect.appendChild(opt);
        });
        tplSelect.addEventListener("change", () => {
            cfg.templateIdx = Number(tplSelect.value);
            this.shiftConfig = cfg;
            this.commitShiftConfig();
        });
        g1.appendChild(tplSelect);
        if (tpl.custom) {
            g1.appendChild(this.numInput(t(lang, "custom"), cfg.customMinutes, 0, 1440, v => {
                cfg.customMinutes = v; this.shiftConfig = cfg; this.commitShiftConfig();
            }));
        }
        const eff = effectiveShiftMinutes(cfg);
        const info = el("div", "pt-cg-info");
        info.textContent = t(lang, "shiftCount") + ": " + data.totals.shiftSlots
            + " (" + tpl.daysOn + " " + t(lang, "days") + " × " + tpl.shiftsPerDay + " " + t(lang, "shifts") + ")";
        g1.appendChild(info);
        const lenInfo = el("div", "pt-cg-info");
        lenInfo.textContent = t(lang, "shiftLength") + ": " + eff + " " + t(lang, "minutes");
        g1.appendChild(lenInfo);
        panel.appendChild(g1);

        // ---- planned stops group (named categories) ----
        const g2 = this.controlGroup(t(lang, "plannedStops"));
        const grid = el("div", "pt-stop-grid");
        PLANNED_STOPS.forEach((stop, i) => {
            const label = el("span", "pt-num-label"); label.textContent = stopLabel(stop, lang);
            const input = document.createElement("input");
            input.type = "number"; input.className = "pt-num-input";
            input.min = "0"; input.max = "999"; input.value = String(cfg.stops[i] || 0);
            input.addEventListener("change", () => {
                let v = Number(input.value); if (isNaN(v)) { v = 0; }
                v = Math.max(0, Math.min(999, v)); input.value = String(v);
                cfg.stops[i] = v; this.shiftConfig = cfg; this.commitShiftConfig();
            });
            grid.appendChild(label); grid.appendChild(input);
        });
        g2.appendChild(grid);
        g2.appendChild(this.summaryLine(t(lang, "plannedStopTotal"),
            plannedStopTotal(cfg) + " " + t(lang, "minutes"), "#f0883e"));
        panel.appendChild(g2);

        // ---- summary group ----
        const stopTot = plannedStopTotal(cfg);
        const slots = data.totals.shiftSlots || 0;
        const baseOpen = Math.max(0, eff - stopTot);
        const g3 = this.controlGroup(t(lang, "summary"));
        g3.appendChild(this.summaryLine(t(lang, "shiftTotal"),
            eff + " × " + slots + " = " + fmt(eff * slots) + " " + t(lang, "minutes"), "#79c0ff"));
        g3.appendChild(this.summaryLine(t(lang, "plannedStopTotal"),
            stopTot + " × " + slots + " = " + fmt(stopTot * slots) + " " + t(lang, "minutes"), "#f0883e"));
        g3.appendChild(this.summaryLine(t(lang, "baseOpenTime") + " / " + t(lang, "shifts"),
            baseOpen + " " + t(lang, "minutes"), "#56d364"));
        g3.appendChild(this.summaryLine(t(lang, "openTime"),
            fmt(data.totals.openTime) + " " + t(lang, "minutes"), "#56d364"));
        panel.appendChild(g3);
        return panel;
    }

    private controlGroup(label: string): HTMLElement {
        const group = el("div", "pt-cg");
        const l = el("div", "pt-cg-label"); l.textContent = label;
        group.appendChild(l);
        return group;
    }

    private numInput(label: string, value: number, min: number, max: number,
                     onChange: (value: number) => void): HTMLElement {
        const row = el("div", "pt-num-row");
        const l = el("span", "pt-num-label"); l.textContent = label;
        const input = document.createElement("input");
        input.type = "number";
        input.className = "pt-num-input";
        input.min = String(min); input.max = String(max);
        input.value = String(value);
        input.addEventListener("change", () => {
            let v = Number(input.value);
            if (isNaN(v)) { v = min; }
            v = Math.max(min, Math.min(max, v));
            input.value = String(v);
            onChange(v);
        });
        row.appendChild(l); row.appendChild(input);
        return row;
    }

    private summaryLine(label: string, value: string, color: string): HTMLElement {
        const row = el("div", "pt-sum-row");
        const l = el("span", "pt-sum-label"); l.textContent = label + ":";
        const v = el("span", "pt-sum-value"); v.textContent = value; v.style.color = color;
        row.appendChild(l); row.appendChild(v);
        return row;
    }

    // ----- production overview chart -----------------------------------------

    private buildChart(data: ProcessedData, lang: Language, scale: number, expanded: boolean): HTMLElement {
        const chart = el("div", expanded ? "pt-chart pt-chart-expanded" : "pt-chart");

        const head = el("div", "pt-chart-head");
        const ttl = el("div", "pt-chart-title");
        ttl.textContent = t(lang, "productionOverview");
        head.appendChild(ttl);
        if (!expanded) {
            const expandBtn = this.button("⛶", false, () => this.showChartPopup(data, lang));
            expandBtn.title = t(lang, "expand");
            head.appendChild(expandBtn);
        }
        head.appendChild(this.legend(lang));
        chart.appendChild(head);

        const scroll = el("div", "pt-chart-scroll");
        let maxVal = 1;
        for (const m of data.machines) { maxVal = Math.max(maxVal, m.produced, m.dailyTarget); }

        for (const area of data.areas) {
            const group = el("div", "pt-chart-group");
            const label = el("div", "pt-group-label");
            label.textContent = area.name;
            label.style.fontSize = Math.round(10 * scale) + "px";
            group.appendChild(label);
            const bars = el("div", "pt-group-bars");
            for (const machine of area.machines) {
                bars.appendChild(this.buildBar(machine, maxVal, lang, scale));
            }
            group.appendChild(bars);
            scroll.appendChild(group);
        }
        chart.appendChild(scroll);
        return chart;
    }

    /** Full-screen overlay showing the production overview chart at a larger scale. */
    private showChartPopup(data: ProcessedData, lang: Language): void {
        const overlay = el("div", "pt-overlay");
        overlay.addEventListener("click", (ev) => { if (ev.target === overlay) { overlay.remove(); } });
        const box = el("div", "pt-chart-popup");
        const close = el("button", "pt-close");
        close.setAttribute("type", "button"); close.textContent = "✕";
        close.addEventListener("click", () => overlay.remove());
        box.appendChild(close);
        box.appendChild(this.buildChart(data, lang, 1.9, true));
        overlay.appendChild(box);
        this.root.appendChild(overlay);
    }

    private legend(lang: Language): HTMLElement {
        const wrap = el("div", "pt-legend");
        const item = (label: string, swatchColor: string, dash: boolean) => {
            const i = el("div", "pt-legend-item");
            const sw = el("span", dash ? "pt-legend-dash" : "pt-legend-swatch");
            if (!dash) { sw.style.background = swatchColor; }
            const tx = el("span"); tx.textContent = label;
            i.appendChild(sw); i.appendChild(tx);
            return i;
        };
        wrap.appendChild(item(t(lang, "produced"), "#56d364", false));
        wrap.appendChild(item(t(lang, "dailyTarget"), "", true));
        return wrap;
    }

    private buildBar(machine: MachineMetrics, maxVal: number, lang: Language, scale: number): HTMLElement {
        const plotH = Math.round(PLOT_H * scale);
        const barW = Math.round(BAR_W * scale);
        const cardW = Math.round(CARD_W * scale);

        const column = el("div", "pt-bar-col");
        const color = fillColor(machine.fillPct);
        column.style.width = cardW + "px";
        column.style.setProperty("--c", color);

        const value = el("div", "pt-bar-value");
        value.textContent = fmt(machine.produced);
        value.style.fontSize = Math.round(11 * scale) + "px";
        column.appendChild(value);

        const plot = el("div", "pt-bar-plot");
        plot.style.width = cardW + "px";
        plot.style.height = plotH + "px";

        const barH = Math.max(2, Math.round((machine.produced / maxVal) * plotH));
        const bar = el("div", "pt-bar");
        bar.style.width = barW + "px";
        bar.style.height = barH + "px";
        plot.appendChild(bar);

        if (machine.dailyTarget > 0) {
            const tH = Math.max(2, Math.round((machine.dailyTarget / maxVal) * plotH));
            const box = el("div", "pt-target-box");
            box.style.width = (barW + 6) + "px";
            box.style.height = tH + "px";
            plot.appendChild(box);
            const line = el("div", "pt-target-line");
            line.style.bottom = (tH - 1) + "px";
            plot.appendChild(line);
        }
        column.appendChild(plot);

        if (machine.dailyTarget > 0) {
            const tgt = el("div", "pt-bar-target");
            tgt.textContent = "⌖ " + fmt(machine.dailyTarget);
            tgt.style.fontSize = Math.round(9 * scale) + "px";
            column.appendChild(tgt);
            const pct = el("div", "pt-bar-pct");
            pct.textContent = Math.round(machine.fillPct) + "%";
            pct.style.fontSize = Math.round(10 * scale) + "px";
            column.appendChild(pct);
        }

        const name = el("div", "pt-bar-name");
        name.textContent = machine.name;
        name.style.fontSize = Math.round(10 * scale) + "px";
        name.style.maxWidth = cardW + "px";
        column.appendChild(name);

        this.wireInteractions(column, machine, lang);
        return column;
    }

    // ----- table view ---------------------------------------------------------

    private buildTable(data: ProcessedData, lang: Language): HTMLElement {
        const wrap = el("div", "pt-table-wrap");
        const table = el("table", "pt-table");
        const thead = el("thead");
        const hr = el("tr");
        const addTh = (label: string, left: boolean) => {
            const th = el("th", left ? "pt-l" : undefined);
            th.textContent = label; hr.appendChild(th);
        };
        addTh(t(lang, "machine"), true);
        addTh(t(lang, "colProduct"), true);
        addTh(t(lang, "colCycle"), false);
        addTh(t(lang, "colProduced"), false);
        addTh(t(lang, "colFtt"), false);
        addTh(t(lang, "colScrap"), false);
        addTh(t(lang, "colDowntimeCodes"), true);
        addTh(t(lang, "colDowntimeMin"), false);
        addTh(t(lang, "colOpenTime"), false);
        addTh(t(lang, "colTarget"), false);
        addTh(t(lang, "colFill"), false);
        thead.appendChild(hr);
        table.appendChild(thead);

        const tbody = el("tbody");
        for (const m of data.machines) {
            const tr = el("tr");
            this.addCell(tr, m.name, true);
            this.addCell(tr, m.products.join(", "), true);
            this.addCell(tr, m.cycleTime ? m.cycleTime.toFixed(2) : "—", false);
            this.addCell(tr, fmt(m.produced), false);
            const ftt = this.addCell(tr, fmt(m.fttLoss), false);
            ftt.style.color = "#e3b341";
            const scrap = this.addCell(tr, fmt(m.scrap), false);
            scrap.style.color = "#ff7b72";
            const codes = el("td", "pt-l");
            this.fillDowntimeBadges(codes, m.downtimeEntries);
            tr.appendChild(codes);
            this.addCell(tr, fmt(m.downtimePlanned + m.downtimeUnplanned), false);
            this.addCell(tr, fmt(m.openTime), false);
            this.addCell(tr, fmt(Math.round(m.dailyTarget)), false);
            const fillC = this.addCell(tr, Math.round(m.fillPct) + "%", false);
            fillC.style.color = fillColor(m.fillPct);
            tr.addEventListener("click", () => this.showDetail(m, lang));
            tbody.appendChild(tr);
        }
        // total row
        const tot = data.totals;
        const tr = el("tr", "pt-total");
        this.addCell(tr, t(lang, "total"), true);
        this.addCell(tr, "", true);
        this.addCell(tr, tot.cycleTime ? tot.cycleTime.toFixed(2) : "—", false);
        this.addCell(tr, fmt(tot.produced), false);
        this.addCell(tr, fmt(tot.fttLoss), false);
        this.addCell(tr, fmt(tot.scrap), false);
        this.addCell(tr, "", true);
        this.addCell(tr, fmt(tot.downtimePlanned + tot.downtimeUnplanned), false);
        this.addCell(tr, fmt(tot.openTime), false);
        this.addCell(tr, fmt(Math.round(tot.dailyTarget)), false);
        const tf = this.addCell(tr, Math.round(tot.fillPct) + "%", false);
        tf.style.color = fillColor(tot.fillPct);
        tbody.appendChild(tr);

        table.appendChild(tbody);
        wrap.appendChild(table);
        return wrap;
    }

    private addCell(tr: HTMLElement, text: string, left: boolean): HTMLElement {
        const td = el("td", left ? "pt-l" : undefined);
        td.textContent = text;
        tr.appendChild(td);
        return td;
    }

    private fillDowntimeBadges(host: HTMLElement, entries: DowntimeEntry[]): void {
        const agg = aggregateDowntime(entries);
        if (!agg.length) { host.textContent = "—"; return; }
        for (const e of agg) {
            const color = e.planned ? "#f0883e" : "#79c0ff";
            const badge = el("span", "pt-dt-badge");
            badge.textContent = e.reason + " (" + e.minutes + "m)" + (e.planned ? " P" : "");
            badge.style.color = color;
            badge.style.background = color + "22";
            badge.style.border = "1px solid " + color + "55";
            host.appendChild(badge);
        }
    }

    // ----- tile view ----------------------------------------------------------

    private buildTiles(data: ProcessedData, lang: Language): HTMLElement {
        const grid = el("div", "pt-tiles");
        for (const m of data.machines) {
            const tile = el("div", "pt-tile");
            const met = m.dailyTarget > 0 && m.produced >= m.dailyTarget;
            tile.style.borderLeftColor = met ? "#56d364" : "#ff7b72";

            const head = el("div", "pt-tile-head");
            const name = el("div", "pt-tile-name"); name.textContent = m.name;
            const prod = el("div", "pt-tile-prod"); prod.textContent = m.products.join(", ");
            head.appendChild(name); head.appendChild(prod);
            if (m.cycleTime > 0) {
                const ct = el("div", "pt-tile-ct");
                ct.textContent = t(lang, "cycleTime") + ": " + m.cycleTime.toFixed(2) + " " + t(lang, "minutes");
                head.appendChild(ct);
            }
            tile.appendChild(head);

            const kgrid = el("div", "pt-tile-grid");
            this.tileKpi(kgrid, t(lang, "produced"), fmt(m.produced), fillColor(m.fillPct));
            this.tileKpi(kgrid, t(lang, "dailyTarget"), fmt(Math.round(m.dailyTarget)), "");
            this.tileKpi(kgrid, t(lang, "colFill"), Math.round(m.fillPct) + "%", fillColor(m.fillPct));
            this.tileKpi(kgrid, t(lang, "fttLoss"), fmt(m.fttLoss), "#e3b341");
            this.tileKpi(kgrid, t(lang, "scrap"), fmt(m.scrap), "#ff7b72");
            this.tileKpi(kgrid, t(lang, "downtime"), fmt(m.downtimePlanned + m.downtimeUnplanned), "#79c0ff");
            tile.appendChild(kgrid);

            const agg = aggregateDowntime(m.downtimeEntries);
            if (agg.length) {
                const foot = el("div", "pt-tile-foot");
                this.fillDowntimeBadges(foot, m.downtimeEntries);
                tile.appendChild(foot);
            }
            tile.addEventListener("click", () => this.showDetail(m, lang));
            grid.appendChild(tile);
        }
        return grid;
    }

    private tileKpi(host: HTMLElement, label: string, value: string, color: string): void {
        const kpi = el("div", "pt-tile-kpi");
        const l = el("div", "pt-tile-kpi-label"); l.textContent = label;
        const v = el("div", "pt-tile-kpi-value"); v.textContent = value;
        if (color) { v.style.color = color; }
        kpi.appendChild(l); kpi.appendChild(v);
        host.appendChild(kpi);
    }

    // ----- OEE view -----------------------------------------------------------

    private buildOee(data: ProcessedData, lang: Language): HTMLElement {
        const wrap = el("div", "pt-oee");

        // machine pills
        const pills = el("div", "pt-machine-pills");
        const allPill = el("button", this.oeeMachine === "" ? "pt-pill pt-pill-active" : "pt-pill");
        allPill.setAttribute("type", "button");
        allPill.textContent = t(lang, "allMachines");
        allPill.addEventListener("click", () => { this.oeeMachine = ""; this.render(); });
        pills.appendChild(allPill);
        for (const m of data.machines) {
            const pill = el("button", this.oeeMachine === m.name ? "pt-pill pt-pill-active" : "pt-pill");
            pill.setAttribute("type", "button");
            pill.textContent = m.name;
            pill.addEventListener("click", () => { this.oeeMachine = m.name; this.render(); });
            pills.appendChild(pill);
        }
        wrap.appendChild(pills);

        const target = this.oeeMachine
            ? (data.machines.find(m => m.name === this.oeeMachine) || data.totals)
            : data.totals;

        const gauges = el("div", "pt-gauges");
        gauges.appendChild(this.gauge(t(lang, "oee"), target.oee, true));
        gauges.appendChild(this.gauge(t(lang, "availability"), target.availability, false));
        gauges.appendChild(this.gauge(t(lang, "quality"), target.quality, false));
        gauges.appendChild(this.gauge(t(lang, "performance"), target.performance, false));
        wrap.appendChild(gauges);

        if (this.oeeMachine === "" && data.machines.length > 1) {
            const title = el("div", "pt-breakdown-title");
            title.textContent = t(lang, "machineBreakdown");
            wrap.appendChild(title);
            for (const m of data.machines) {
                wrap.appendChild(this.breakdownRow(m));
            }
        }
        return wrap;
    }

    private gauge(label: string, ratio: number, isOee: boolean): HTMLElement {
        const pct = Math.round(ratio * 100);
        const color = gaugeColor(pct);
        const box = el("div", isOee ? "pt-gauge pt-gauge-oee" : "pt-gauge");
        const lbl = el("div", "pt-gauge-label"); lbl.textContent = label;
        box.appendChild(lbl);
        box.appendChild(svgGauge(pct, color));
        return box;
    }

    private breakdownRow(m: MachineMetrics): HTMLElement {
        const pct = Math.round(m.oee * 100);
        const color = gaugeColor(pct);
        const row = el("div", "pt-bd-row");
        const name = el("div", "pt-bd-name"); name.textContent = m.name;
        const track = el("div", "pt-bd-track");
        const fill = el("div", "pt-bd-fill");
        fill.style.width = Math.max(0, Math.min(100, pct)) + "%";
        fill.style.background = color;
        track.appendChild(fill);
        const pctEl = el("div", "pt-bd-pct"); pctEl.textContent = pct + "%"; pctEl.style.color = color;
        const formula = el("div", "pt-bd-formula");
        formula.textContent = "A:" + Math.round(m.availability * 100) + "% Q:"
            + Math.round(m.quality * 100) + "% P:" + Math.round(m.performance * 100) + "%";
        row.appendChild(name); row.appendChild(track); row.appendChild(pctEl); row.appendChild(formula);
        return row;
    }

    // ----- detail modal -------------------------------------------------------

    private showDetail(machine: MachineMetrics, lang: Language): void {
        const overlay = el("div", "pt-overlay");
        overlay.addEventListener("click", (ev) => { if (ev.target === overlay) { overlay.remove(); } });

        const panel = el("div", "pt-panel");

        const close = el("button", "pt-close");
        close.setAttribute("type", "button"); close.textContent = "✕";
        close.addEventListener("click", () => overlay.remove());
        panel.appendChild(close);

        const head = el("div", "pt-panel-head");
        const title = el("span", "pt-panel-title"); title.textContent = machine.name;
        const sub = el("span", "pt-panel-sub"); sub.textContent = machine.products.join(", ");
        head.appendChild(title); head.appendChild(sub);
        if (machine.cycleTime > 0) {
            const ct = el("span", "pt-tile-ct");
            ct.textContent = t(lang, "cycleTime") + ": " + machine.cycleTime.toFixed(2) + " " + t(lang, "minutes");
            head.appendChild(ct);
        }
        panel.appendChild(head);

        const kpis = el("div", "pt-kpis");
        this.kpi(kpis, t(lang, "produced"), fmt(machine.produced) + " / " + fmt(Math.round(machine.dailyTarget)),
            fillColor(machine.fillPct));
        this.kpi(kpis, t(lang, "fttLoss"), fmt(machine.fttLoss), "#e3b341");
        this.kpi(kpis, t(lang, "scrap"), fmt(machine.scrap), "#ff7b72");
        panel.appendChild(kpis);

        if (this.settings.display.showOee.value) {
            const oee = el("div", "pt-kpis");
            oee.style.gridTemplateColumns = "1fr 1fr 1fr 1fr";
            this.kpiPct(oee, t(lang, "oee"), machine.oee);
            this.kpiPct(oee, t(lang, "availability"), machine.availability);
            this.kpiPct(oee, t(lang, "quality"), machine.quality);
            this.kpiPct(oee, t(lang, "performance"), machine.performance);
            panel.appendChild(oee);
        }

        const cols = el("div", "pt-dt-cols");
        cols.appendChild(this.downtimeCol(t(lang, "planned"),
            machine.downtimeEntries.filter(d => d.planned), "#f0883e", lang));
        cols.appendChild(this.downtimeCol(t(lang, "unplanned"),
            machine.downtimeEntries.filter(d => !d.planned), "#79c0ff", lang));
        panel.appendChild(cols);

        overlay.appendChild(panel);
        this.root.appendChild(overlay);
    }

    private kpi(host: HTMLElement, label: string, value: string, color: string): void {
        const box = el("div", "pt-kpi");
        const l = el("div", "pt-kpi-label"); l.textContent = label;
        const v = el("div", "pt-kpi-value"); v.textContent = value; if (color) { v.style.color = color; }
        box.appendChild(l); box.appendChild(v); host.appendChild(box);
    }

    private kpiPct(host: HTMLElement, label: string, ratio: number): void {
        this.kpi(host, label, Math.round(ratio * 100) + "%", gaugeColor(Math.round(ratio * 100)));
    }

    private downtimeCol(title: string, entries: DowntimeEntry[], color: string, lang: Language): HTMLElement {
        const wrap = el("div");
        const head = el("div", "pt-dt-col-title");
        head.textContent = title + " (" + entries.length + ")";
        wrap.appendChild(head);
        const agg = aggregateDowntime(entries);
        if (!agg.length) {
            const none = el("div", "pt-dt-empty"); none.textContent = "—"; wrap.appendChild(none); return wrap;
        }
        for (const e of agg) {
            const row = el("div", "pt-dt-row");
            row.style.background = color + "15";
            row.style.borderLeft = "3px solid " + color;
            const r = el("span"); r.textContent = e.reason;
            const m = el("span", "pt-dt-min"); m.textContent = e.minutes + " " + t(lang, "minutes");
            m.style.color = color;
            row.appendChild(r); row.appendChild(m);
            wrap.appendChild(row);
        }
        return wrap;
    }

    // ----- report modal -------------------------------------------------------

    private showReport(data: ProcessedData, lang: Language): void {
        const overlay = el("div", "pt-overlay");
        overlay.addEventListener("click", (ev) => { if (ev.target === overlay) { overlay.remove(); } });
        const panel = el("div", "pt-panel pt-report");

        const machines = data.machines;
        const tot = data.totals;
        const oeePct = tot.dailyTarget > 0 ? Math.round((tot.produced / tot.dailyTarget) * 100) : 0;
        const scrapRate = (tot.produced + tot.scrap) > 0
            ? +((tot.scrap / (tot.produced + tot.scrap)) * 100).toFixed(1) : 0;
        const totDt = tot.downtimePlanned + tot.downtimeUnplanned;
        const aboveT = machines.filter(m => m.dailyTarget > 0 && m.produced >= m.dailyTarget).length;
        const belowT = machines.length - aboveT;
        const status = gaugeColor(oeePct);
        const statusText = oeePct >= 85 ? t(lang, "onTarget") : oeePct >= 60 ? t(lang, "attention") : t(lang, "critical");

        // gradient header
        const header = el("div", "pt-report-header");
        const close = el("button", "pt-report-close");
        close.setAttribute("type", "button"); close.textContent = "✕";
        close.addEventListener("click", () => overlay.remove());
        header.appendChild(close);
        const sub = el("div", "pt-report-sub"); sub.textContent = t(lang, "dailyReport");
        const ttl = el("div", "pt-report-title"); ttl.textContent = t(lang, "productionReport");
        const meta = el("div", "pt-report-meta");
        const dateStr = new Date().toLocaleDateString(lang === "hu" ? "hu-HU" : "en-US",
            { year: "numeric", month: "long", day: "numeric" });
        meta.textContent = dateStr + " • " + machines.length + " " + t(lang, "line");
        header.appendChild(sub); header.appendChild(ttl); header.appendChild(meta);
        panel.appendChild(header);

        const body = el("div", "pt-report-body");

        // status bar
        const statusBar = el("div", "pt-report-status");
        statusBar.style.borderLeftColor = status;
        const dot = el("div", "pt-report-dot"); dot.style.background = status; dot.style.boxShadow = "0 0 10px " + status + "66";
        const stMid = el("div", "pt-report-status-mid");
        const stLabel = el("div", "pt-report-status-label");
        const stLabelText = document.createTextNode(t(lang, "overallStatus") + ": ");
        const stState = el("span"); stState.textContent = statusText; stState.style.color = status;
        stLabel.appendChild(stLabelText); stLabel.appendChild(stState);
        const stSub = el("div", "pt-report-status-sub");
        stSub.textContent = aboveT + " " + t(lang, "onTargetLines") + " • " + belowT + " " + t(lang, "belowLines");
        stMid.appendChild(stLabel); stMid.appendChild(stSub);
        const stPct = el("div", "pt-report-status-pct"); stPct.textContent = oeePct + "%"; stPct.style.color = status;
        statusBar.appendChild(dot); statusBar.appendChild(stMid); statusBar.appendChild(stPct);
        body.appendChild(statusBar);

        // KPI cards
        const kpiGrid = el("div", "pt-report-kpis");
        this.reportKpi(kpiGrid, t(lang, "totalProduced"), fmt(tot.produced),
            t(lang, "plan") + ": " + fmt(Math.round(tot.dailyTarget)), oeePct, status);
        this.reportKpi(kpiGrid, t(lang, "planFulfillment"), oeePct + "%",
            oeePct >= 85 ? "✓ " + t(lang, "onTarget") : t(lang, "attention"), oeePct, status);
        this.reportKpi(kpiGrid, t(lang, "scrap"), fmt(tot.scrap),
            scrapRate + "% " + t(lang, "scrapRate"), Math.min(100, scrapRate * 10),
            scrapRate <= 2 ? "#56d364" : scrapRate <= 5 ? "#e3b341" : "#ff7b72");
        const dtPerMachine = machines.length ? Math.round(totDt / machines.length) : 0;
        this.reportKpi(kpiGrid, t(lang, "downtime"), fmt(totDt) + " " + t(lang, "minutes"),
            dtPerMachine + " " + t(lang, "perMachine"), Math.min(100, dtPerMachine / 4),
            dtPerMachine < 30 ? "#56d364" : dtPerMachine < 60 ? "#e3b341" : "#ff7b72");
        body.appendChild(kpiGrid);

        // line performance table
        if (machines.length) {
            const card = el("div", "pt-report-table-card");
            const cardHead = el("div", "pt-report-table-head");
            const ch = el("div", "pt-report-table-title"); ch.textContent = t(lang, "linePerformance");
            cardHead.appendChild(ch);
            card.appendChild(cardHead);

            const table = el("table", "pt-table");
            const thead = el("thead"); const hr = el("tr");
            const addTh = (label: string, left: boolean) => { const th = el("th", left ? "pt-l" : undefined); th.textContent = label; hr.appendChild(th); };
            addTh(t(lang, "line"), true); addTh(t(lang, "colProduct"), true);
            addTh(t(lang, "produced"), false); addTh(t(lang, "colTarget"), false);
            addTh(t(lang, "fulfillment"), false); addTh(t(lang, "scrap"), false); addTh("DT", false);
            thead.appendChild(hr); table.appendChild(thead);
            const tbody = el("tbody");
            for (const m of machines) {
                const pct = m.dailyTarget > 0 ? Math.round((m.produced / m.dailyTarget) * 100) : 0;
                const tr = el("tr");
                this.addCell(tr, m.name, true);
                this.addCell(tr, m.products.join(", "), true);
                this.addCell(tr, fmt(m.produced), false);
                this.addCell(tr, fmt(Math.round(m.dailyTarget)), false);
                const fc = this.addCell(tr, pct + "%", false); fc.style.color = gaugeColor(pct);
                const sc = this.addCell(tr, fmt(m.scrap), false); if (m.scrap > 0) { sc.style.color = "#ff7b72"; }
                const dc = this.addCell(tr, fmt(m.downtimePlanned + m.downtimeUnplanned) + "m", false);
                if (m.downtimePlanned + m.downtimeUnplanned > 45) { dc.style.color = "#e3b341"; }
                tbody.appendChild(tr);
            }
            table.appendChild(tbody);
            card.appendChild(table);
            body.appendChild(card);
        }

        panel.appendChild(body);
        overlay.appendChild(panel);
        this.root.appendChild(overlay);
    }

    private reportKpi(host: HTMLElement, label: string, value: string, sub: string,
                      pct: number, color: string): void {
        const card = el("div", "pt-report-kpi");
        const bar = el("div", "pt-report-kpi-bar");
        const fill = el("div", "pt-report-kpi-fill");
        fill.style.width = Math.min(100, Math.max(0, pct)) + "%"; fill.style.background = color;
        bar.appendChild(fill); card.appendChild(bar);
        const l = el("div", "pt-report-kpi-label"); l.textContent = label;
        const v = el("div", "pt-report-kpi-value"); v.textContent = value;
        const s = el("div", "pt-report-kpi-sub"); s.textContent = sub;
        card.appendChild(l); card.appendChild(v); card.appendChild(s);
        host.appendChild(card);
    }

    // ----- interactions -------------------------------------------------------

    private wireInteractions(element: HTMLElement, machine: MachineMetrics, lang: Language): void {
        element.addEventListener("click", (ev: MouseEvent) => {
            ev.stopPropagation();
            this.selectionManager.select(machine.selectionId, ev.ctrlKey || ev.metaKey);
            this.showDetail(machine, lang);
        });
        element.addEventListener("contextmenu", (ev: MouseEvent) => {
            ev.preventDefault();
            this.selectionManager.showContextMenu(machine.selectionId, { x: ev.clientX, y: ev.clientY });
        });
        element.addEventListener("mousemove", (ev: MouseEvent) => {
            this.tooltipService.show({
                coordinates: [ev.clientX, ev.clientY],
                isTouchEvent: false,
                dataItems: this.tooltipItems(machine, lang),
                identities: [machine.selectionId]
            });
        });
        element.addEventListener("mouseout", () => {
            this.tooltipService.hide({ immediately: false, isTouchEvent: false });
        });
    }

    private tooltipItems(machine: MachineMetrics, lang: Language): powerbi.extensibility.VisualTooltipDataItem[] {
        return [
            { displayName: t(lang, "machine"), value: machine.name },
            { displayName: t(lang, "produced"), value: fmt(machine.produced) },
            { displayName: t(lang, "dailyTarget"), value: fmt(Math.round(machine.dailyTarget)) },
            { displayName: t(lang, "oee"), value: Math.round(machine.oee * 100) + "%" }
        ];
    }

    private persist(property: string, value: powerbi.PrimitiveValue): void {
        this.host.persistProperties({
            merge: [{ objectName: "display", selector: undefined, properties: { [property]: value } }]
        });
    }

}

// ----- small DOM / SVG helpers ------------------------------------------------

function el(tag: string, className?: string): HTMLElement {
    const node = document.createElement(tag);
    if (className) { node.className = className; }
    return node as HTMLElement;
}

function clear(node: HTMLElement): void {
    while (node.firstChild) { node.removeChild(node.firstChild); }
}

function fmt(value: number): string {
    return Math.round(value).toLocaleString();
}

interface AggDowntime { reason: string; minutes: number; planned: boolean; }

/** Combine repeated downtime reasons into a single badge each. */
function aggregateDowntime(entries: DowntimeEntry[]): AggDowntime[] {
    const map = new Map<string, AggDowntime>();
    for (const e of entries) {
        const key = e.reason + "|" + e.planned;
        const cur = map.get(key);
        if (cur) { cur.minutes += e.minutes; }
        else { map.set(key, { reason: e.reason, minutes: e.minutes, planned: e.planned }); }
    }
    return [...map.values()].sort((a, b) => b.minutes - a.minutes);
}

const SVG_NS = "http://www.w3.org/2000/svg";

/** Build a circular OEE gauge as an SVG (no innerHTML). */
function svgGauge(pct: number, color: string): SVGElement {
    const size = 120, r = 50, cx = size / 2, cy = size / 2;
    const circ = 2 * Math.PI * r;
    const svg = document.createElementNS(SVG_NS, "svg");
    svg.setAttribute("width", String(size));
    svg.setAttribute("height", String(size));
    svg.setAttribute("viewBox", "0 0 " + size + " " + size);

    const track = document.createElementNS(SVG_NS, "circle");
    track.setAttribute("cx", String(cx)); track.setAttribute("cy", String(cy)); track.setAttribute("r", String(r));
    track.setAttribute("fill", "none"); track.setAttribute("stroke", "#21262d"); track.setAttribute("stroke-width", "10");
    svg.appendChild(track);

    const arc = document.createElementNS(SVG_NS, "circle");
    arc.setAttribute("cx", String(cx)); arc.setAttribute("cy", String(cy)); arc.setAttribute("r", String(r));
    arc.setAttribute("fill", "none"); arc.setAttribute("stroke", color); arc.setAttribute("stroke-width", "10");
    arc.setAttribute("stroke-linecap", "round");
    arc.setAttribute("stroke-dasharray", String(circ));
    arc.setAttribute("stroke-dashoffset", String(circ * (1 - Math.max(0, Math.min(100, pct)) / 100)));
    arc.setAttribute("transform", "rotate(-90 " + cx + " " + cy + ")");
    svg.appendChild(arc);

    const text = document.createElementNS(SVG_NS, "text");
    text.setAttribute("x", String(cx)); text.setAttribute("y", String(cy));
    text.setAttribute("text-anchor", "middle"); text.setAttribute("dominant-baseline", "central");
    text.setAttribute("fill", color);
    text.setAttribute("font-size", "22"); text.setAttribute("font-weight", "700");
    text.textContent = pct + "%";
    svg.appendChild(text);
    return svg;
}
