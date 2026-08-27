/**
 * Transforms the categorical dataView into per-machine production metrics
 * (OEE, downtime split, daily target) grouped by area. No DAX — all derived here.
 */
import powerbi from "powerbi-visuals-api";
import DataView = powerbi.DataView;
import DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
import ISelectionId = powerbi.visuals.ISelectionId;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import { TrackerSettingsModel } from "./settings";

export interface DowntimeEntry { reason: string; minutes: number; planned: boolean; }

export interface MachineMetrics {
    name: string;
    area: string;
    areaSort: string;
    products: string[];
    produced: number;
    fttLoss: number;
    scrap: number;
    downtimePlanned: number;
    downtimeUnplanned: number;
    downtimeEntries: DowntimeEntry[];
    cycleTime: number;
    shiftSlots: number;
    openTime: number;
    runTime: number;
    dailyTarget: number;
    availability: number;
    performance: number;
    quality: number;
    oee: number;
    fillPct: number;
    selectionId: ISelectionId;
}

export interface AreaGroup { name: string; sort: string; machines: MachineMetrics[]; }

export interface FilterValues {
    shifts: string[]; lines: string[]; locations: string[]; products: string[]; categories: string[];
}

export interface FilterState {
    shift: string; line: string; location: string; product: string; category: string;
    dateFrom: string; dateTo: string;
}

/** Shift schedule parameters that drive the shift count and open time. */
export interface ScheduleParams {
    shiftLengthMinutes: number;
    plannedStopMinutes: number;
    shiftsPerDay: number;
    daysOn: number;
    daysOff: number;
    weekPattern: number[] | null;   // days of week worked (0=Sun..6=Sat); null = use daysOn/daysOff ratio
}

export interface ProcessedData {
    machines: MachineMetrics[];
    areas: AreaGroup[];
    totals: MachineMetrics;
    filters: FilterValues;
    hasData: boolean;
    shiftCount: number;       // shifts in the period (from the schedule)
    factoryOpenTime: number;  // shiftCount × (shiftLength − plannedStop)
    // v5.4.0 free tier: machines beyond the cap are dropped before any metric is
    // computed, so the chart, table and totals all agree on the same machine set.
    machinesHidden: number;   // how many machines the cap removed (0 when licensed)
    machinesTotal: number;    // machines present in the data, before the cap
}

export const EMPTY_FILTER: FilterState =
    { shift: "", line: "", location: "", product: "", category: "", dateFrom: "", dateTo: "" };

function isoFromValue(raw: powerbi.PrimitiveValue | null | undefined): string {
    if (raw == null) { return ""; }
    const date = raw instanceof Date ? raw : new Date(String(raw));
    if (isNaN(date.getTime())) { return ""; }
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return date.getFullYear() + "-" + m + "-" + d;
}

function clamp01(x: number): number { return x < 0 ? 0 : x > 1 ? 1 : x; }

function categoryIndex(categories: DataViewCategoryColumn[], role: string): number {
    return categories.findIndex(c => c.source.roles && c.source.roles[role]);
}
function valueIndex(values: powerbi.DataViewValueColumns, role: string): number {
    for (let i = 0; i < values.length; i++) {
        if (values[i].source.roles && values[i].source.roles[role]) { return i; }
    }
    return -1;
}

/** Mutable accumulator while scanning rows. */
interface MachineAccumulator {
    name: string; area: string; areaSort: string;
    products: Set<string>;
    produced: number; fttLoss: number; scrap: number;
    downtimePlanned: number; downtimeUnplanned: number;
    downtimeEntries: DowntimeEntry[];
    sumProducedCt: number; sumCt: number; ctCount: number;
    shiftSlots: Set<string>;
    firstRow: number;
}

export function transform(
    dataView: DataView | undefined,
    settings: TrackerSettingsModel,
    host: IVisualHost,
    filter: FilterState,
    schedule?: ScheduleParams,
    maxMachines?: number
): ProcessedData {
    const empty: ProcessedData = {
        machines: [], areas: [], totals: blankTotals(),
        filters: { shifts: [], lines: [], locations: [], products: [], categories: [] }, hasData: false,
        shiftCount: 0, factoryOpenTime: 0, machinesHidden: 0, machinesTotal: 0
    };
    const categorical = dataView && dataView.categorical;
    if (!categorical || !categorical.categories || !categorical.categories.length) { return empty; }

    const cats = categorical.categories;
    const vals = categorical.values;
    const iMachine = categoryIndex(cats, "machine");
    if (iMachine < 0) { return empty; }
    const machineColumn = cats[iMachine];

    const iProduct = categoryIndex(cats, "product");
    const iReason = categoryIndex(cats, "downtimeCode");      // "Downtime Reason" (display label)
    const iCode = categoryIndex(cats, "downtimeCodeKey");     // "Downtime Code" (drives planned)
    const iShift = categoryIndex(cats, "shiftName");
    const iDate = categoryIndex(cats, "dateField");
    const iLine = categoryIndex(cats, "productionLine");
    const iCategory = categoryIndex(cats, "productCategory");
    const iArea = categoryIndex(cats, "areaName");
    const iAreaSort = categoryIndex(cats, "areaSort");

    const iProduced = vals ? valueIndex(vals, "producedQty") : -1;
    const iFtt = vals ? valueIndex(vals, "fttLossQty") : -1;
    const iScrap = vals ? valueIndex(vals, "scrapQty") : -1;
    const iDowntime = vals ? valueIndex(vals, "downtimeMin") : -1;
    const iCycle = vals ? valueIndex(vals, "cycleTime") : -1;

    const rowCount = machineColumn.values.length;
    const cell = (idx: number, row: number): string =>
        idx >= 0 ? String(cats[idx].values[row] ?? "").trim() : "";
    const num = (idx: number, row: number): number =>
        idx >= 0 && vals ? (Number(vals[idx].values[row]) || 0) : 0;

    const shifts = new Set<string>(), lines = new Set<string>(), locations = new Set<string>(),
        products = new Set<string>(), categories = new Set<string>();
    const dateSet = new Set<string>();   // distinct production dates after filtering
    const machineMap = new Map<string, MachineAccumulator>();
    const order: string[] = [];

    for (let row = 0; row < rowCount; row++) {
        const machine = String(machineColumn.values[row] ?? "").trim();
        if (!machine) { continue; }

        const shift = cell(iShift, row), line = cell(iLine, row), location = cell(iArea, row);
        const product = cell(iProduct, row), category = cell(iCategory, row);
        const iso = iDate >= 0 ? isoFromValue(cats[iDate].values[row]) : "";
        if (shift) { shifts.add(shift); }
        if (line) { lines.add(line); }
        if (location) { locations.add(location); }
        if (product) { products.add(product); }
        if (category) { categories.add(category); }

        // active filters
        if (filter.shift && shift && shift !== filter.shift) { continue; }
        if (filter.line && line && line !== filter.line) { continue; }
        if (filter.location && location && location !== filter.location) { continue; }
        if (filter.product && product && product !== filter.product) { continue; }
        if (filter.category && category && category !== filter.category) { continue; }
        if (filter.dateFrom && iso && iso < filter.dateFrom) { continue; }
        if (filter.dateTo && iso && iso > filter.dateTo) { continue; }

        if (iso) { dateSet.add(iso); }

        let acc = machineMap.get(machine);
        if (!acc) {
            acc = {
                name: machine, area: cell(iArea, row), areaSort: cell(iAreaSort, row),
                products: new Set<string>(), produced: 0, fttLoss: 0, scrap: 0,
                downtimePlanned: 0, downtimeUnplanned: 0, downtimeEntries: [],
                sumProducedCt: 0, sumCt: 0, ctCount: 0, shiftSlots: new Set<string>(), firstRow: row
            };
            machineMap.set(machine, acc);
            order.push(machine);
        }

        const produced = num(iProduced, row);
        const cycle = num(iCycle, row);
        acc.produced += produced;
        acc.fttLoss += num(iFtt, row);
        acc.scrap += num(iScrap, row);
        if (product) { acc.products.add(product); }
        if (cycle > 0) { acc.sumProducedCt += produced * cycle; acc.sumCt += cycle; acc.ctCount++; }
        if (iso || shift) { acc.shiftSlots.add(iso + "|" + shift); }

        const downtime = num(iDowntime, row);
        if (downtime > 0) {
            const reason = cell(iReason, row) || cell(iCode, row) || "—";
            const code = cell(iCode, row) || cell(iReason, row);
            const planned = code.startsWith("S");
            if (planned) { acc.downtimePlanned += downtime; } else { acc.downtimeUnplanned += downtime; }
            acc.downtimeEntries.push({ reason, minutes: downtime, planned });
        }
    }

    // Shift count + factory open time come from the schedule (template), not the
    // raw date×shift combinations in the data — so switching shift template (e.g.
    // 12h continuous = 2 shifts/day vs 8h = 3 shifts/day) changes the shift count.
    const sched: ScheduleParams = schedule || {
        shiftLengthMinutes: settings.shiftSchedule.shiftLengthMinutes.value || 480,
        plannedStopMinutes: settings.shiftSchedule.plannedStopMinutes.value || 0,
        shiftsPerDay: settings.shiftSchedule.shiftsPerDay.value || 1,
        daysOn: settings.shiftSchedule.workingDaysPerWeek.value || 5,
        daysOff: Math.max(0, 7 - (settings.shiftSchedule.workingDaysPerWeek.value || 5)),
        weekPattern: null
    };
    const productive = Math.max(1, (sched.shiftLengthMinutes || 0) - (sched.plannedStopMinutes || 0));
    const shiftCount = computeShiftCount([...dateSet], sched, !!filter.shift);
    const factoryOpenTime = shiftCount * productive;

    // v5.4.0: apply the free-tier machine cap before finalizing, so totals and
    // OEE averages reflect exactly the machines the user can actually see.
    const machinesTotal = order.length;
    const capped = typeof maxMachines === "number" && maxMachines >= 0 && machinesTotal > maxMachines
        ? order.slice(0, maxMachines)
        : order;
    const machinesHidden = machinesTotal - capped.length;

    const machines: MachineMetrics[] = capped.map(name => finalize(machineMap.get(name), factoryOpenTime, host, machineColumn));

    // group by area (sorted by areaSort then name)
    const areaMap = new Map<string, AreaGroup>();
    for (const m of machines) {
        const key = m.area || "—";
        let g = areaMap.get(key);
        if (!g) { g = { name: key, sort: m.areaSort || key, machines: [] }; areaMap.set(key, g); }
        g.machines.push(m);
    }
    const areas = [...areaMap.values()].sort((a, b) => a.sort.localeCompare(b.sort, undefined, { numeric: true }));

    return {
        machines, areas, totals: aggregateTotals(machines, factoryOpenTime, shiftCount),
        filters: {
            shifts: [...shifts].sort(), lines: [...lines].sort(), locations: [...locations].sort(),
            products: [...products].sort(), categories: [...categories].sort()
        },
        hasData: machines.length > 0,
        shiftCount, factoryOpenTime, machinesHidden, machinesTotal
    };
}

/**
 * Number of shifts in the data's date range, given the schedule.
 * Mirrors the original getScheduleInfo: working days in range (by week pattern,
 * or daysOn/daysOff ratio) × shifts per day, plus weekend extras. When a single
 * shift is filtered, shifts/day collapses to 1.
 */
function computeShiftCount(isoDates: string[], sched: ScheduleParams, singleShift: boolean): number {
    const spd = singleShift ? 1 : Math.max(1, sched.shiftsPerDay || 1);
    const unique = [...new Set(isoDates)].sort();
    if (!unique.length) { return spd; }
    const start = new Date(unique[0] + "T12:00:00");
    const end = new Date(unique[unique.length - 1] + "T12:00:00");
    const present = new Set(unique);
    const wp = sched.weekPattern;
    let workdays = 0, extra = 0;
    if (wp && wp.length) {
        for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
            const key = d.getFullYear() + "-"
                + String(d.getMonth() + 1).padStart(2, "0") + "-"
                + String(d.getDate()).padStart(2, "0");
            if (wp.indexOf(d.getDay()) >= 0) { workdays++; }
            else if (present.has(key)) { extra += spd; }
        }
    } else {
        const totalDays = Math.round((end.getTime() - start.getTime()) / 864e5) + 1;
        const on = sched.daysOn || 5, off = sched.daysOff || 2;
        workdays = Math.max(1, Math.round((totalDays * on) / (on + off)));
    }
    workdays = Math.max(1, workdays);
    return workdays * spd + extra;
}

function finalize(acc: MachineAccumulator, openTime: number,
                  host: IVisualHost, machineColumn: DataViewCategoryColumn): MachineMetrics {
    const cycleTime = acc.produced > 0 && acc.sumProducedCt > 0
        ? acc.sumProducedCt / acc.produced
        : (acc.ctCount > 0 ? acc.sumCt / acc.ctCount : 0);
    const runTime = Math.max(0, openTime - acc.downtimeUnplanned);
    const dailyTarget = cycleTime > 0 ? openTime / cycleTime : 0;
    const idealCount = cycleTime > 0 ? runTime / cycleTime : 0;
    const good = Math.max(0, acc.produced - acc.fttLoss - acc.scrap);

    const availability = openTime > 0 ? clamp01(runTime / openTime) : 0;
    const performance = idealCount > 0 ? clamp01(acc.produced / idealCount) : 0;
    const quality = acc.produced > 0 ? clamp01(good / acc.produced) : 0;
    const oee = availability * performance * quality;
    const fillPct = dailyTarget > 0 ? (acc.produced / dailyTarget) * 100 : 0;

    const selectionId = host.createSelectionIdBuilder()
        .withCategory(machineColumn, acc.firstRow).createSelectionId();

    return {
        name: acc.name, area: acc.area, areaSort: acc.areaSort, products: [...acc.products].sort(),
        produced: acc.produced, fttLoss: acc.fttLoss, scrap: acc.scrap,
        downtimePlanned: acc.downtimePlanned, downtimeUnplanned: acc.downtimeUnplanned,
        downtimeEntries: acc.downtimeEntries, cycleTime, shiftSlots: acc.shiftSlots.size,
        openTime, runTime, dailyTarget, availability, performance, quality, oee, fillPct, selectionId
    };
}

function blankTotals(): MachineMetrics {
    return {
        name: "", area: "", areaSort: "", products: [], produced: 0, fttLoss: 0, scrap: 0,
        downtimePlanned: 0, downtimeUnplanned: 0, downtimeEntries: [], cycleTime: 0, shiftSlots: 0,
        openTime: 0, runTime: 0, dailyTarget: 0, availability: 0, performance: 0, quality: 0,
        oee: 0, fillPct: 0, selectionId: undefined
    };
}

function aggregateTotals(machines: MachineMetrics[], factoryOpenTime: number, shiftCount: number): MachineMetrics {
    const total = blankTotals();
    // All machines share the factory open time; aggregate OEE is the average machine.
    let sumProducedCt = 0, sumA = 0, sumP = 0, sumQ = 0, sumOee = 0;
    const n = machines.length || 1;
    for (const m of machines) {
        total.produced += m.produced; total.fttLoss += m.fttLoss; total.scrap += m.scrap;
        total.downtimePlanned += m.downtimePlanned; total.downtimeUnplanned += m.downtimeUnplanned;
        total.dailyTarget += m.dailyTarget;
        sumProducedCt += m.produced * m.cycleTime;
        sumA += m.availability; sumP += m.performance; sumQ += m.quality; sumOee += m.oee;
    }
    total.cycleTime = total.produced > 0 ? sumProducedCt / total.produced : 0;
    total.openTime = factoryOpenTime;
    total.availability = sumA / n;
    total.performance = sumP / n;
    total.quality = sumQ / n;
    total.oee = sumOee / n;
    total.runTime = Math.max(0, factoryOpenTime * total.availability);
    total.fillPct = total.dailyTarget > 0 ? (total.produced / total.dailyTarget) * 100 : 0;
    total.shiftSlots = shiftCount;
    return total;
}

/** Bar colour by fill %: >=90 green, 85 yellow, <=70 red, smooth blend. */
export function fillColor(pct: number): string {
    const G = [52, 199, 89], Y = [245, 197, 24], R = [229, 72, 77];
    const lerp = (a: number[], b: number[], tt: number) =>
        a.map((v, i) => Math.round(v + (b[i] - v) * tt));
    let c: number[];
    if (pct >= 90) { c = G; }
    else if (pct >= 85) { c = lerp(Y, G, (pct - 85) / 5); }
    else if (pct >= 70) { c = lerp(R, Y, (pct - 70) / 15); }
    else { c = R; }
    return "#" + c.map(v => Math.max(0, Math.min(255, v)).toString(16).padStart(2, "0")).join("");
}

/** OEE gauge colour: >=85 green, >=60 yellow, else red (GitHub-dark palette). */
export function gaugeColor(pct: number): string {
    if (pct >= 85) { return "#56d364"; }
    if (pct >= 60) { return "#e3b341"; }
    return "#ff7b72";
}
