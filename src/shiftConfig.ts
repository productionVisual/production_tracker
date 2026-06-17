/**
 * Shift schedule templates and planned-stop categories for the detailed
 * shift-settings panel (recreated from the original Production Tracker).
 */
import { Language } from "./i18n";
import { ScheduleParams } from "./dataModel";

export interface ShiftTemplate {
    hu: string; en: string;
    minutes: number;      // 0 = custom (user enters minutes)
    daysOn: number;
    daysOff: number;
    shiftsPerDay: number;
    weekPattern: number[] | null;   // worked weekdays (0=Sun..6=Sat); null = use daysOn/daysOff ratio
    custom?: boolean;
}

export interface PlannedStopCategory {
    hu: string; en: string;
    defaultMinutes: number;
}

const WD = [1, 2, 3, 4, 5];           // Mon–Fri
const WD6 = [1, 2, 3, 4, 5, 6];       // Mon–Sat
const ALL = [0, 1, 2, 3, 4, 5, 6];    // every day

export const SHIFT_TEMPLATES: ShiftTemplate[] = [
    { hu: "5/2 nyolc óra (8h)", en: "5/2 eight hours (8h)", minutes: 480, daysOn: 5, daysOff: 2, shiftsPerDay: 2, weekPattern: WD },
    { hu: "5/3 nyolc óra (8h)", en: "5/3 eight hours (8h)", minutes: 480, daysOn: 5, daysOff: 2, shiftsPerDay: 3, weekPattern: WD },
    { hu: "5/3 tizenkét óra (12h)", en: "5/3 twelve hours (12h)", minutes: 720, daysOn: 5, daysOff: 3, shiftsPerDay: 2, weekPattern: null },
    { hu: "4/2 nyolc óra (8h)", en: "4/2 eight hours (8h)", minutes: 480, daysOn: 4, daysOff: 2, shiftsPerDay: 2, weekPattern: null },
    { hu: "4/2 tizenkét óra (12h)", en: "4/2 twelve hours (12h)", minutes: 720, daysOn: 4, daysOff: 2, shiftsPerDay: 2, weekPattern: null },
    { hu: "6/2 nyolc óra (8h)", en: "6/2 eight hours (8h)", minutes: 480, daysOn: 6, daysOff: 2, shiftsPerDay: 2, weekPattern: WD6 },
    { hu: "6/2 tizenkét óra (12h)", en: "6/2 twelve hours (12h)", minutes: 720, daysOn: 6, daysOff: 2, shiftsPerDay: 2, weekPattern: WD6 },
    { hu: "2/2 nyolc óra (8h)", en: "2/2 eight hours (8h)", minutes: 480, daysOn: 2, daysOff: 2, shiftsPerDay: 2, weekPattern: null },
    { hu: "2/2 tizenkét óra (12h)", en: "2/2 twelve hours (12h)", minutes: 720, daysOn: 2, daysOff: 2, shiftsPerDay: 2, weekPattern: null },
    { hu: "Folyamatos nyolc óra (8h)", en: "Continuous eight hours (8h)", minutes: 480, daysOn: 7, daysOff: 0, shiftsPerDay: 3, weekPattern: ALL },
    { hu: "Folyamatos tizenkét óra (12h)", en: "Continuous twelve hours (12h)", minutes: 720, daysOn: 7, daysOff: 0, shiftsPerDay: 2, weekPattern: ALL },
    { hu: "Egyedi", en: "Custom", minutes: 0, daysOn: 5, daysOff: 2, shiftsPerDay: 2, weekPattern: WD, custom: true }
];

export const PLANNED_STOPS: PlannedStopCategory[] = [
    { hu: "Ebédszünet", en: "Lunch break", defaultMinutes: 30 },
    { hu: "Reggeli szünet", en: "Morning break", defaultMinutes: 10 },
    { hu: "Délutáni szünet", en: "Afternoon break", defaultMinutes: 10 },
    { hu: "Műszakváltás", en: "Shift change", defaultMinutes: 15 },
    { hu: "Tervezett karbantartás", en: "Planned maintenance", defaultMinutes: 0 },
    { hu: "Egyéb tervezett állás", en: "Other planned stop", defaultMinutes: 0 }
];

export function templateLabel(tpl: ShiftTemplate, lang: Language): string {
    return lang === "hu" ? tpl.hu : tpl.en;
}

export function stopLabel(stop: PlannedStopCategory, lang: Language): string {
    return lang === "hu" ? stop.hu : stop.en;
}

/** Runtime, persisted shift configuration (serialised to the shiftConfig property). */
export interface ShiftConfig {
    templateIdx: number;
    customMinutes: number;
    stops: number[];
}

export function effectiveShiftMinutes(cfg: ShiftConfig): number {
    const tpl = SHIFT_TEMPLATES[cfg.templateIdx] || SHIFT_TEMPLATES[0];
    return tpl.custom ? cfg.customMinutes : tpl.minutes;
}

export function plannedStopTotal(cfg: ShiftConfig): number {
    return cfg.stops.reduce((sum, v) => sum + (v || 0), 0);
}

/** Build the schedule parameters the data transform needs from a config. */
export function scheduleParamsFor(cfg: ShiftConfig): ScheduleParams {
    const tpl = SHIFT_TEMPLATES[cfg.templateIdx] || SHIFT_TEMPLATES[0];
    return {
        shiftLengthMinutes: effectiveShiftMinutes(cfg),
        plannedStopMinutes: plannedStopTotal(cfg),
        shiftsPerDay: tpl.shiftsPerDay,
        daysOn: tpl.daysOn,
        daysOff: tpl.daysOff,
        weekPattern: tpl.weekPattern
    };
}
