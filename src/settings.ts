/**
 * Format-pane settings for the Production Tracker, built with the Power BI
 * formatting-model utils (no DAX, configured from the Format pane).
 */
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import Model = formattingSettings.Model;
import SimpleCard = formattingSettings.SimpleCard;
import NumUpDown = formattingSettings.NumUpDown;
import ToggleSwitch = formattingSettings.ToggleSwitch;
import ItemDropdown = formattingSettings.ItemDropdown;
import { Language } from "./i18n";

const LANGUAGE_ITEMS = [
    { value: "en", displayName: "English" },
    { value: "hu", displayName: "Magyar" }
];

/** Shift schedule — used to compute Open Time, Daily Target and OEE. */
class ShiftScheduleCard extends SimpleCard {
    shiftLengthMinutes = new NumUpDown({
        name: "shiftLengthMinutes", displayName: "Shift length (min)", value: 480
    });
    shiftsPerDay = new NumUpDown({
        name: "shiftsPerDay", displayName: "Shifts per day", value: 3
    });
    plannedStopMinutes = new NumUpDown({
        name: "plannedStopMinutes", displayName: "Planned stop / shift (min)", value: 30
    });
    workingDaysPerWeek = new NumUpDown({
        name: "workingDaysPerWeek", displayName: "Working days / week", value: 5
    });

    name = "shiftSchedule";
    displayName = "Shift schedule";
    slices = [this.shiftLengthMinutes, this.shiftsPerDay, this.plannedStopMinutes, this.workingDaysPerWeek];
}

/** Display options. */
class DisplayCard extends SimpleCard {
    showOee = new ToggleSwitch({ name: "showOee", displayName: "Show OEE metrics", value: true });
    language = new ItemDropdown({
        name: "language", displayName: "Language",
        items: LANGUAGE_ITEMS, value: LANGUAGE_ITEMS[0]
    });

    name = "display";
    displayName = "Display";
    slices = [this.showOee, this.language];
}

export class TrackerSettingsModel extends Model {
    shiftSchedule = new ShiftScheduleCard();
    display = new DisplayCard();
    cards = [this.shiftSchedule, this.display];

    get languageCode(): Language {
        return this.display.language.value.value === "hu" ? "hu" : "en";
    }
}
