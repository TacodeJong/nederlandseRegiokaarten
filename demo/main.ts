// Renders the visual outside Power BI with the sample data, for screenshots: demo/preview.html?view=gemeenten|provincies|punten
import { Visual } from "../src/visual";
import * as sample from "../samples/sample-data.json";

type Row = Record<string, string | number>;

const params = new URLSearchParams(location.search);
const view = params.get("view") || "gemeenten";
const element = document.getElementById("visual") as HTMLElement;

class SelectionId {
    constructor(private key: string) { }
    getKey() { return this.key; }
    getSelector() { return { data: [this.key] }; }
    equals(other: SelectionId) { return other && other.getKey() === this.key; }
    includes(other: SelectionId) { return this.equals(other); }
    hasIdentity() { return true; }
}
const host: any = {
    locale: "nl-NL",
    tooltipService: { enabled: () => true, show: () => undefined, move: () => undefined, hide: () => undefined },
    colorPalette: { getColor: () => ({ value: "#118DFF" }), isHighContrast: false },
    createSelectionIdBuilder: () => {
        const parts: string[] = [];
        const builder: any = {
            withCategory: (column: any, index: number) => { parts.push(`${column.source.queryName}:${index}`); return builder; },
            withMeasure: (measure: string) => { parts.push(measure); return builder; },
            withSeries: () => builder,
            createSelectionId: () => new SelectionId(parts.join("|"))
        };
        return builder;
    },
    createSelectionManager: () => ({
        select: () => Promise.resolve([]),
        clear: () => Promise.resolve(),
        getSelectionIds: () => [],
        hasSelection: () => false,
        registerOnSelectCallback: () => undefined,
        showContextMenu: () => Promise.resolve()
    }),
    persistProperties: () => undefined,
    eventService: { renderingStarted: () => undefined, renderingFinished: () => undefined, renderingFailed: () => undefined },
    hostCapabilities: { allowInteractions: true },
    displayWarningIcon: () => undefined,
    launchUrl: () => undefined,
    refreshHostData: () => undefined
};

const category = (role: string, name: string, values: (string | number)[]) => ({
    source: { displayName: name, queryName: `Data.${name}`, roles: { [role]: true }, type: { text: typeof values[0] === "string", numeric: typeof values[0] === "number" } },
    values,
    identity: values.map((_, index) => ({ key: `${name}:${index}` }))
});
const measure = (role: string, name: string, values: number[]) => ({
    source: { displayName: name, queryName: `Data.${name}`, roles: { [role]: true }, isMeasure: true, type: { numeric: true } },
    values
});
const prop = (properties: Record<string, unknown>) => properties;

const rows: Row[] = view === "provincies" ? (sample as any).provincies : view === "punten" ? (sample as any).punten : (sample as any).gemeenten;
const categorical = view === "punten"
    ? {
        categories: [
            category("latitude", "Latitude", rows.map(row => row.Latitude)),
            category("longitude", "Longitude", rows.map(row => row.Longitude)),
            category("pointCategory", "Type", rows.map(row => row.Type)),
            category("pointTooltip", "Locatie", rows.map(row => row.Locatie))
        ],
        values: [] as any[]
    }
    : {
        categories: [category("category", view === "provincies" ? "Provincie" : "Gemeente", rows.map(row => row[view === "provincies" ? "Provincie" : "Gemeente"]))],
        values: [measure("measure", "Score", rows.map(row => Number(row.Score)))]
    };
const objects: any = {
    gemeenten: {
        mapSettings: prop({ selectedMapSlice: "gemeenten 2023", useMidColorSlice: true }),
        labelSettings: prop({ showLabels: false })
    },
    provincies: { mapSettings: prop({ selectedMapSlice: "provincies" }), labelSettings: prop({ showLabels: true, showValues: true, showRegionName: true }) },
    punten: { mapSettings: prop({ selectedMapSlice: "provincies" }), pointSettings: prop({ showPoints: true, useLegendColors: true, showPointLegend: true, pointSize: 8 }) }
}[view];

const viewport = { width: element.clientWidth, height: element.clientHeight };
const visual = new Visual({ element, host } as any);
visual.update({
    viewport,
    type: 2,
    dataViews: [{ metadata: { columns: [], objects }, categorical }]
} as any);
document.body.setAttribute("data-ready", "1");
