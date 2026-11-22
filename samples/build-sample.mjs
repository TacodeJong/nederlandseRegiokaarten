// Generates the sample data (CSV + JSON) and samples/pbip/NederlandseRegiokaarten-sample.pbip from dist/*.pbiviz.
// Usage: npx pbiviz package && node samples/build-sample.mjs
import { execFileSync } from "node:child_process";
import { randomUUID, createHash } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { geoCentroid, geoContains } from "d3-geo";
import { feature } from "topojson-client";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, "..");
// Keep names short: Windows MAX_PATH (260) breaks Power BI on nested CustomVisuals paths.
const name = "NLMap-sample";
const out = path.join(here, "pbip");
const report = path.join(out, `${name}.Report`);
const model = path.join(out, `${name}.SemanticModel`);

const write = (file, content) => {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, typeof content === "string" ? content : JSON.stringify(content, null, 2));
};
const stableId = seed => {
    const hex = createHash("sha1").update(seed).digest("hex");
    return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20, 32)}`;
};
const features = file => {
    const topology = JSON.parse(fs.readFileSync(path.join(root, "topojson", file), "utf8"));
    return feature(topology, Object.values(topology.objects)[0]).features;
};
const noise = text => parseInt(createHash("md5").update(text).digest("hex").slice(0, 6), 16) / 0xffffff;

// ---------- sample data (fictitious values) ----------
const provinces = features("provincies.json");
const provinceOf = point => provinces.find(province => geoContains(province, point))?.properties.statnaam
    ?? provinces.map(province => [province, Math.hypot(...geoCentroid(province).map((v, i) => v - point[i]))]).sort((a, b) => a[1] - b[1])[0][0].properties.statnaam;

const gemeenten = features("gemeenten_2023.json").map(gemeente => {
    const centroid = geoCentroid(gemeente);
    const value = Math.round(40 + (centroid[1] - 50.75) * 18 + noise(gemeente.properties.statnaam) * 35);
    return { Gemeente: gemeente.properties.statnaam, Provincie: provinceOf(centroid), Score: value };
}).sort((a, b) => a.Gemeente.localeCompare(b.Gemeente, "nl"));

const provincieRows = provinces.map(province => {
    const rows = gemeenten.filter(row => row.Provincie === province.properties.statnaam);
    return {
        Provincie: province.properties.statnaam,
        Score: Math.round(rows.reduce((sum, row) => sum + row.Score, 0) / Math.max(1, rows.length)),
        Gemeenten: rows.length
    };
});

const capitals = ["Groningen", "Leeuwarden", "Assen", "Zwolle", "Lelystad", "Arnhem", "Utrecht", "Haarlem", "'s-Gravenhage", "Middelburg", "'s-Hertogenbosch", "Maastricht"];
const extraPoints = ["Amsterdam", "Rotterdam", "Eindhoven", "Tilburg", "Almere", "Breda", "Nijmegen", "Apeldoorn", "Enschede", "Amersfoort", "Zaanstad", "Haarlemmermeer", "Den Helder", "Emmen", "Venlo", "Deventer", "Leiden", "Dordrecht", "Alkmaar", "Hoorn", "Sittard-Geleen", "Heerenveen", "Goes", "Terneuzen"];
const gemeenteFeatures = features("gemeenten_2023.json");
const punten = [...capitals.map(city => [city, "Provinciehoofdstad"]), ...extraPoints.map(city => [city, "Grote gemeente"])].map(([city, type]) => {
    const match = gemeenteFeatures.find(gemeente => gemeente.properties.statnaam === city);
    if (!match) throw new Error(`Unknown gemeente ${city}`);
    const [lon, lat] = geoCentroid(match);
    return { Locatie: city, Type: type, Latitude: Number(lat.toFixed(5)), Longitude: Number(lon.toFixed(5)) };
});

const csv = rows => {
    const columns = Object.keys(rows[0]);
    return [columns.join(";"), ...rows.map(row => columns.map(column => row[column]).join(";"))].join("\n") + "\n";
};
write(path.join(here, "gemeenten-sample.csv"), csv(gemeenten));
write(path.join(here, "provincies-sample.csv"), csv(provincieRows));
write(path.join(here, "punten-sample.csv"), csv(punten));
write(path.join(here, "sample-data.json"), { gemeenten, provincies: provincieRows, punten });

// ---------- semantic model ----------
fs.rmSync(out, { recursive: true, force: true });
const tables = { Gemeenten: gemeenten, Provincies: provincieRows, Punten: punten };
const isInt = (rows, column) => rows.every(row => Number.isInteger(row[column]));
const mType = (rows, column) => typeof rows[0][column] === "number" ? (isInt(rows, column) ? "Int64.Type" : "number") : "text";
const mValue = value => typeof value === "number" ? String(value) : `"${String(value).replace(/"/g, '""')}"`;
const tableTmdl = (table, rows, measures) => {
    const columns = Object.keys(rows[0]);
    const mTable = `#table(type table [${columns.map(column => `${column} = ${mType(rows, column)}`).join(", ")}], {\n${rows.map(row => `\t\t\t\t\t{${columns.map(column => mValue(row[column])).join(", ")}}`).join(",\n")}\n\t\t\t\t})`;
    const columnTmdl = column => {
        const numeric = typeof rows[0][column] === "number";
        const category = column === "Latitude" ? "Latitude" : column === "Longitude" ? "Longitude" : column === "Provincie" ? "StateOrProvince" : column === "Gemeente" || column === "Locatie" ? "City" : null;
        return [
            `\tcolumn ${column}`,
            `\t\tdataType: ${numeric ? (isInt(rows, column) ? "int64" : "double") : "string"}`,
            `\t\tlineageTag: ${stableId(table + column)}`,
            `\t\tsummarizeBy: ${numeric && !category ? "sum" : "none"}`,
            `\t\tsourceColumn: ${column}`,
            category ? `\t\tdataCategory: ${category}` : null,
            "",
            `\t\tannotation SummarizationSetBy = Automatic`,
            ""
        ].filter(line => line != null).join("\n");
    };
    return `table ${table}
\tlineageTag: ${stableId("table" + table)}

${measures.map(([measure, expression]) => `\tmeasure '${measure}' = ${expression}\n\t\tformatString: #,0\n\t\tlineageTag: ${stableId(table + "measure" + measure)}\n`).join("\n")}
${columns.map(columnTmdl).join("\n")}
\tpartition ${table} = m
\t\tmode: import
\t\tsource = \`\`\`
\t\t\t\tlet
\t\t\t\t    Source = ${mTable}
\t\t\t\tin
\t\t\t\t    Source
\t\t\t\t\`\`\`

\tannotation PBI_ResultType = Table
`;
};
write(path.join(model, "definition", "tables", "Gemeenten.tmdl"), tableTmdl("Gemeenten", gemeenten, [["Gemiddelde score", "AVERAGE ( Gemeenten[Score] )"]]));
write(path.join(model, "definition", "tables", "Provincies.tmdl"), tableTmdl("Provincies", provincieRows, [["Provinciescore", "AVERAGE ( Provincies[Score] )"]]));
write(path.join(model, "definition", "tables", "Punten.tmdl"), tableTmdl("Punten", punten, []));
write(path.join(model, "definition", "relationships.tmdl"), `relationship ${stableId("relGemProv")}
\tfromColumn: Gemeenten.Provincie
\ttoColumn: Provincies.Provincie
`);
write(path.join(model, "definition", "model.tmdl"), `model Model
\tculture: nl-NL
\tdefaultPowerBIDataSourceVersion: powerBI_V3
\tsourceQueryCulture: nl-NL
\tdataAccessOptions
\t\tlegacyRedirects
\t\treturnErrorValuesAsNull

annotation __PBI_TimeIntelligenceEnabled = 0

annotation PBI_QueryOrder = ${JSON.stringify(Object.keys(tables))}

${Object.keys(tables).map(table => `ref table ${table}`).join("\n")}
`);
write(path.join(model, "definition", "database.tmdl"), "database\n\tcompatibilityLevel: 1606\n");
write(path.join(model, "definition.pbism"), {
    $schema: "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
    version: "4.2",
    settings: {}
});
write(path.join(model, ".platform"), {
    $schema: "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
    metadata: { type: "SemanticModel", displayName: name },
    config: { version: "2.0", logicalId: randomUUID() }
});

// ---------- report ----------
const pbiviz = JSON.parse(fs.readFileSync(path.join(root, "pbiviz.json"), "utf8"));
const guid = pbiviz.visual.guid;
const packageFile = fs.existsSync(path.join(root, "dist")) && fs.readdirSync(path.join(root, "dist")).find(file => file.startsWith(guid) && file.endsWith(".pbiviz"));
if (!packageFile) throw new Error("Run `npx pbiviz package` first.");
const visualFolder = path.join(report, "CustomVisuals", guid);
fs.mkdirSync(visualFolder, { recursive: true });
execFileSync("tar", ["-xf", path.join(root, "dist", packageFile), "-C", visualFolder]);
// report.json paths are basenames; Power BI resolves them via package.json → resources/.
const resources = fs.readdirSync(path.join(visualFolder, "resources"));

const literal = value => ({ expr: { Literal: { Value: typeof value === "string" ? `'${value}'` : typeof value === "number" ? `${value}D` : String(value) } } });
const objects = spec => Object.fromEntries(Object.entries(spec).map(([object, properties]) => [object, [{
    properties: Object.fromEntries(Object.entries(properties).map(([key, value]) => [key, typeof value === "string" && value.startsWith("#") ? { solid: { color: literal(value) } } : literal(value)]))
}]]));
const column = (entity, property) => ({ field: { Column: { Expression: { SourceRef: { Entity: entity } }, Property: property } }, queryRef: `${entity}.${property}`, nativeQueryRef: property, active: true });
const measure = (entity, property) => ({ field: { Measure: { Expression: { SourceRef: { Entity: entity } }, Property: property } }, queryRef: `${entity}.${property}`, nativeQueryRef: property });
const title = text => ({ title: [{ properties: { show: literal(true), text: literal(text), fontSize: literal(12) } }] });

const pages = [];
const page = (id, displayName, visuals) => {
    pages.push(id);
    const folder = path.join(report, "definition", "pages", id);
    write(path.join(folder, "page.json"), {
        $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json",
        name: id, displayName, displayOption: "FitToPage", height: 768, width: 1366
    });
    visuals.forEach((visual, index) => write(path.join(folder, "visuals", visual.name, "visual.json"), {
        $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.4.0/schema.json",
        name: visual.name,
        position: { ...visual.position, z: index * 1000, tabOrder: index * 1000 },
        visual: {
            visualType: visual.type,
            query: { queryState: visual.query },
            ...(visual.objects ? { objects: objects(visual.objects) } : {}),
            visualContainerObjects: title(visual.title),
            drillFilterOtherVisuals: true
        }
    }));
};
const map = (name, titleText, position, query, settings) => ({ name, type: guid, title: titleText, position, query, objects: settings });
const slicer = (name, entity, property, position) => ({
    name, type: "slicer", title: property, position,
    query: { Values: { projections: [column(entity, property)] } }
});

page("p1gemeenten00000001", "Gemeenten", [
    map("map_gemeenten", "Score per gemeente (fictieve waarden)", { x: 16, y: 16, width: 1000, height: 736 }, {
        category: { projections: [column("Gemeenten", "Gemeente")] },
        measure: { projections: [measure("Gemeenten", "Gemiddelde score")] }
    }, { mapSettings: { selectedMapSlice: "gemeenten 2023", useMidColorSlice: true } }),
    slicer("slicer_provincie", "Provincies", "Provincie", { x: 1032, y: 16, width: 318, height: 420 })
]);
page("p2provincies0000002", "Provincies", [
    map("map_provincies", "Gemiddelde score per provincie", { x: 16, y: 16, width: 1334, height: 736 }, {
        category: { projections: [column("Provincies", "Provincie")] },
        measure: { projections: [measure("Provincies", "Provinciescore")] }
    }, { mapSettings: { selectedMapSlice: "provincies" }, labelSettings: { showLabels: true, showValues: true, showRegionName: true } })
]);
page("p3punten000000000003", "Punten", [
    map("map_punten", "Locaties als punten op de kaart", { x: 16, y: 16, width: 1334, height: 736 }, {
        latitude: { projections: [column("Punten", "Latitude")] },
        longitude: { projections: [column("Punten", "Longitude")] },
        pointCategory: { projections: [column("Punten", "Type")] },
        pointTooltip: { projections: [column("Punten", "Locatie")] }
    }, { mapSettings: { selectedMapSlice: "provincies" }, pointSettings: { showPoints: true, useLegendColors: true, showPointLegend: true, pointSize: 8 } })
]);

write(path.join(report, "definition", "pages", "pages.json"), {
    $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
    pageOrder: pages, activePageName: pages[0]
});
write(path.join(report, "definition", "version.json"), {
    $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
    version: "2.0.0"
});
write(path.join(report, "definition", "report.json"), {
    $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.3.0/schema.json",
    themeCollection: { baseTheme: { name: "CY24SU10", reportVersionAtImport: { visual: "1.8.95", report: "2.0.95", page: "1.3.95" }, type: "SharedResources" } },
    resourcePackages: [
        { name: "SharedResources", type: "SharedResources", items: [{ name: "CY24SU10", path: "BaseThemes/CY24SU10.json", type: "BaseTheme" }] },
        {
            name: guid, type: "CustomVisual", items: [
                ...resources.map(file => ({ name: file, path: file, type: "CustomVisualMetadata" }))
            ]
        }
    ],
    settings: { useStylableVisualContainerHeader: true, exportDataMode: "AllowSummarized", defaultDrillFilterOtherVisuals: true }
});
write(path.join(report, "definition.pbir"), {
    $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
    version: "4.0",
    datasetReference: { byPath: { path: `../${name}.SemanticModel` } }
});
write(path.join(report, ".platform"), {
    $schema: "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
    metadata: { type: "Report", displayName: name },
    config: { version: "2.0", logicalId: randomUUID() }
});
write(path.join(out, `${name}.pbip`), {
    $schema: "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
    version: "1.0",
    artifacts: [{ report: { path: `${name}.Report` } }],
    settings: { enableAutoRecovery: true }
});
console.log(`Wrote ${path.join(out, name + ".pbip")} (${gemeenten.length} gemeenten, ${punten.length} punten)`);
