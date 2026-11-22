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
const provinceOf = point => {
    const match = provinces.find(province => geoContains(province, point))
        ?? provinces.map(province => [province, Math.hypot(...geoCentroid(province).map((v, i) => v - point[i]))]).sort((a, b) => a[1] - b[1])[0][0];
    return { Provincie: match.properties.statnaam, ProvincieStatcode: match.properties.statcode };
};

const gemeenten = features("gemeenten_2023.json").map(gemeente => {
    const centroid = geoCentroid(gemeente);
    const value = Math.round(40 + (centroid[1] - 50.75) * 18 + noise(gemeente.properties.statnaam) * 35);
    const province = provinceOf(centroid);
    return {
        Statcode: gemeente.properties.statcode,
        Gemeente: gemeente.properties.statnaam,
        Provincie: province.Provincie,
        ProvincieStatcode: province.ProvincieStatcode,
        Score: value
    };
}).sort((a, b) => a.Gemeente.localeCompare(b.Gemeente, "nl"));

const provincieRows = provinces.map(province => {
    const rows = gemeenten.filter(row => row.ProvincieStatcode === province.properties.statcode);
    return {
        Statcode: province.properties.statcode,
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
    const province = provinceOf([lon, lat]);
    return {
        Locatie: city,
        Type: type,
        Latitude: Number(lat.toFixed(5)),
        Longitude: Number(lon.toFixed(5)),
        // CBS codes from topojson so Filter can scope points/clusters to regions
        Statcode: match.properties.statcode,
        ProvincieStatcode: province.ProvincieStatcode
    };
});

// GGD + locaties: public stand-in for the Koeltorens-style demo (points/clusters on GGD map).
// Fictitious locations only — do not publish real org pbix data to the public repo.
const ggdFeatures = features("ggdregios.json");
const ggd = ggdFeatures.map(feature => ({
    Statcode: feature.properties.statcode,
    GGD: feature.properties.statnaam
})).sort((a, b) => a.GGD.localeCompare(b.GGD, "nl"));

const statuses = ["GEACCEPTEERD", "NIEUW", "AFGEWEZEN"];
const locaties = ggdFeatures.flatMap((feature) => {
    const [lon0, lat0] = geoCentroid(feature);
    return [0, 1, 2].map(offset => {
        const jitter = (noise(`${feature.properties.statcode}-${offset}`) - 0.5) * 0.18;
        const lon = lon0 + jitter;
        const lat = lat0 + (noise(`${feature.properties.statcode}-y-${offset}`) - 0.5) * 0.12;
        const inside = geoContains(feature, [lon, lat]) ? [lon, lat] : [lon0, lat0];
        return {
            Locatie: `Locatie ${feature.properties.statcode}-${offset + 1}`,
            Type: offset === 2 ? "type_afgewezen" : offset === 1 ? "type_nieuw" : "type_geaccepteerd",
            Status: statuses[offset % statuses.length],
            Latitude: Number(inside[1].toFixed(5)),
            Longitude: Number(inside[0].toFixed(5)),
            GGDStatcode: feature.properties.statcode,
            GGD: feature.properties.statnaam
        };
    });
});

const mapCatalog = [
    ["Provincies", "provincies", "provincies.json", "PV"],
    ["Wijken", "wijken", "wijken.json", "WK"],
    ["Gemeenten 2023", "gemeenten 2023", "gemeenten_2023.json", "GM"],
    ["Gemeenten 2021", "gemeenten 2021", "gemeenten_2021.json", "GM"],
    ["Gemeenten 2020", "gemeenten 2020", "gemeenten_2020.json", "GM"],
    ["Gemeenten 2019", "gemeenten 2019", "gemeenten_2019.json", "GM"],
    ["Gemeenten 2016", "gemeenten 2016", "gemeenten_2016.json", "GM"],
    ["Gemeenten 2012", "gemeenten 2012", "gemeenten_2012.json", "GM"],
    ["GGD Regios", "ggdregios", "ggdregios.json", "GG"],
    ["Veiligheidsregios", "veiligheidsregios", "veiligheidsregios.json", "VR"],
    ["Politieregios", "politieregios", "politieregios.json", "PO"]
].map(([kaart, instelling, file, prefix], index) => {
    const regions = features(file);
    return {
        Volgnummer: index + 1,
        Kaart: kaart,
        Instelling: instelling,
        StatcodePrefix: prefix,
        Regios: regions.length,
        VoorbeeldStatcode: regions[0]?.properties?.statcode || regions[0]?.properties?.wijkcode || ""
    };
});

const csv = rows => {
    const columns = Object.keys(rows[0]);
    return [columns.join(";"), ...rows.map(row => columns.map(column => row[column]).join(";"))].join("\n") + "\n";
};
write(path.join(here, "gemeenten-sample.csv"), csv(gemeenten));
write(path.join(here, "provincies-sample.csv"), csv(provincieRows));
write(path.join(here, "punten-sample.csv"), csv(punten));
write(path.join(here, "ggd-sample.csv"), csv(ggd));
write(path.join(here, "locaties-sample.csv"), csv(locaties));
write(path.join(here, "kaarten-sample.csv"), csv(mapCatalog));
write(path.join(here, "sample-data.json"), { gemeenten, provincies: provincieRows, punten, ggd, locaties, kaarten: mapCatalog });

// ---------- semantic model ----------
fs.rmSync(out, { recursive: true, force: true });
const tables = { Kaarten: mapCatalog, Gemeenten: gemeenten, Provincies: provincieRows, Punten: punten, GGD: ggd, Locaties: locaties };
const isInt = (rows, column) => rows.every(row => Number.isInteger(row[column]));
// File.Contents gives mashup a real data source. Pure #table queries make Desktop's
// DataSourceDiscoveryVisitor NullReferenceException during refresh / composite checks.
const mCsvPath = file => path.join(here, file).replace(/\\/g, "\\\\");
const tableTmdl = (table, rows, measures, csvFile) => {
    const columns = Object.keys(rows[0]);
    const typePairs = columns.map(column => {
        const numeric = typeof rows[0][column] === "number";
        return `{"${column}", ${numeric ? (isInt(rows, column) ? "Int64.Type" : "type number") : "type text"}}`;
    }).join(", ");
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
\t\t\t\t    Source = Csv.Document(File.Contents("${mCsvPath(csvFile)}"), [Delimiter=";", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
\t\t\t\t    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
\t\t\t\t    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {${typePairs}}, "en-US")
\t\t\t\tin
\t\t\t\t    #"Changed Type"
\t\t\t\t\`\`\`

\tannotation PBI_ResultType = Table
`;
};
write(path.join(model, "definition", "tables", "Kaarten.tmdl"), tableTmdl("Kaarten", mapCatalog, [], "kaarten-sample.csv"));
write(path.join(model, "definition", "tables", "Gemeenten.tmdl"), tableTmdl("Gemeenten", gemeenten, [["Gemiddelde score", "AVERAGE ( Gemeenten[Score] )"]], "gemeenten-sample.csv"));
write(path.join(model, "definition", "tables", "Provincies.tmdl"), tableTmdl("Provincies", provincieRows, [["Provinciescore", "AVERAGE ( Provincies[Score] )"]], "provincies-sample.csv"));
write(path.join(model, "definition", "tables", "Punten.tmdl"), tableTmdl("Punten", punten, [], "punten-sample.csv"));
write(path.join(model, "definition", "tables", "GGD.tmdl"), tableTmdl("GGD", ggd, [], "ggd-sample.csv"));
write(path.join(model, "definition", "tables", "Locaties.tmdl"), tableTmdl("Locaties", locaties, [], "locaties-sample.csv"));
// Locaties → GGD so a GGD slicer filters points/clusters without putting GGD in the visual wells.
write(path.join(model, "definition", "relationships.tmdl"), `relationship Locaties_GGD
\tjoinOnDateBehavior: datePartOnly
\tfromColumn: Locaties.GGDStatcode
\ttoColumn: GGD.Statcode
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
const page = (id, displayName, visuals, interactions = []) => {
    pages.push(id);
    const folder = path.join(report, "definition", "pages", id);
    write(path.join(folder, "page.json"), {
        $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json",
        name: id, displayName, displayOption: "FitToPage", height: 768, width: 1366,
        ...(interactions.length ? { visualInteractions: interactions } : {})
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
const tableVisual = (name, titleText, position, entity, properties) => ({
    name, type: "tableEx", title: titleText, position,
    query: { Values: { projections: properties.map(property => column(entity, property)) } }
});

page("p0kaarten0000000000", "Kaarten", [
    tableVisual(
        "table_kaarten",
        "Beschikbare kaarten in Nederlandse Regiokaarten",
        { x: 16, y: 16, width: 1334, height: 736 },
        "Kaarten",
        ["Volgnummer", "Kaart", "Instelling", "StatcodePrefix", "Regios", "VoorbeeldStatcode"]
    )
]);
page("p1gemeenten00000001", "Gemeenten", [
    map("map_gemeenten", "Score per gemeente (fictieve waarden)", { x: 16, y: 16, width: 1000, height: 736 }, {
        category: { projections: [column("Gemeenten", "Statcode")] },
        measure: { projections: [measure("Gemeenten", "Gemiddelde score")] }
        // Filter uses topojson codes for the selected map (GM## here). Province slicer
        // still filters the data rows via ProvincieStatcode on the same table.
    }, { mapSettings: { selectedMapSlice: "gemeenten 2023", useMidColorSlice: true } }),
    slicer("slicer_provincie", "Gemeenten", "ProvincieStatcode", { x: 1032, y: 16, width: 318, height: 420 })
]);
page("p2provincies0000002", "Provincies", [
    map("map_provincies", "Gemiddelde score per provincie", { x: 16, y: 16, width: 1334, height: 736 }, {
        category: { projections: [column("Provincies", "Statcode")] },
        measure: { projections: [measure("Provincies", "Provinciescore")] }
    }, { mapSettings: { selectedMapSlice: "provincies" }, labelSettings: { showLabels: true, showValues: true, showRegionName: true } })
]);
page("p3punten000000000003", "Punten", [
    map("map_punten", "Locaties als punten op de kaart", { x: 16, y: 16, width: 1000, height: 736 }, {
        latitude: { projections: [column("Punten", "Latitude")] },
        longitude: { projections: [column("Punten", "Longitude")] },
        pointCategory: { projections: [column("Punten", "Type")] },
        pointTooltip: { projections: [column("Punten", "Locatie")] },
        // ProvincieStatcode matches provincies topojson.statcode (PV##)
        filter: { projections: [column("Punten", "ProvincieStatcode")] }
    }, { mapSettings: { selectedMapSlice: "provincies" }, pointSettings: { showPoints: true, useLegendColors: true, showPointLegend: true, pointSize: 8 } }),
    slicer("slicer_punten_pv", "Punten", "ProvincieStatcode", { x: 1032, y: 16, width: 318, height: 420 })
]);

// Showcase pages (Koeltorens-style): GGD basemap + points/clusters, GGD slicer via relationship
const ggdPointQuery = {
    latitude: { projections: [column("Locaties", "Latitude")] },
    longitude: { projections: [column("Locaties", "Longitude")] },
    pointCategory: { projections: [column("Locaties", "Status")] },
    pointTooltip: { projections: [column("Locaties", "Locatie"), column("Locaties", "Type")] }
};
const ggdPointSettings = {
    mapSettings: { selectedMapSlice: "ggdregios", showExtraFeatures: false, regionBorderThickness: 1 },
    pointSettings: {
        showPoints: true, pointSize: 5, useLegendColors: true, showPointLegend: true,
        pointLegendPosition: "bottom-right", enableClustering: false, pointOpacity: 100
    },
    labelSettings: { showLabels: false, showValues: false },
    legendSettings: { showLegend: false }
};
const ggdClusterSettings = {
    ...ggdPointSettings,
    pointSettings: {
        ...ggdPointSettings.pointSettings,
        enableClustering: true, showClusterCount: true, showPointLegend: false,
        minClusterSize: 20, maxClusterSize: 20, clusterBorderThickness: 4, clusterCountSize: 14
    }
};

page("p4puntdata0000000004", "Puntdata", [
    map("map_puntdata", "Locaties per GGD-regio (fictieve data)", { x: 28, y: 30, width: 715, height: 642 }, ggdPointQuery, ggdPointSettings),
    tableVisual("table_status_punten", "Status", { x: 774, y: 57, width: 156, height: 173 }, "Locaties", ["Status"]),
    slicer("slicer_ggd_punten", "GGD", "GGD", { x: 1032, y: 57, width: 300, height: 420 })
], [{ source: "table_status_punten", target: "map_puntdata", type: "DataFilter" }]);

page("p5clusters0000000005", "Clusterbubbles", [
    map("map_clusters", "Locaties per GGD-regio (clusters)", { x: 28, y: 30, width: 715, height: 642 }, ggdPointQuery, ggdClusterSettings),
    tableVisual("table_status_clusters", "Status", { x: 752, y: 57, width: 156, height: 173 }, "Locaties", ["Status"]),
    slicer("slicer_ggd_clusters", "GGD", "GGD", { x: 752, y: 249, width: 528, height: 423 })
], [{ source: "table_status_clusters", target: "map_clusters", type: "DataFilter" }]);

write(path.join(report, "definition", "pages", "pages.json"), {
    $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
    pageOrder: pages, activePageName: "p4puntdata0000000004"
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
console.log(`Wrote ${path.join(out, name + ".pbip")} (${gemeenten.length} gemeenten, ${punten.length} punten, ${locaties.length} locaties / ${ggd.length} GGD)`);
