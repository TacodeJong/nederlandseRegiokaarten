/*
*  Power BI Visual CLI
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*  MIT License
*
*  Permission is hereby granted, free of charge, to any person obtaining a copy
*  of this software and associated documentation files (the ""Software""), to deal
*  in the Software without restriction, including without limitation the rights
*  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*  copies of the Software, and to permit persons to whom the Software is
*  furnished to do so, subject to the following conditions:
*
*  The above copyright notice and this permission notice shall be included in
*  all copies or substantial portions of the Software.
*
*  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
*  THE SOFTWARE.
*/
"use strict";

import "core-js/stable";
import "./../style/visual.less";
import * as d3 from "d3";
import { zoom, ZoomBehavior, ZoomTransform } from "d3-zoom";
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import DataView = powerbi.DataView;
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";


import { VisualSettings } from "./settings";
import { select as d3Select } from "d3-selection";
import { ITooltipServiceWrapper, TooltipEventArgs, createTooltipServiceWrapper } from "powerbi-visuals-utils-tooltiputils";
import { geoPath, geoMercator, GeoPermissibleObjects, geoContains } from "d3-geo";
import { feature } from "topojson-client";
import * as topojson from 'topojson-client';

import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;

import * as provinciesTopojson from "../topojson/provincies.json";
import * as gemeentenTopojson_2023 from "../topojson/gemeenten_2023.json";
import * as gemeentenTopojson_2021 from "../topojson/gemeenten_2021.json";
import * as gemeentenTopojson_2020 from "../topojson/gemeenten_2020.json";
import * as gemeentenTopojson_2019 from "../topojson/gemeenten_2019.json";  
import * as gemeentenTopojson_2016 from "../topojson/gemeenten_2016.json";
import * as gemeentenTopojson_2012 from "../topojson/gemeenten_2012.json";


import * as wijkenTopojson from "../topojson/wijken.json";
import * as ggdregiosTopojson from "../topojson/ggdregios.json";
import * as politieregiosTopojson from "../topojson/politieregios.json";
import * as veiligheidsregiosTopojson from "../topojson/veiligheidsregios.json";
// import * as KONwerkGebiedenTopojson from "../topojson/werkgebieden.json";
// Define the interface for the GeoJSON feature properties
interface GeoJsonFeatureProperties {
    naam: string;
    rubriek: string;
    code: string;
}

// Define the interface for the GeoJSON feature
interface GeoJsonFeature {
    type: string;
    properties: GeoJsonFeatureProperties;
    geometry: {
        type: string;
        coordinates: any[];
    };
    value: number;
    customLabel: string;
    tooltipData: VisualTooltipDataItem[];
    selectionId: powerbi.visuals.ISelectionId;
    categoryIndex?: number;
    isHighlighted?: boolean;
}

// Define the type for the GeoJSON data
interface GeoJsonData {
    type: "FeatureCollection";
    features: GeoJsonFeature[];
    bbox?: number[];
}

// One mapped point. selectionId is the identity of the underlying data row, so
// clicking a mark can cross-filter other visuals down to those rows.
interface PointDatum {
    lat: number;
    lon: number;
    legend?: string;
    sourceIndex?: number;
    selectionId?: powerbi.visuals.ISelectionId;
}

interface ProjectedPointDatum extends PointDatum {
    projectedX: number;
    projectedY: number;
}

// Region values indexed for hash-based matching, with the original order kept
// for the substring fallback
interface RegionLookup {
    byName: Map<string, number>;
    byCode: Map<string, number>;
    normalized: string[];
}

// One screen-space cluster bubble, with its per-legend-category breakdown
interface ClusterDatum {
    id: string;
    count: number;
    centroid: [number, number];
    legend?: string;
    legendCounts: Map<string, number>;
    // The rows behind this bubble, so it can be filtered and partially highlighted
    points: ProjectedPointDatum[];
}

export class Visual implements IVisual {
    private target: HTMLElement;
    private settings: VisualSettings;
    private svg: d3.Selection<SVGSVGElement, any, HTMLElement, any>;
    private zoomBehavior: ZoomBehavior<SVGSVGElement, unknown>;
    private container: d3.Selection<SVGGElement, any, HTMLElement, any>;
    private tooltipServiceWrapper: ITooltipServiceWrapper;
    private featuresGroup: d3.Selection<SVGGElement, any, HTMLElement, any>;
    private labelsGroup: d3.Selection<SVGGElement, any, HTMLElement, any>;
    private pointsGroup: d3.Selection<SVGGElement, any, HTMLElement, any>;
    private measureValues: number[] = [];
    private selectionManager: powerbi.extensibility.ISelectionManager;
    private selectionIdBuilder: powerbi.extensibility.ISelectionIdBuilder;
    private tooltipMap: Map<string, VisualTooltipDataItem[]>;
    private host: powerbi.extensibility.visual.IVisualHost;
    private currentFieldSet: string = ""; // Cache for current field set
    private filterValues: string[] = [];
    private formattingSettingsService: FormattingSettingsService;
    private dataViews: powerbi.DataView[];
    private viewport: powerbi.IViewport;
    private checkbox: HTMLInputElement; // Define the checkbox property
    private legendCheckbox: HTMLInputElement; // Define the legend checkbox property
    private pointLegendCheckbox: HTMLInputElement; // Define the point legend checkbox property
    private labelsCheckbox: HTMLInputElement; // Define the labels checkbox property
    private valuesCheckbox: HTMLInputElement; // Define the values checkbox property
    private uiContainer: HTMLDivElement; // Define a container for UI elements
    private uiHeader: HTMLDivElement; // Define a header for the cockpit
    private uiContent: HTMLDivElement; // Define a content container for the cockpit
    private legendContainer: HTMLDivElement; // Define a container for the legend
    private pointLegendContainer: HTMLDivElement; // Define a container for the point legend
    private debugConsole: HTMLDivElement; // Define a container for the debug console
    private visualHostTooltipService: powerbi.extensibility.ITooltipService;
    
    // Store unfiltered values for color scale preservation
    private unfilteredMeasureValues: number[] = [];
    
    // Cache global min/max for color scale across filtering
    private scaleMin: number | null = null;
    private scaleMax: number | null = null;
    private scaleLocked: boolean = true;        // default behavior: keep scale
    private measureKey: string | null = null;   // to detect measure changes
    
    // Local state for cockpit toggles to avoid formatting pane reloads
    private localShowRegionName: boolean = true;
    private localShowLabels: boolean = true;
    private localShowValues: boolean = true;
    private localShowLegend: boolean = false;
    private localShowPointLegend: boolean = false;
    private localCockpitCollapsed: boolean = false;
    private isUsingPointCounts: boolean = false;
    private legendCategories: string[] = [];
    // Every category ever seen, in the order they claimed a colour slot
    private legendSlotOrder: string[] = [];
    private readonly maxLegendColorSlots: number = 12;
    private legendCategoryColumn: powerbi.DataViewCategoryColumn | undefined;
    // The 'Regio' column. In point mode lat/lon/category are grouping columns too,
    // so categories[0] is not necessarily the region field.
    private regionCategoryColumn: powerbi.DataViewCategoryColumn | undefined;

    // Legend categories the user clicked in the point legend; empty means "no highlight"
    private selectedLegendCategories: Set<string> = new Set<string>();
    private readonly legendDimOpacity: number = 0.12;
    // Keys of the data rows currently cross-filtered by clicking a point or cluster
    private selectedPointKeys: Set<string> | null = null;
    private dataView: powerbi.DataView;
    
    // Performance optimization caches
    private baseGeojsonCache: Map<string, {type: string, features: any[]}> = new Map();
    private normalizedKeyCache: Map<string, string> = new Map();
    private spatialIndex: Map<string, any[]> | null = null;
    private spatialIndexKey: string | null = null;
    private spatialIndexGridSize: number = 0.2; // degrees
    private cachedProjection: any = null;
    private lastRenderedProjectionKey: string | null = null;
    private projectedPointsCache: Map<string, {projectedX: number, projectedY: number}> = new Map();
    private maxPointsLimit: number = 10000; // Configurable limit for very large datasets
    
    // Cache for detecting changes that require cache invalidation
    private lastViewportSize: {width: number, height: number} | null = null;
    private lastProjectionKey: string | null = null;

    // Zoom-based point clustering state
    private clusterProjectedPoints: ProjectedPointDatum[] | null = null;
    private clusterPointTooltipValues: any[] | undefined;
    private clusterLegendFieldName: string | undefined;
    private lastClusterZoomBucket: number | null = null;
    private clusterScreenCellSize: number = 60; // screen pixels — larger = more aggressive clustering when zoomed out
    private clusterZoomRaf: number | null = null;


    private updateLabelScaling(transform: d3.ZoomTransform) {
        this.labelsGroup
            .attr("transform", transform.toString())
            .selectAll("text")
            .style("font-size", `${this.settings.labelSettings.labelSettingsGroup.fontSize.value / transform.k}px`)
            .attr("dy", (d, i) => i % 2 === 0 ? "0" : `${1.2 / transform.k}em`); 
            
        this.labelsGroup.selectAll("g.label").each((d, i, nodes) => {
            const label = d3.select(nodes[i]);
            const text = label.select("text").node() as SVGTextElement;
            const rect = label.select("rect");
            const bbox = text.getBBox();
            const padding = 3 / transform.k;
            
            rect
                .attr("x", bbox.x - padding)
                .attr("y", bbox.y - padding)
                .attr("width", bbox.width + (2 * padding))
                .attr("height", bbox.height + (2 * padding))
                .attr("rx", 5 / transform.k)
                .attr("ry", 5 / transform.k)
                .style("stroke-width", (this.settings.labelSettings.labelSettingsGroup.showLabelBorder.value.valueOf() ? 
                    ((this.settings.labelSettings.labelSettingsGroup as any).labelBorderThickness?.value ?? 1) / transform.k : 0) + "px");
        });
    }

    /**
     * Keep point/cluster screen size constant while the map zooms.
     */
    private updatePointScaling(transform: d3.ZoomTransform) {
        const k = transform.k || 1;

        this.pointsGroup.attr("transform", transform.toString());

        this.pointsGroup.selectAll<SVGCircleElement, unknown>("circle").each(function () {
            const circle = d3.select(this);
            const baseR = Number(circle.attr("data-base-r"));
            const baseStroke = Number(circle.attr("data-base-stroke"));
            if (!isNaN(baseR)) {
                circle.attr("r", baseR / k);
            }
            if (!isNaN(baseStroke)) {
                circle.attr("stroke-width", baseStroke / k);
            }
        });

        this.pointsGroup.selectAll<SVGTextElement, unknown>("text").each(function () {
            const text = d3.select(this);
            const baseFont = Number(text.attr("data-base-font"));
            if (!isNaN(baseFont)) {
                text.attr("font-size", baseFont / k);
            }
        });
    }

    private onZoom(event: d3.D3ZoomEvent<SVGSVGElement, unknown>) {
        this.featuresGroup.attr("transform", event.transform.toString());
        this.updateLabelScaling(event.transform);

        // Re-cluster on zoom so bubbles merge when zooming out and split when zooming in
        if (this.clusterProjectedPoints && this.clusterProjectedPoints.length > 0
            && this.settings?.pointSettings?.clusterSettingsGroup?.enableClustering?.value) {
            if (this.clusterZoomRaf !== null) {
                cancelAnimationFrame(this.clusterZoomRaf);
            }
            this.clusterZoomRaf = requestAnimationFrame(() => {
                this.clusterZoomRaf = null;
                this.drawZoomClusters(event.transform, false);
            });
        } else {
            this.updatePointScaling(event.transform);
        }
    }

    private getClusterZoomBucket(k: number): number {
        // Coarse buckets so we don't rebuild on every tiny zoom tick
        return Math.round(Math.log2(Math.max(k, 1)) * 4);
    }

    private updateZoomExtents(width: number, height: number) {
        this.zoomBehavior
            .translateExtent([[0, 0], [width, height]])
            .extent([[0, 0], [width, height]]);
    }

    constructor(options: VisualConstructorOptions) {
        const width = options.element.clientWidth;
        const height = options.element.clientHeight;
        const margin = 20; // Add margin of 20 pixels

        this.formattingSettingsService = new FormattingSettingsService();
        this.target = options.element;
        this.svg = d3Select(this.target)
            .append('svg');
        this.svg
            .attr("width", width)
            .attr("height", height)
            .attr("viewBox", `-${margin} -${margin} ${width + 2*margin} ${height + 2*margin}`)
            .attr("preserveAspectRatio", "xMinYMin meet");
        
        // Separate groups for features, labels, and points
        this.featuresGroup = this.svg.append('g').classed('features', true); // Group for features
        this.pointsGroup = this.svg.append('g').classed('points', true);     // Group for points
        // Labels sit above regions; disable hit-testing so clicks reach region paths
        this.labelsGroup = this.svg.append('g')
            .classed('labels', true)
            .style('pointer-events', 'none');

        // Initialize the zoom behavior
        this.zoomBehavior = zoom<SVGSVGElement, unknown>()
            .scaleExtent([1, 8]) // Set the zoom scale limits (min and max zoom)
            .translateExtent([[0, 0], [options.element.clientWidth, options.element.clientHeight]]) // Restrict panning
            .on("zoom", (event) => this.onZoom(event)); // Handle zoom events

        this.svg.call(this.zoomBehavior); // Attach the zoom behavior to the SVG

        this.tooltipServiceWrapper = createTooltipServiceWrapper(
            options.host.tooltipService,
            options.element
        );
        this.host = options.host;
        this.selectionManager = this.host.createSelectionManager(); // Initialize selection manager
        this.selectionIdBuilder = this.host.createSelectionIdBuilder();
        this.visualHostTooltipService = this.host.tooltipService;

        this.updateZoomExtents(options.element.clientWidth, options.element.clientHeight);
        this.tooltipMap = new Map<string, VisualTooltipDataItem[]>();  // Initialize tooltipMap

        // Create a container for UI elements (cockpit)
        this.uiContainer = d3Select(this.target)
            .append("div")
            .style("position", "absolute")
            .style("top", "20px")
            .style("left", "5px")
            .style("width", "200px")
            .style("background-color", "rgba(255, 255, 255, 0.95)")
            .style("border", "0.2px solid #ccc")
            .style("border-radius", "5px")
            .style("box-shadow", "0 0 5px rgba(0, 0, 0, 0.3)")
            .style("display", "none")
            .node() as HTMLDivElement;

        // Create header for the cockpit
        this.uiHeader = d3Select(this.uiContainer)
            .append("div")
            .style("display", "flex")
            .style("justify-content", "space-between")
            .style("align-items", "center")
            .style("padding", "1px 1px")
            .style("border-bottom", "0.2px solid #eee")
            .style("cursor", "pointer")
            .style("user-select", "none")
            .style("background-color", "#f8f9fa")
            .style("border-radius", "5px 5px 0 0")
            .node() as HTMLDivElement;

        // Add title to header
        d3Select(this.uiHeader)
            .append("span")
            .style("font-weight", "bold")
            .style("font-size", "12px")
            .style("color", "#333")
            .text("Cockpit");

        // Add collapse/expand button to header
        const collapseButton = d3Select(this.uiHeader)
            .append("span")
            .style("font-size", "14px")
            .style("color", "#666")
            .style("transition", "transform 0.2s ease")
            .text("▼")
            .node() as HTMLSpanElement;

        // Create content container
        this.uiContent = d3Select(this.uiContainer)
            .append("div")
            .style("padding", "10px")
            .style("transition", "max-height 0.3s ease, opacity 0.3s ease")
            .style("overflow", "hidden")
            .node() as HTMLDivElement;

        // Add click handler to header for collapse/expand
        d3Select(this.uiHeader).on("click", () => {
            this.toggleCockpitCollapse();
        });

        // Add checkbox for toggling labels visibility (first)
        const labelsDiv = d3Select(this.uiContent)
            .append("div")
            .style("margin-bottom", "5px");

        // Add a checkbox to toggle region names (second)
        const regionNamesDiv = d3Select(this.uiContent)
            .append("div")
            .style("margin-bottom", "5px");

        // Add checkbox for toggling values (third)
        const valuesDiv = d3Select(this.uiContent)
            .append("div")
            .style("margin-bottom", "5px");
            
        // Add checkbox for toggling legend visibility (fourth)
        const legendDiv = d3Select(this.uiContent)
            .append("div")
            .style("margin-bottom", "5px");

        // Add checkbox for toggling point legend visibility (fifth)
        const pointLegendDiv = d3Select(this.uiContent)
            .append("div")
            .style("margin-bottom", "5px");

        labelsDiv.append("label")
            .text("Toon labels")
            .style("margin-right", "5px");

        regionNamesDiv.append("label")
            .text("Toon regionamen")
            .style("margin-right", "5px");

        valuesDiv.append("label")
            .text("Toon waarden")
            .style("margin-right", "5px");

        legendDiv.append("label")
            .text("Toon legenda")
            .style("margin-right", "5px");

        pointLegendDiv.append("label")
            .text("Toon puntlegenda")
            .style("margin-right", "5px");

        // Add checkbox for toggling labels visibility
        this.labelsCheckbox = labelsDiv.append("input")
            .attr("type", "checkbox")
            .on("change", (event) => {
                this.toggleLabelsVisibility(event.target.checked);
            })
            .node() as HTMLInputElement;

        this.checkbox = regionNamesDiv.append("input")
            .attr("type", "checkbox")
            .on("change", (event) => {
                this.toggleRegionNames(event.target.checked);
            })
            .node() as HTMLInputElement;

        // Add checkbox for toggling values visibility
        this.valuesCheckbox = valuesDiv.append("input")
            .attr("type", "checkbox")
            .on("change", (event) => {
                this.toggleValuesVisibility(event.target.checked);
            })
            .node() as HTMLInputElement;

        // Add checkbox for toggling legend visibility
        this.legendCheckbox = legendDiv.append("input")
            .attr("type", "checkbox")
            .attr("checked", false) // Start unchecked since legend starts hidden
            .on("change", (event) => {
                this.toggleLegendVisibility(event.target.checked);
            })
            .node() as HTMLInputElement;

        // Add checkbox for toggling point legend visibility
        this.pointLegendCheckbox = pointLegendDiv.append("input")
            .attr("type", "checkbox")
            .attr("checked", false) // Start unchecked since point legend starts hidden
            .on("change", (event) => {
                this.togglePointLegendVisibility(event.target.checked);
            })
            .node() as HTMLInputElement;

        // Create a container for the gradient legend
        this.legendContainer = d3Select(this.target)
            .append("div")
            .style("position", "absolute")
            .style("bottom", "10px")
            .style("left", "10px")
            .style("background-color", "rgba(255, 255, 255, 0.8)")
            .style("padding", "10px")
            .style("border", "0.2px solid #000000")
            .style("box-shadow", "0 0 2px rgba(0, 0, 0, 0.2)")
            .style("display", "none")
            .node() as HTMLDivElement;

        // Create a container for the point legend
        this.pointLegendContainer = d3Select(this.target)
            .append("div")
            .style("position", "absolute")
            .style("bottom", "10px")
            .style("right", "10px")
            .style("background-color", "rgba(255, 255, 255, 0.8)")
            .style("padding", "10px")
            .style("border", "0.2px solid #000000")
            .style("box-shadow", "0 0 2px rgba(0, 0, 0, 0.2)")
            .style("display", "none")
            .node() as HTMLDivElement;

        // Create a container for the debug console
        this.debugConsole = d3Select(this.target)
            .append("div")
            .style("position", "absolute")
            .style("bottom", "10px")
            .style("right", "10px")
            .style("width", "300px")
            .style("height", "150px")
            .style("overflow-y", "auto")
            .style("background-color", "rgba(0, 0, 0, 0.7)")
            .style("color", "white")
            .style("padding", "10px")
            .style("font-size", "12px")
            .style("box-shadow", "0 0 5px rgba(0, 0, 0, 0.3)")
            .style("display", "none")
            .style("z-index", "1000")
            .node() as HTMLDivElement;

    }

    private updateLegendPositioning() {
        if (!this.legendContainer) return;

        const position = this.settings.legendSettings.legendSettingsGroup.legendPosition.value.value as string;
        // Calculate width based on margins - will be updated after SVG is created
        const useAbsoluteValues = ((this.settings.mapSettings.mapColorSettingsGroup as any).useAbsoluteValues?.value) ?? false;
        const marginRight = useAbsoluteValues ? 40 : 70;
        const marginLeft = useAbsoluteValues ? 0 : 50;
        const gradientWidth = 20;
        const width = marginLeft + gradientWidth + marginRight + 20; // Add padding (10px on each side)
        const height = 170; // Fixed height
        const backgroundColor = "#ffffff";
        const backgroundOpacity = 0.8;
        const borderColor = "#000000";
        const borderThickness = 0.2;

        // Update positioning
        d3Select(this.legendContainer)
            .style("width", `${width}px`)
            .style("height", `${height}px`)
            .style("background-color", `rgba(${this.hexToRgb(backgroundColor)}, ${backgroundOpacity})`)
            .style("border", `${borderThickness}px solid ${borderColor}`)
            .style("top", position.includes("top") ? "10px" : "auto")
            .style("bottom", position.includes("bottom") ? "10px" : "auto")
            .style("left", position.includes("left") ? "10px" : "auto")
            .style("right", position.includes("right") ? "10px" : "auto");
    }

    private hexToRgb(hex: string): string {
        const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
        if (result) {
            const r = parseInt(result[1], 16);
            const g = parseInt(result[2], 16);
            const b = parseInt(result[3], 16);
            return `${r}, ${g}, ${b}`;
        }
        return "255, 255, 255";
    }

    private createLegend(forceShow?: boolean) {
		// Clear existing legend content
            d3Select(this.legendContainer).selectAll("*").remove();

		// Check if legend should be shown based on formatting pane setting or force override
		const shouldShow = forceShow !== undefined ? forceShow : this.settings.legendSettings.legendSettingsGroup.showLegend.value;
		
		if (!shouldShow) {
			d3Select(this.legendContainer).style("display", "none");
			return;
		}
		
		// Show legend
		d3Select(this.legendContainer).style("display", "block");

		// If no measure values, hide legend and exit
		if (!this.unfilteredMeasureValues || this.unfilteredMeasureValues.length === 0) {
			d3Select(this.legendContainer).style("display", "none");
			return;
		}

        // Update legend positioning and styling
        this.updateLegendPositioning();

            // Get colors and positions from settings first
            const minColor = this.settings.mapSettings.mapColorSettingsGroup.minColorSlice.value.value;
            const midColor = this.settings.mapSettings.mapColorSettingsGroup.midColorSlice.value.value;
            const maxColor = this.settings.mapSettings.mapColorSettingsGroup.maxColorSlice.value.value;
            const useMidColor = this.settings.mapSettings.mapColorSettingsGroup.useMidColorSlice.value;
            const useAbsoluteValues = ((this.settings.mapSettings.mapColorSettingsGroup as any).useAbsoluteValues?.value) ?? false;

            const gradientHeight = 150;
            const gradientWidth = 20;
            const marginRight = useAbsoluteValues ? 40 : 70; // Space for right axis labels (smaller for absolute values, larger for percentage mode with data values)
            const marginLeft = useAbsoluteValues ? 0 : 50; // Space for left axis labels (only needed for percentages, increased to fit "100%")
            const axisLineSpacing = 2; // Thin spacing between gradient bar and axes

            const svg = d3Select(this.legendContainer)
                .append("svg")
                .attr("width", marginLeft + gradientWidth + marginRight)
                .attr("height", gradientHeight + 30)
                .style("background-color", "rgb(255, 255, 255)");

            const minPos = this.settings.mapSettings.mapColorSettingsGroup.minPosition.value;
            const midPos = this.settings.mapSettings.mapColorSettingsGroup.midPosition.value;
            const maxPos = this.settings.mapSettings.mapColorSettingsGroup.maxPosition.value;

		// Use cached domain when locked
		const minValue = (this.scaleLocked && this.scaleMin != null)
			? this.scaleMin
			: Math.min(...this.unfilteredMeasureValues);

		const maxValue = (this.scaleLocked && this.scaleMax != null)
			? this.scaleMax
			: Math.max(...this.unfilteredMeasureValues);

            // Convert positions to percentages if using absolute values
            let minPosPercent, midPosPercent, maxPosPercent;
            if (useAbsoluteValues && maxPos !== minPos) {
                // When using absolute values, position stops proportionally based on their values
                // Gradient goes from y1="100%" (bottom) to y2="0%" (top)
                // In SVG gradients: offset="0%" is at start (bottom), offset="100%" is at end (top)
                // So minPos should be at offset 0% (bottom) and maxPos at offset 100% (top)
                const range = maxPos - minPos;
                minPosPercent = 0;   // minPos always at bottom (offset 0% = start point)
                maxPosPercent = 100; // maxPos always at top (offset 100% = end point)
                if (useMidColor && range !== 0) {
                    // Calculate where midPos falls proportionally between minPos and maxPos
                    // Then convert to offset: 0% = bottom (minPos), 100% = top (maxPos)
                    midPosPercent = ((midPos - minPos) / range) * 100;
                } else {
                    midPosPercent = 50; // Default middle if not used
                }
            } else if (useAbsoluteValues) {
                // Edge case: minPos === maxPos
                minPosPercent = 0;
                maxPosPercent = 100;
                midPosPercent = 50;
            } else {
                // Use positions as percentages (current behavior)
                minPosPercent = minPos;
                midPosPercent = midPos;
                maxPosPercent = maxPos;
            }

            // Create gradient
            const gradient = svg.append("defs")
                .append("linearGradient")
                .attr("id", "legendGradient")
                .attr("x1", "0%")
                .attr("x2", "0%")
                .attr("y1", "100%")
                .attr("y2", "0%");

            // SVG gradients require stops to be ordered from smallest to largest offset
            // Create stops array and sort by offset
            if (useMidColor) {
                const stops = [
                    { offset: minPosPercent, color: minColor },
                    { offset: midPosPercent, color: midColor },
                    { offset: maxPosPercent, color: maxColor }
                ].sort((a, b) => a.offset - b.offset); // Sort by offset ascending
                
                stops.forEach(stop => {
                gradient.append("stop")
                        .attr("offset", `${stop.offset}%`)
                        .attr("stop-color", stop.color);
                });
            } else {
                const stops = [
                    { offset: minPosPercent, color: minColor },
                    { offset: maxPosPercent, color: maxColor }
                ].sort((a, b) => a.offset - b.offset); // Sort by offset ascending
                
                stops.forEach(stop => {
                gradient.append("stop")
                        .attr("offset", `${stop.offset}%`)
                        .attr("stop-color", stop.color);
                });
            }

            // Create the gradient rectangle (moved right to make space for left axis)
            const gradientX = marginLeft;
            svg.append("rect")
                .attr("x", gradientX)
                .attr("y", 5)
                .attr("width", gradientWidth)
                .attr("height", gradientHeight)
                .style("fill", "url(#legendGradient)")
                .style("stroke", "black")
                .style("stroke-width", "0.5px");
            
            // Add thin vertical line on the left side of gradient bar
            svg.append("line")
                .attr("x1", gradientX)
                .attr("x2", gradientX)
                .attr("y1", 5)
                .attr("y2", gradientHeight + 5)
                .style("stroke", "black")
                .style("stroke-width", "0.5px");
            
            // Add thin vertical line on the right side of gradient bar
            svg.append("line")
                .attr("x1", gradientX + gradientWidth)
                .attr("x2", gradientX + gradientWidth)
                .attr("y1", 5)
                .attr("y2", gradientHeight + 5)
                .style("stroke", "black")
                .style("stroke-width", "0.5px");

            // Create scales for both axes
            // When using absolute values, the right axis needs to position the manually set values
            // based on where they fall in the gradient (which uses percentages)
            let valueScale: d3.ScaleLinear<number, number>;
            let rightAxisTickValues: number[];
            
            if (useAbsoluteValues) {
                // When using absolute values, create a scale that maps minPos to bottom and maxPos to top
                // The scale domain should be [minPos, maxPos] to position them correctly
                valueScale = d3.scaleLinear()
                    .domain([minPos, maxPos]) // minPos at bottom, maxPos at top
                    .range([gradientHeight + 5, 5]); // bottom to top
                
                // Show ticks at the manually set positions (ordered from min to max)
                if (useMidColor) {
                    rightAxisTickValues = [minPos, midPos, maxPos];
                } else {
                    rightAxisTickValues = [minPos, maxPos];
                }
            } else {
                // When using percentages, use the data range
                valueScale = d3.scaleLinear()
                .domain([minValue, maxValue])
                .range([gradientHeight + 5, 5]);

                // Show the actual data values
                rightAxisTickValues = useMidColor ? 
                    [minValue, (minValue + maxValue) / 2, maxValue] : 
                    [minValue, maxValue];
            }

            // Create left axis scale - only show when NOT using absolute values
            let leftAxisScale: d3.ScaleLinear<number, number>;
            let leftAxisTickValues: number[];
            let leftAxisFormat: (d: number) => string;
            let showLeftAxis = !useAbsoluteValues; // Hide left axis when using absolute values
            
            if (!useAbsoluteValues) {
                // When using percentages, show percentages
                leftAxisScale = d3.scaleLinear()
                .domain([0, 100])
                .range([gradientHeight + 5, 5]);
                leftAxisTickValues = [0, 50, 100];
                leftAxisFormat = (d: number) => d + "%";
            } else {
                // Left axis not used in absolute values mode, but need to define for TypeScript
                leftAxisScale = d3.scaleLinear()
                    .domain([0, 100])
                    .range([gradientHeight + 5, 5]);
                leftAxisTickValues = [];
                leftAxisFormat = (d: number) => "";
            }

            const valueAxis = d3.axisRight(valueScale)
                .tickValues(rightAxisTickValues)
                .tickFormat(d => Number(d).toFixed(this.settings.labelSettings.labelSettingsGroup.numOfDigits.value));

            const leftAxis = d3.axisLeft(leftAxisScale)
                .tickValues(leftAxisTickValues)
                .tickFormat(leftAxisFormat);

            // Get font settings
            const fontSize = 12; // Fixed font size
            const fontColor = "#000000"; // Fixed font color

            // Add left axis - only show when NOT using absolute values
            if (showLeftAxis) {
            svg.append("g")
                .attr("class", "legend-axis")
                    .attr("transform", `translate(${gradientX - axisLineSpacing}, 0)`)
                    .call(leftAxis)
                .style("font-size", `${fontSize}px`)
                .style("fill", fontColor);
            }

            // Add value axis (right side) - positioned just after the gradient bar
            svg.append("g")
                .attr("class", "legend-axis")
                .attr("transform", `translate(${gradientX + gradientWidth + axisLineSpacing}, 0)`)
                .call(valueAxis)
                .style("font-size", `${fontSize}px`)
                .style("fill", fontColor);

            // Style the axis lines and ticks - make them thin lines
            // Hide the domain path (the vertical line) - we only want tick lines and labels
            svg.selectAll(".legend-axis path")
                .style("stroke", "none")
                .style("display", "none");
            // Style tick lines as thin lines
            svg.selectAll(".legend-axis line")
                .style("stroke", "black")
                .style("stroke-width", "0.5px");
        }
    
    private createPointLegend(forceShow?: boolean) {
        // Clear existing point legend content
        d3Select(this.pointLegendContainer).selectAll("*").remove();
        
        // Check if point legend should be shown based on formatting pane setting or force override
        const shouldShow = forceShow !== undefined ? forceShow : this.settings.pointSettings.pointSettingsGroup.showPointLegend.value;
        
        if (!shouldShow) {
            d3Select(this.pointLegendContainer).style("display", "none");
            return;
        }
        
        // Show point legend
        d3Select(this.pointLegendContainer).style("display", "block");
        
        // If no legend categories or legend colors not enabled, hide and exit
        if (this.legendCategories.length === 0 || !this.settings.pointSettings.pointSettingsGroup.useLegendColors.value) {
            d3Select(this.pointLegendContainer).style("display", "none");
            return;
        }
        
        // Update point legend positioning and styling
        this.updatePointLegendPositioning();
        
        // Show the container
        d3Select(this.pointLegendContainer).style("display", "block");
        
        // Use fixed styling
        const fontSize = 11;
        const fontColor = "#000000";
        
        // Create title
        d3Select(this.pointLegendContainer)
            .append("div")
            .style("font-weight", "bold")
            .style("margin-bottom", "8px")
            .style("font-size", `${fontSize + 1}px`)
            .style("color", fontColor)
            .text("Puntlegenda");
        
        // Create legend items
        this.legendCategories.forEach((category, index) => {
            const color = this.getLegendColor(category);
            
            const itemDiv = d3Select(this.pointLegendContainer)
                .append("div")
                .classed("point-legend-item", true)
                .attr("data-category", category)
                .attr("title", "Klik om deze categorie te markeren (ctrl-klik voor meerdere)")
                .style("display", "flex")
                .style("align-items", "center")
                .style("margin-bottom", "4px")
                .style("cursor", "pointer")
                .style("user-select", "none")
                .on("click", (event: MouseEvent) => {
                    event.stopPropagation();
                    this.toggleLegendCategory(category, event.ctrlKey || event.metaKey || event.shiftKey);
                });
            
            // Color circle
            itemDiv.append("div")
                .style("width", "12px")
                .style("height", "12px")
                .style("border-radius", "50%")
                .style("background-color", color)
                .style("margin-right", "6px")
                .style("flex-shrink", "0");
            
            // Category label
            itemDiv.append("div")
                .style("font-size", `${fontSize}px`)
                .style("color", fontColor)
                .style("white-space", "nowrap")
                .text(category);
        });

        this.updatePointLegendSelectionStyles();
    }

    /**
     * Cross-filter on a region. Returns false when the region has no matching row in
     * the 'Regio' field, so the caller can let the click fall through.
     */
    private selectRegionFeature(event: MouseEvent, feature: GeoJsonFeature): boolean {
        if (!feature || !feature.selectionId) {
            return false;
        }

        const isCtrlPressed = event.ctrlKey || event.metaKey;
        this.selectionManager.select(feature.selectionId, isCtrlPressed)
            .then((ids: powerbi.visuals.ISelectionId[]) => {
                // Selecting a region replaces any point-level selection
                this.setSelectedPointKeys([]);
                this.updateSelection(ids);
                if (ids.length === 0) {
                    const borderSettings = this.settings.mapSettings.regionBorderSettingsGroup;
                    this.featuresGroup.selectAll<SVGPathElement, GeoJsonFeature>("path")
                        .attr("opacity", 1)
                        .style("stroke-width", (borderSettings.showRegionBorder.value ? borderSettings.regionBorderThickness.value : 0) + "px");
                }
            })
            .catch((error) => {
                this.logDebugMessage(`Selection error: ${error}`);
            });

        return true;
    }

    /**
     * Cross-filter on the data rows behind a point or cluster, so a table visual can
     * list exactly the records that mark represents.
     */
    private selectPointIdentities(event: MouseEvent, ids: powerbi.visuals.ISelectionId[]) {
        const valid = (ids || []).filter(id => !!id);
        if (!valid.length) {
            return;
        }

        event.stopPropagation();
        const isCtrlPressed = event.ctrlKey || event.metaKey;
        this.selectionManager.select(valid, isCtrlPressed)
            .then((selected: powerbi.visuals.ISelectionId[]) => {
                this.setSelectedPointKeys(selected);
            })
            .catch((error) => {
                this.logDebugMessage(`Point selection error: ${error}`);
            });
    }

    /**
     * Remember which rows are selected so unselected marks can be dimmed.
     */
    private setSelectedPointKeys(ids: powerbi.visuals.ISelectionId[]) {
        if (!ids || ids.length === 0) {
            this.selectedPointKeys = null;
        } else {
            this.selectedPointKeys = new Set(ids.map(id => id.getKey()));
        }
        this.applyPointSelectionStyles();
    }

    private isPointSelected(point: {selectionId?: powerbi.visuals.ISelectionId}): boolean {
        if (!this.selectedPointKeys) {
            return true;
        }
        return !!point.selectionId && this.selectedPointKeys.has(point.selectionId.getKey());
    }

    /**
     * Dim marks that are not part of the current row selection.
     */
    private applyPointSelectionStyles() {
        if (this.settings.pointSettings.clusterSettingsGroup.enableClustering.value
            && this.clusterProjectedPoints && this.clusterProjectedPoints.length) {
            this.drawZoomClusters(d3.zoomTransform(this.svg.node()), true);
        } else {
            this.applyPointHighlightStyles();
        }
    }

    private isLegendHighlightActive(): boolean {
        return this.selectedLegendCategories.size > 0;
    }

    private legendMatches(legend?: string): boolean {
        if (!this.isLegendHighlightActive()) {
            return true;
        }
        return !!legend && this.selectedLegendCategories.has(legend);
    }

    /**
     * Plain click selects a single category (clicking it again clears);
     * ctrl/shift-click adds or removes categories from the current set.
     */
    private toggleLegendCategory(category: string, additive: boolean) {
        if (additive) {
            if (this.selectedLegendCategories.has(category)) {
                this.selectedLegendCategories.delete(category);
            } else {
                this.selectedLegendCategories.add(category);
            }
        } else {
            const alreadyOnlySelection = this.selectedLegendCategories.size === 1
                && this.selectedLegendCategories.has(category);
            this.selectedLegendCategories.clear();
            if (!alreadyOnlySelection) {
                this.selectedLegendCategories.add(category);
            }
        }

        this.applyLegendHighlight();
    }

    private clearLegendHighlight() {
        if (!this.isLegendHighlightActive()) {
            return;
        }
        this.selectedLegendCategories.clear();
        this.applyLegendHighlight();
    }

    /** Drop selections for categories that no longer exist after a data refresh. */
    private pruneLegendSelection() {
        if (!this.isLegendHighlightActive()) {
            return;
        }
        Array.from(this.selectedLegendCategories).forEach(category => {
            if (this.legendCategories.indexOf(category) < 0) {
                this.selectedLegendCategories.delete(category);
            }
        });
    }

    private applyLegendHighlight() {
        this.updatePointLegendSelectionStyles();

        if (this.settings.pointSettings.clusterSettingsGroup.enableClustering.value
            && this.clusterProjectedPoints && this.clusterProjectedPoints.length) {
            this.drawZoomClusters(d3.zoomTransform(this.svg.node()), true);
        } else {
            this.applyPointHighlightStyles();
        }
    }

    /** Dim point circles outside the selected legend categories or the row selection. */
    private applyPointHighlightStyles() {
        const dim = this.legendDimOpacity;
        const inFocus = (d: ProjectedPointDatum) =>
            !d || (this.legendMatches(d.legend) && this.isPointSelected(d));

        this.pointsGroup.selectAll<SVGCircleElement, ProjectedPointDatum>("circle")
            .attr("opacity", function (d) {
                const base = Number(d3.select(this).attr("data-base-opacity"));
                const baseOpacity = isNaN(base) ? 1 : base;
                return inFocus(d) ? baseOpacity : baseOpacity * dim;
            });
    }

    /** Reflect the current selection in the legend swatches themselves. */
    private updatePointLegendSelectionStyles() {
        if (!this.pointLegendContainer) {
            return;
        }

        const active = this.isLegendHighlightActive();
        d3Select(this.pointLegendContainer).selectAll<HTMLDivElement, unknown>("div.point-legend-item")
            .each((_, i, nodes) => {
                const item = d3Select(nodes[i]);
                const category = item.attr("data-category");
                const selected = this.selectedLegendCategories.has(category);
                item.style("opacity", !active || selected ? "1" : "0.4")
                    .style("font-weight", selected ? "bold" : "normal");
            });
    }

    private updatePointLegendPositioning() {
        if (!this.pointLegendContainer) return;

        const position = this.settings.pointSettings.pointSettingsGroup.pointLegendPosition.value.value as string;

        // Update positioning with fixed styling
        d3Select(this.pointLegendContainer)
            .style("background-color", "rgba(255, 255, 255, 0.8)")
            .style("border", "0.2px solid #000000")
            .style("top", position.includes("top") ? "10px" : "auto")
            .style("bottom", position.includes("bottom") ? "10px" : "auto")
            .style("left", position.includes("left") ? "10px" : "auto")
            .style("right", position.includes("right") ? "10px" : "auto");
    }

    /**
     * Update the showRegionName setting based on whether the checkbox is checked.
     */
    private toggleRegionNames(isChecked: boolean) {
        // Don't proceed if region names toggle is disabled in cockpit
        const regionNamesToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showRegionNamesToggle.value;
        const labelsToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showLabelsToggle.value;
        const masterLabelsEnabled = this.settings.labelSettings.labelSettingsGroup.showLabels.value;
        
        if (!regionNamesToggleEnabled) {
            return;
        }
        
        // If master labels are OFF but cockpit allows toggling, check visual labels state
        if (!masterLabelsEnabled && labelsToggleEnabled && !this.localShowLabels) {
            return;
        }
        
        // Update local state only (don't modify settings to avoid formatting pane refresh)
        this.localShowRegionName = isChecked;
        
        // Update labels display directly
        this.updateLabelsVisibility();
    }

    private toggleValuesVisibility(isChecked: boolean) {
        // Don't proceed if values toggle is disabled in cockpit
        const valuesToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showValuesToggle.value;
        const labelsToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showLabelsToggle.value;
        const masterLabelsEnabled = this.settings.labelSettings.labelSettingsGroup.showLabels.value;
        
        if (!valuesToggleEnabled) {
            return;
        }
        
        // If master labels are OFF but cockpit allows toggling, check visual labels state
        if (!masterLabelsEnabled && labelsToggleEnabled && !this.localShowLabels) {
            return;
        }
        
        // Update local state only (don't modify settings to avoid formatting pane refresh)
        this.localShowValues = isChecked;
        
        // Update labels display directly
        this.updateLabelsVisibility();
    }

    private toggleLegendVisibility(isChecked: boolean) {
        // Don't proceed if legend toggle is disabled in cockpit
        const legendToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showLegendToggle.value;
        
        if (!legendToggleEnabled) {
            return;
        }
        
        // Cockpit toggle works as override - can show legend even if disabled in formatting pane
        if (isChecked) {
            // Force show legend regardless of formatting pane setting
            this.createLegend(true);
        } else {
            // Hide legend (respects formatting pane setting)
            this.createLegend(false);
        }
    }

    private togglePointLegendVisibility(isChecked: boolean) {
        // Don't proceed if point legend toggle is disabled in cockpit
        const pointLegendToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showPointLegendToggle.value;
        
        if (!pointLegendToggleEnabled) {
            return;
        }
        
        // Cockpit toggle works as override - can show point legend even if disabled in formatting pane
        if (isChecked) {
            // Force show point legend regardless of formatting pane setting
            this.createPointLegend(true);
        } else {
            // Hide point legend (respects formatting pane setting)
            this.createPointLegend(false);
        }
    }

    private toggleLabelsVisibility(isChecked: boolean) {
        // Don't proceed if labels toggle is disabled in cockpit
        const labelsToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showLabelsToggle.value;
        
        if (!labelsToggleEnabled) {
            return;
        }
        
        // Update local state only
        this.localShowLabels = isChecked;
        
        // Update labels display directly
        this.updateLabelsVisibility();
        
        // Update toggle states directly
        this.updateRegionNamesToggleState();
        this.updateValuesToggleState();
    }

    private updateRegionNamesToggleState() {
        // Disable/enable region names toggle based on cockpit settings and master settings
        if (this.checkbox) {
            const regionNamesToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showRegionNamesToggle.value;
            const labelsToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showLabelsToggle.value;
            const masterLabelsEnabled = this.settings.labelSettings.labelSettingsGroup.showLabels.value;
            
            // If master settings have labels ON, toggles depend on visual labels state
            // If master settings have labels OFF but cockpit allows toggling, toggles depend on visual labels state
            let shouldDisable = !regionNamesToggleEnabled;
            
            if (masterLabelsEnabled || (!masterLabelsEnabled && labelsToggleEnabled)) {
                // Labels are controlled by visual state - depend on visual labels state
                shouldDisable = shouldDisable || !this.localShowLabels;
            }
            
            this.checkbox.disabled = shouldDisable;
            
            // Style the checkbox and label to show disabled state
            const regionNamesDiv = d3Select(this.uiContent).selectAll("div").nodes()[1] as HTMLDivElement;
            if (regionNamesDiv) {
                const label = d3Select(regionNamesDiv).select("label");
                const checkbox = d3Select(regionNamesDiv).select("input");
                
                if (!shouldDisable) {
                    label.style("color", "black").style("opacity", "1");
                    checkbox.style("opacity", "1").style("cursor", "pointer");
                } else {
                    label.style("color", "gray").style("opacity", "0.5");
                    checkbox.style("opacity", "0.5").style("cursor", "not-allowed");
                }
            }
        }
    }

    private updateValuesToggleState() {
        // Disable/enable values toggle based on cockpit settings and master settings
        if (this.valuesCheckbox) {
            const valuesToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showValuesToggle.value;
            const labelsToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showLabelsToggle.value;
            const masterLabelsEnabled = this.settings.labelSettings.labelSettingsGroup.showLabels.value;
            
            // If master settings have labels ON, toggles depend on visual labels state
            // If master settings have labels OFF but cockpit allows toggling, toggles depend on visual labels state
            let shouldDisable = !valuesToggleEnabled;
            
            if (masterLabelsEnabled || (!masterLabelsEnabled && labelsToggleEnabled)) {
                // Labels are controlled by visual state - depend on visual labels state
                shouldDisable = shouldDisable || !this.localShowLabels;
            }
            
            this.valuesCheckbox.disabled = shouldDisable;
            
            // Style the checkbox and label to show disabled state
            const valuesDiv = d3Select(this.uiContent).selectAll("div").nodes()[2] as HTMLDivElement;
            if (valuesDiv) {
                const label = d3Select(valuesDiv).select("label");
                const checkbox = d3Select(valuesDiv).select("input");
                
                if (!shouldDisable) {
                    label.style("color", "black").style("opacity", "1");
                    checkbox.style("opacity", "1").style("cursor", "pointer");
                } else {
                    label.style("color", "gray").style("opacity", "0.5");
                    checkbox.style("opacity", "0.5").style("cursor", "not-allowed");
                }
            }
        }
    }

    private updateLabelsVisibility() {
        // Simple wrapper that calls the existing updateLabelsOnly method
        this.updateLabelsOnly();
    }

    private updateLabelsOnly() {
        // Update only the labels without full re-render
        const labels = this.labelsGroup.selectAll<SVGGElement, GeoJsonFeature>("g.label");
        
        // Use requestAnimationFrame to batch DOM updates and improve performance
        requestAnimationFrame(() => {
            labels.each((d: GeoJsonFeature, i, nodes) => {
                const group = d3.select(nodes[i]);
                const text = group.select("text");
                const rect = group.select("rect");
                
                // Re-calculate label text based on current settings
                let labelText: string;
                if (this.settings.labelSettings.labelSettingsGroup.useCustomLabel.value.valueOf()) {
                    labelText = d.customLabel && d.customLabel.length > 0
                        ? d.customLabel
                        : "---";
                } else {
                    // Prepare label text
                    const labelParts: string[] = [];
                    if (this.localShowRegionName && d.properties.rubriek !== "custom") {
                        labelParts.push(d.properties.naam);
                    }
                    if (this.settings.mapSettings.mapSelectSettingsGroup.showExtraFeatures.value.valueOf() && d.properties.rubriek === "custom") {
                        labelParts.push(d.properties.naam);
                    }
                    // Add value with proper formatting
                    if (this.localShowValues && this.measureValues && this.measureValues.length > 0) {
                        if (this.isUsingPointCounts) {
                            // For point counts, show as integer
                            labelParts.push(d.value.toString());
                        } else {
                            // For regular measure values, use formatting
                            const measureValue = this.dataViews && this.dataViews[0] && this.dataViews[0].categorical && this.dataViews[0].categorical.values
                                ? this.dataViews[0].categorical.values[0] : undefined;
                            const formatString = (measureValue && measureValue.source && measureValue.source.format) || "General";
                            const formattedValue = formatString === "General" 
                                ? d.value.toFixed(this.settings.labelSettings.labelSettingsGroup.numOfDigits.value)
                                : valueFormatter.format(d.value, formatString);
                            labelParts.push(formattedValue);
                        }
                    }

                    // Join parts
                    if (this.settings.mapSettings.mapSelectSettingsGroup.showExtraFeatures.value.valueOf() && d.properties.rubriek === "custom") {
                        labelText = labelParts.join(": ");
                    } else {
                        labelText = labelParts.join("\n");
                    }
                }
                
                // Check if there's actually any content to display
                const hasContent = labelText && labelText.trim().length > 0;
                
                // Update text content
                text.selectAll("tspan").remove();
                if (hasContent) {
                    labelText.split("\n").forEach((line, index) => {
                        text.append("tspan")
                            .attr("x", text.attr("x"))
                            .attr("dy", index === 0 ? "0" : "1.2em")
                            .text(line);
                    });
                }
                
                // Update visibility based on local showLabels state AND whether there's content
                const shouldShow = this.localShowLabels && hasContent;
                text.style("display", shouldShow ? null : "none");
                rect.style("display", shouldShow ? null : "none");
                
                // Update background rectangle if visible - only call getBBox when necessary
                if (shouldShow) {
                    const textNode = text.node() as SVGTextElement;
                    if (textNode) {
                        // Use a more efficient approach - only call getBBox once per visible label
                        try {
                            const textBBox = textNode.getBBox();
                            // Get current zoom transform to scale properties correctly
                            const currentTransform = d3.zoomTransform(this.svg.node());
                            const padding = 3 / currentTransform.k;
                            
                            rect.attr("x", textBBox.x - padding)
                                .attr("y", textBBox.y - padding)
                                .attr("rx", 5 / currentTransform.k)
                                .attr("ry", 5 / currentTransform.k)
                                .attr("width", textBBox.width + 2 * padding)
                                .attr("height", textBBox.height + 2 * padding)
                                .attr("fill", this.settings.labelSettings.labelSettingsGroup.showLabelBackground.value.valueOf() ? this.settings.labelSettings.labelSettingsGroup.labelBackgroundColor.value.value : "none")
                                .attr("opacity", this.settings.labelSettings.labelSettingsGroup.labelBackgroundOpacity.value / 100)
                                .style("stroke", this.settings.labelSettings.labelSettingsGroup.showLabelBorder.value.valueOf() ? (this.settings.labelSettings.labelSettingsGroup as any).labelBorderColor?.value?.value : "none")
                                .style("stroke-width", (this.settings.labelSettings.labelSettingsGroup.showLabelBorder.value.valueOf() ? ((this.settings.labelSettings.labelSettingsGroup as any).labelBorderThickness?.value ?? 1) / currentTransform.k : 0) + "px");
                        } catch (error) {
                            // If getBBox fails, skip background rectangle update
                            console.warn("Failed to get text bounding box for label:", error);
                        }
                    }
                }
            });
        });
    }
    
    private toggleCockpitVisibility(isChecked: boolean) {
        d3Select(this.uiContainer).style("display", isChecked ? "block" : "none");
    }

    private toggleCockpitCollapse() {
        this.localCockpitCollapsed = !this.localCockpitCollapsed;
        this.updateCockpitCollapseState();
    }

    private updateCockpitCollapseState() {
        const content = d3Select(this.uiContent);
        const header = d3Select(this.uiHeader);
        const collapseButton = header.select("span:last-child");
        
        if (this.localCockpitCollapsed) {
            // Collapse: hide content and rotate arrow
            content.style("max-height", "0")
                  .style("opacity", "0")
                  .style("padding", "0 10px");
            collapseButton.style("transform", "rotate(-90deg)");
        } else {
            // Expand: show content and rotate arrow back
            content.style("max-height", "200px")
                  .style("opacity", "1")
                  .style("padding", "10px");
            collapseButton.style("transform", "rotate(0deg)");
        }
    }

    private initializeCockpitState() {
        // Initialize cockpit collapse state based on setting
        this.localCockpitCollapsed = this.settings.cockpitSettings.cockpitSettingsGroup.startCockpitCollapsed.value;
        this.updateCockpitCollapseState();
    }

    private updateCockpitOptionsVisibility() {
        // Check if cockpit is enabled at all
        const cockpitEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.enableCockpit.value;
        if (!cockpitEnabled) {
            this.toggleCockpitVisibility(false);
            return;
        }

        // Show/hide labels toggle based on setting (first div in content)
        const labelsDiv = d3Select(this.uiContent).selectAll("div").nodes()[0] as HTMLDivElement;
        const showLabelsToggle = this.settings.cockpitSettings.cockpitSettingsGroup.showLabelsToggle.value;
        if (labelsDiv) {
            d3Select(labelsDiv).style("display", showLabelsToggle ? "block" : "none");
        }

        // Show/hide region names toggle based on setting (second div in content)
        const regionNamesDiv = d3Select(this.uiContent).selectAll("div").nodes()[1] as HTMLDivElement;
        const showRegionNamesToggle = this.settings.cockpitSettings.cockpitSettingsGroup.showRegionNamesToggle.value;
        const masterLabelsEnabled = this.settings.labelSettings.labelSettingsGroup.showLabels.value;
        if (regionNamesDiv) {
            // Show region names toggle if its setting is enabled AND labels are not completely disabled
            const shouldShowRegionNames = showRegionNamesToggle && (masterLabelsEnabled || showLabelsToggle);
            d3Select(regionNamesDiv).style("display", shouldShowRegionNames ? "block" : "none");
        }

        // Show/hide values toggle based on setting (third div in content)
        const valuesDiv = d3Select(this.uiContent).selectAll("div").nodes()[2] as HTMLDivElement;
        const showValuesToggle = this.settings.cockpitSettings.cockpitSettingsGroup.showValuesToggle.value;
        if (valuesDiv) {
            // Show values toggle if its setting is enabled AND labels are not completely disabled
            const shouldShowValues = showValuesToggle && (masterLabelsEnabled || showLabelsToggle);
            d3Select(valuesDiv).style("display", shouldShowValues ? "block" : "none");
        }

        // Show/hide legend toggle based on setting (fourth div in content)
        const legendDiv = d3Select(this.uiContent).selectAll("div").nodes()[3] as HTMLDivElement;
        const showLegendToggle = this.settings.cockpitSettings.cockpitSettingsGroup.showLegendToggle.value;
        if (legendDiv) {
            d3Select(legendDiv).style("display", showLegendToggle ? "block" : "none");
        }

        // Show/hide point legend toggle based on setting (fifth div in content)
        const pointLegendDiv = d3Select(this.uiContent).selectAll("div").nodes()[4] as HTMLDivElement;
        const showPointLegendToggle = this.settings.cockpitSettings.cockpitSettingsGroup.showPointLegendToggle.value;
        if (pointLegendDiv) {
            d3Select(pointLegendDiv).style("display", showPointLegendToggle ? "block" : "none");
        }

        // Hide entire cockpit if no toggles are actually visible
        const anyToggleVisible = showLabelsToggle || 
                                 (showRegionNamesToggle && (masterLabelsEnabled || showLabelsToggle)) || 
                                 (showValuesToggle && (masterLabelsEnabled || showLabelsToggle)) || 
                                 showLegendToggle ||
                                 showPointLegendToggle;
        this.toggleCockpitVisibility(anyToggleVisible);
    }

    private syncLocalStateWithSettings() {
        // Always sync local state with main settings (main settings are the source of truth)
        // The cockpit toggles only control what the end-user can see/access
        this.localShowLabels = this.settings.labelSettings.labelSettingsGroup.showLabels.value;
        this.localShowRegionName = this.settings.labelSettings.labelSettingsGroup.showRegionName.value;
        this.localShowValues = this.settings.labelSettings.labelSettingsGroup.showValues.value;
        this.localShowLegend = this.settings.legendSettings.legendSettingsGroup.showLegend.value;
        this.localShowPointLegend = this.settings.pointSettings.pointSettingsGroup.showPointLegend.value;
    }

    private syncCheckboxStates() {
        // Sync region names checkbox with local state
        if (this.checkbox) {
            this.checkbox.checked = this.localShowRegionName;
        }

        // Sync values checkbox with local state
        if (this.valuesCheckbox) {
            this.valuesCheckbox.checked = this.localShowValues;
        }

        // Sync legend checkbox with local state
        if (this.legendCheckbox) {
            this.legendCheckbox.checked = this.localShowLegend;
        }

        // Sync point legend checkbox with local state
        if (this.pointLegendCheckbox) {
            this.pointLegendCheckbox.checked = this.localShowPointLegend;
        }

        // Sync labels checkbox with local state
        if (this.labelsCheckbox) {
            this.labelsCheckbox.checked = this.localShowLabels;
        }
    }

    private logDebugMessage(message: string | (() => string)) {
        // Only log in development mode to improve performance
        // Check if we're in development mode by looking for debug console
        if (this.debugConsole && this.debugConsole.style.display !== 'none') {
            const messageElement = document.createElement("div");
            messageElement.textContent = typeof message === "function" ? message() : message;
            this.debugConsole.appendChild(messageElement);

            // Scroll to the bottom to show the latest message
            this.debugConsole.scrollTop = this.debugConsole.scrollHeight;
        }
    }

    /**
     * Build a spatial index for efficient point-to-region mapping
     */
    private buildSpatialIndex(geojson: any): Map<string, any[]> {
        const spatialIndex = new Map<string, any[]>();
        const gridSize = this.spatialIndexGridSize;

        geojson.features.forEach((feature: any) => {
            try {
                const bounds = d3.geoBounds(feature);
                // Cache the bbox on the entry so point tests can reject cheaply
                // before paying for a full geoContains on detailed geometry.
                const entry = {
                    feature,
                    minLon: bounds[0][0],
                    minLat: bounds[0][1],
                    maxLon: bounds[1][0],
                    maxLat: bounds[1][1]
                };

                const minCellLon = Math.floor(entry.minLon / gridSize);
                const minCellLat = Math.floor(entry.minLat / gridSize);
                const maxCellLon = Math.floor(entry.maxLon / gridSize);
                const maxCellLat = Math.floor(entry.maxLat / gridSize);

                // Add feature to all grid cells it overlaps
                for (let lon = minCellLon; lon <= maxCellLon; lon++) {
                    for (let lat = minCellLat; lat <= maxCellLat; lat++) {
                        const key = `${lon},${lat}`;
                        const cell = spatialIndex.get(key);
                        if (cell) {
                            cell.push(entry);
                        } else {
                            spatialIndex.set(key, [entry]);
                        }
                    }
                }
            } catch (error) {
                // If bounds calculation fails, add to a fallback bucket
                const fallbackKey = "fallback";
                const fallback = spatialIndex.get(fallbackKey);
                const entry = { feature, minLon: -Infinity, minLat: -Infinity, maxLon: Infinity, maxLat: Infinity };
                if (fallback) {
                    fallback.push(entry);
                } else {
                    spatialIndex.set(fallbackKey, [entry]);
                }
            }
        });
        
        return spatialIndex;
    }

    /**
     * Get grid cell key for a point coordinate
     */
    private getGridKey(lon: number, lat: number): string {
        const gridSize = this.spatialIndexGridSize;
        return `${Math.floor(lon / gridSize)},${Math.floor(lat / gridSize)}`;
    }

    /**
     * First region whose geometry contains the point, or null.
     * Walks the grid cell then the fallback bucket, rejecting on bbox before geoContains.
     */
    private findContainingFeature(lon: number, lat: number): any | null {
        if (!this.spatialIndex) {
            return null;
        }

        const coords: [number, number] = [lon, lat];
        const buckets = [this.spatialIndex.get(this.getGridKey(lon, lat)), this.spatialIndex.get("fallback")];

        for (const bucket of buckets) {
            if (!bucket) {
                continue;
            }
            for (const entry of bucket) {
                if (lon < entry.minLon || lon > entry.maxLon || lat < entry.minLat || lat > entry.maxLat) {
                    continue;
                }
                try {
                    if (geoContains(entry.feature as any, coords)) {
                        return entry.feature;
                    }
                } catch {
                    continue;
                }
            }
        }
        return null;
    }

    /**
     * Optimized point counting using spatial index
     */
    /**
     * Rebuild the spatial index only when the region set actually changed.
     */
    private ensureSpatialIndex(geojson: any) {
        const features = geojson?.features || [];
        let key = String(features.length);
        for (const feature of features) {
            key += "|" + (feature.properties?.code || feature.properties?.naam);
        }

        if (!this.spatialIndex || this.spatialIndexKey !== key) {
            this.spatialIndex = this.buildSpatialIndex(geojson);
            this.spatialIndexKey = key;
        }
    }

    /**
     * Single point-in-region pass producing both the points that landed inside a
     * region and the per-region tally. Walking the geometry twice (once to filter,
     * once to count) was the most expensive part of applying a filter.
     */
    private assignPointsToRegions(
        pointData: PointDatum[],
        geojson: any
    ): { inside: PointDatum[], counts: Map<string, number> } {
        const counts = new Map<string, number>();
        const inside: PointDatum[] = [];

        if (!pointData.length || !geojson?.features?.length) {
            return { inside, counts };
        }

        this.ensureSpatialIndex(geojson);

        for (const point of pointData) {
            const feature = this.findContainingFeature(point.lon, point.lat);
            if (feature) {
                inside.push(point);
                const regionKey = feature.properties.code || feature.properties.naam;
                counts.set(regionKey, (counts.get(regionKey) || 0) + 1);
            }
        }

        return { inside, counts };
    }

    /**
     * Pre-compute projections for all points to avoid redundant calculations
     */
    private precomputeProjections(pointData: PointDatum[], projection: any): ProjectedPointDatum[] {
        return pointData.map(point => {
            const cacheKey = `${point.lon},${point.lat}`;
            
            // Check cache first
            if (this.projectedPointsCache.has(cacheKey)) {
                const cached = this.projectedPointsCache.get(cacheKey)!;
                return {
                    ...point,
                    projectedX: cached.projectedX,
                    projectedY: cached.projectedY
                };
            }
            
            // Compute projection
            const coords = projection([point.lon, point.lat]);
            const projectedX = coords ? coords[0] : 0;
            const projectedY = coords ? coords[1] : 0;
            
            // Cache the result
            this.projectedPointsCache.set(cacheKey, { projectedX, projectedY });
            
            return {
                ...point,
                projectedX,
                projectedY
            };
        });
    }

    /**
     * Limit point data for very large datasets to maintain performance
     */
    private limitPointData(pointData: PointDatum[]): PointDatum[] {
        if (pointData.length <= this.maxPointsLimit) {
            return pointData;
        }
        
        // Sample points evenly across the dataset to maintain distribution
        const step = Math.ceil(pointData.length / this.maxPointsLimit);
        const limitedData = pointData.filter((_, index) => index % step === 0);
        
        // Log warning about data limitation
        console.warn(`Point data limited from ${pointData.length} to ${limitedData.length} points for performance`);
        
        return limitedData;
    }

    /**
     * Blank cells must not become 0 — rows without coordinates are not points.
     */
    private toCoordinate(value: any): number {
        if (value === null || value === undefined || value === "") {
            return NaN;
        }
        return Number(value);
    }

    /**
     * Called O(regions x categoryValues) times per update — memoised because the
     * regex passes otherwise dominate matching on the larger maps.
     */
    private normalizeRegionKey(value: any): string {
        const raw = String(value ?? "");
        const cached = this.normalizedKeyCache.get(raw);
        if (cached !== undefined) {
            return cached;
        }

        const normalized = raw
            .toLowerCase()
            .trim()
            .replace(/^ggd\s+/i, "")
            .replace(/\s+/g, " ");
        this.normalizedKeyCache.set(raw, normalized);
        return normalized;
    }

    private regionMatchesFilterValue(naam: any, code: any, filterValue: string): boolean {
        const filter = this.normalizeRegionKey(filterValue);
        if (!filter) {
            return false;
        }

        const name = this.normalizeRegionKey(naam);
        const regionCode = this.normalizeRegionKey(code);

        return (
            name === filter ||
            this.regionCodesMatch(regionCode, filter) ||
            (name.length > 0 && (name.includes(filter) || filter.includes(name)))
        );
    }

    /**
     * PO1 and PO01 are the same code. PO1 must not match PO10.
     */
    private regionCodesMatch(left: string, right: string): boolean {
        if (!left || !right) {
            return false;
        }
        if (left === right) {
            return true;
        }
        const pattern = /^([a-z]+)0*(\d+)$/;
        const a = left.match(pattern);
        const b = right.match(pattern);
        return !!a && !!b && a[1] === b[1] && a[2] === b[2];
    }

    /**
     * Collapses PO1 and PO01 onto one key (but keeps PO10 distinct) so codes can
     * be compared with a Map lookup instead of a regex pair per comparison.
     */
    private canonicalRegionCode(value: string): string {
        const match = value.match(/^([a-z]+)0*(\d+)$/);
        return match ? match[1] + match[2] : value;
    }

    /**
     * Index a list of region values so each feature can be matched with a hash lookup.
     * Scanning every value per feature is O(regions x values), which on the gemeenten
     * maps is >100k comparisons and the single biggest cost of applying a filter.
     */
    private buildRegionLookup(values: any[]): RegionLookup {
        const byName = new Map<string, number>();
        const byCode = new Map<string, number>();
        const normalized: string[] = [];

        values.forEach((value, index) => {
            const key = this.normalizeRegionKey(value);
            normalized.push(key);
            if (!key) {
                return;
            }
            if (!byName.has(key)) {
                byName.set(key, index);
            }
            const code = this.canonicalRegionCode(key);
            if (!byCode.has(code)) {
                byCode.set(code, index);
            }
        });

        return { byName, byCode, normalized };
    }

    /**
     * Index of the first value matching this region, or -1. Exact name and code hits
     * are preferred over the fuzzy substring fallback, which is only scanned when
     * both lookups miss.
     */
    private findRegionIndex(lookup: RegionLookup, naam: any, code: any): number {
        const name = this.normalizeRegionKey(naam);
        if (name) {
            const byName = lookup.byName.get(name);
            if (byName !== undefined) {
                return byName;
            }
        }

        const regionCode = this.normalizeRegionKey(code);
        if (regionCode) {
            const byCode = lookup.byCode.get(this.canonicalRegionCode(regionCode));
            if (byCode !== undefined) {
                return byCode;
            }
        }

        if (name) {
            for (let i = 0; i < lookup.normalized.length; i++) {
                const candidate = lookup.normalized[i];
                if (candidate && (name.includes(candidate) || candidate.includes(name))) {
                    return i;
                }
            }
        }

        return -1;
    }

    /**
     * Resolve which map regions should scope points/clusters
     * (Filter well, external highlight/selection, or hard-filtered Regio data).
     */
    private getPointScopeGeojson(
        geojson: any,
        hasHighlights: boolean,
        highlights: any[] | undefined,
        categoryValues: any[]
    ): { geojson: any, restrict: boolean } {
        const allFeatures = geojson?.features || [];
        if (!allFeatures.length) {
            return { geojson, restrict: false };
        }

        // Filter well already trimmed geojson.features
        if (this.filterValues.length > 0) {
            return { geojson, restrict: true };
        }

        // External/cross-highlight: only highlighted Regio rows
        if (hasHighlights && highlights && categoryValues.length > 0) {
            const active = new Set<string>();
            categoryValues.forEach((value, index) => {
                if (highlights[index] !== undefined && highlights[index] !== null) {
                    active.add(String(value).toLowerCase().trim());
                }
            });
            if (active.size > 0) {
                const activeLookup = this.buildRegionLookup(Array.from(active));
                const features = allFeatures.filter((feature: any) =>
                    this.findRegionIndex(activeLookup, feature.properties?.naam, feature.properties?.code) !== -1
                );
                if (features.length > 0) {
                    return { geojson: { type: geojson.type, features }, restrict: true };
                }
            }
        }

        // Hard filter / slicer: Regio category list is a subset of the map
        if (categoryValues.length > 0 && categoryValues.length < allFeatures.length) {
            const categoryLookup = this.buildRegionLookup(categoryValues);
            const features = allFeatures.filter((feature: any) =>
                this.findRegionIndex(categoryLookup, feature.properties?.naam, feature.properties?.code) !== -1
            );
            if (features.length > 0 && features.length < allFeatures.length) {
                return { geojson: { type: geojson.type, features }, restrict: true };
            }
        }

        return { geojson, restrict: false };
    }

    /**
     * Format a custom label measure value using Power BI format or numOfDigits.
     */
    private formatCustomLabelValue(raw: any, formatString: string | undefined): string {
        if (raw === null || raw === undefined || raw === "") {
            return "---";
        }

        const format = formatString || "General";
        const numeric = typeof raw === "number" ? raw : Number(raw);
        const isNumeric = typeof raw === "number" || (typeof raw === "string" && raw.trim() !== "" && !isNaN(numeric));

        if (isNumeric && !isNaN(numeric)) {
            if (format === "General" || format === "") {
                const digits = this.settings?.labelSettings?.labelSettingsGroup?.numOfDigits?.value ?? 1;
                return numeric.toFixed(digits);
            }
            return valueFormatter.format(numeric, format);
        }

        return String(raw);
    }

    /**
     * Detect map type from statcode prefix in category values
     */
    private detectMapTypeFromStatcode(categoryValues: any[]): string | null {
        if (!categoryValues || categoryValues.length === 0) {
            return null;
        }

        // Sample a few values to detect the pattern
        const sampleSize = Math.min(10, categoryValues.length);
        const samples = categoryValues.slice(0, sampleSize).map(v => String(v).toUpperCase().trim());

        // Count occurrences of each pattern to determine the most common type
        const patternCounts: { [key: string]: number } = {
            "provincies": 0,
            "gemeenten": 0,
            "wijken": 0,
            "ggdregios": 0,
            "veiligheidsregios": 0,
            "politieregios": 0
        };

        // Check for statcode prefixes and patterns
        for (const sample of samples) {
            if (sample.startsWith("PV")) {
                patternCounts["provincies"]++;
            } else if (sample.startsWith("GM")) {
                patternCounts["gemeenten"]++;
            } else if (sample.startsWith("WK") || sample.includes("WIJK")) {
                patternCounts["wijken"]++;
            } else if (sample.startsWith("GG")) {
                patternCounts["ggdregios"]++;
            } else if (sample.startsWith("VR") || sample.includes("VEILIGHEIDSREGIO")) {
                patternCounts["veiligheidsregios"]++;
            } else if (/^PO\d/.test(sample) || sample.includes("POLITIEREGIO")) {
                patternCounts["politieregios"]++;
            }
        }

        // Find the most common pattern
        let maxCount = 0;
        let detectedType: string | null = null;
        for (const [type, count] of Object.entries(patternCounts)) {
            if (count > maxCount) {
                maxCount = count;
                detectedType = type;
            }
        }

        // If gemeenten is detected, default to latest version
        if (detectedType === "gemeenten") {
            return "gemeenten 2023";
        }

        return detectedType;
    }

    /**
     * Get the appropriate region type label based on the selected map
     */
    private getRegionTypeLabel(selectedMap: any): string {
        const mapValue = selectedMap?.value || selectedMap;
        const regionTypeMap: { [key: string]: string } = {
            "provincies": "Provincie",
            "wijken": "Wijk", 
            "gemeenten 2023": "Gemeente",
            "gemeenten 2021": "Gemeente",
            "gemeenten 2020": "Gemeente",
            "gemeenten 2019": "Gemeente",
            "gemeenten 2016": "Gemeente",
            "gemeenten 2012": "Gemeente",
            "ggdregios": "GGD Regio",
            "veiligheidsregios": "Veiligheidsregio",
            "politieregios": "Politieregio"
        };
        
        return regionTypeMap[mapValue] || "Regio";
    }

    /**
     * Create tooltip data for a point
     */
    private createPointTooltipData(point: PointDatum, index: number, pointTooltipValues?: any[], legendFieldName?: string): VisualTooltipDataItem[] {
        const tooltipData: VisualTooltipDataItem[] = [];
        const dataIndex = point.sourceIndex !== undefined ? point.sourceIndex : index;
        
        // Add category information if enabled
        if (this.settings.tooltipSettings.tooltipSettingsGroup.showPointCategory.value) {
            tooltipData.push({
                displayName: "Punt",
                value: `#${dataIndex + 1}`
            });
        }
        
        // Add legend information if available and enabled
        if (this.settings.tooltipSettings.tooltipSettingsGroup.showPointLegend.value && point.legend) {
            // Use actual field name if available, otherwise fallback to "Categorie"
            const categoryFieldName = legendFieldName || "Categorie";
            tooltipData.push({
                displayName: categoryFieldName,
                value: point.legend
            });
        }
        
        // Add custom point tooltip fields if available
        if (pointTooltipValues && pointTooltipValues.length > 0) {
            pointTooltipValues.forEach(tooltipField => {
                if (tooltipField.values && tooltipField.values[dataIndex] !== undefined && tooltipField.values[dataIndex] !== null) {
                    tooltipData.push({
                        displayName: tooltipField.displayName,
                        value: tooltipField.values[dataIndex]
                    });
                }
            });
        }
        
        // Add coordinates if enabled
        if (this.settings.tooltipSettings.tooltipSettingsGroup.showPointCoordinates.value) {
            tooltipData.push({
                displayName: "Latitude",
                value: point.lat.toFixed(6)
            });
            tooltipData.push({
                displayName: "Longitude", 
                value: point.lon.toFixed(6)
            });
        }
        
        return tooltipData;
    }
    
    public update(options: VisualUpdateOptions) {
        // Store options for later use
        this.dataViews = options.dataViews;
        this.viewport = options.viewport;

        // Get settings (safe when no dataView)
        const dataView: DataView | undefined = options.dataViews && options.dataViews[0] ? options.dataViews[0] : undefined;
        this.dataView = dataView; // Store for use in formatting model
        this.settings = this.formattingSettingsService.populateFormattingSettingsModel(VisualSettings, dataView as any);

        this.updateCockpitOptionsVisibility();
        this.syncLocalStateWithSettings();
        this.initializeCockpitState();
        this.syncCheckboxStates();
        this.updateRegionNamesToggleState();
        this.updateValuesToggleState();

        // Always render, even without data
        const hasDataView = !!(options.dataViews && options.dataViews[0]);

        // Log the update call (only in development)
        this.logDebugMessage("Update called with options");

        // Continue rendering regardless of dataset presence

        // this.settings = Visual.parseSettings(options.dataViews[0]);

        const width = options.viewport.width;
        const height = options.viewport.height;
        const margin = 15; // Same margin as in constructor
        const categorical: any = dataView && dataView.categorical ? dataView.categorical : { categories: [], values: [] };

        this.updateZoomExtents(width, height);

        this.svg
            .attr("width", width)
            .attr("height", height)
            .attr("viewBox", `-${margin} -${margin} ${width + 2*margin} ${height + 2*margin}`)
            .attr("preserveAspectRatio", "xMinYMin meet");

        
        const newFieldSet = dataView && dataView.categorical && dataView.categorical.values
            ? dataView.categorical.values.map(value => value.source.displayName).join(",")
            : "";
        if (this.currentFieldSet !== newFieldSet) {
            this.currentFieldSet = newFieldSet; // Update the cached field set
            this.tooltipMap.clear(); // Clear precomputed data
        }

        // Get category values specifically from the 'Regio' field (category role), not from lat/lon
        const categoryColumn = dataView && dataView.categorical && dataView.categorical.categories
            ? dataView.categorical.categories.find(cat => cat.source.roles.category)
            : undefined;
        const categoryValues = categoryColumn ? categoryColumn.values : [];

        // Clicking a region cross-filters on the 'Regio' field. Point setups often leave
        // Regio empty and put the region key in 'Filter' instead, so fall back to that —
        // without it there is no column to build a selection id against and clicks do nothing.
        this.regionCategoryColumn = categoryColumn
            || (dataView && dataView.categorical && dataView.categorical.categories
                ? dataView.categorical.categories.find(cat => cat.source.roles.filter)
                : undefined);
        const measureValues = dataView && dataView.categorical && dataView.categorical.values
            ? (dataView.categorical.values.find(value => value.source.roles.measure)?.values.map(v => Number(v)) || [])
            : [];
        
        // Extract lat/lon data from categories (not aggregated)
        const latitudeValues = dataView && dataView.categorical && dataView.categorical.categories
            ? (dataView.categorical.categories.find(cat => cat.source.roles.latitude)?.values.map(v => this.toCoordinate(v)) || [])
            : [];
        const longitudeValues = dataView && dataView.categorical && dataView.categorical.categories
            ? (dataView.categorical.categories.find(cat => cat.source.roles.longitude)?.values.map(v => this.toCoordinate(v)) || [])
            : [];
        // Store legend category column for identity-based color binding
        this.legendCategoryColumn = dataView && dataView.categorical && dataView.categorical.categories
            ? dataView.categorical.categories.find(cat => cat.source.roles.pointCategory)
            : undefined;
        
        // Get the legend field name for tooltip display
        const legendFieldName = this.legendCategoryColumn?.source?.displayName;
        
        const pointCategoryValues = this.legendCategoryColumn
            ? this.legendCategoryColumn.values.map(v => (v === null || v === undefined) ? "" : String(v))
            : [];
        
        // Every grouping column shares one identity per data row, so any of them yields
        // the row identity needed to cross-filter a single point.
        const pointIdentityColumn = dataView && dataView.categorical && dataView.categorical.categories
            ? (dataView.categorical.categories.find(cat => cat.source.roles.latitude)
                || dataView.categorical.categories[0])
            : undefined;

        // Create point data array (only if both lat and lon are present)
        const pointData: PointDatum[] = [];
        if (latitudeValues.length > 0 && longitudeValues.length > 0) {
            // Include point category values in minLength calculation to ensure data alignment
            const lengths = [latitudeValues.length, longitudeValues.length];
            if (pointCategoryValues.length > 0) {
                lengths.push(pointCategoryValues.length);
            }
            const minLength = Math.min(...lengths);
            
            for (let i = 0; i < minLength; i++) {
                const lat = latitudeValues[i];
                const lon = longitudeValues[i];
                if (Number.isFinite(lat) && Number.isFinite(lon)) {
                    const point: PointDatum = {
                        lat,
                        lon,
                        sourceIndex: i
                    };
                    if (pointCategoryValues.length > i && pointCategoryValues[i]) {
                        point.legend = pointCategoryValues[i];
                    }
                    if (pointIdentityColumn) {
                        point.selectionId = this.host.createSelectionIdBuilder()
                            .withCategory(pointIdentityColumn, i)
                            .createSelectionId();
                    }
                    pointData.push(point);
                }
            }
        }

        // Extract unique legend categories for formatting pane
        if (pointCategoryValues.length > 0) {
            this.legendCategories = Array.from(new Set(pointCategoryValues.filter(v => v && v !== ""))).sort();
        } else {
            this.legendCategories = [];
        }
        this.resolveLegendSlots();
        this.pruneLegendSelection();
        
        // Process tooltip data during update
        // Tooltip fields with "kind": "Grouping" are in categorical.categories, not categorical.values
        const tooltipValues = dataView && dataView.categorical && dataView.categorical.categories
            ? dataView.categorical.categories
                .filter(category => category.source.roles.tooltip)  // Get fields marked as tooltip
                .map(category => ({
                    displayName: category.source.displayName,
                    values: category.values.map(v => v?.toString() ?? "N/A")  // Handle null values
                }))
            : [];

        // Process point tooltip data during update (from categories)
        const pointTooltipValues = dataView && dataView.categorical && dataView.categorical.categories
            ? dataView.categorical.categories
                .filter(category => category.source.roles.pointTooltip)  // Get fields marked as pointTooltip
                .map(category => ({
                    displayName: category.source.displayName,
                    values: category.values.map(v => v?.toString() ?? "N/A")  // Handle null values
                }))
            : [];

        this.measureValues = measureValues;
        
        // Identify the active measure column
        const measureColumn = (dataView && dataView.categorical && dataView.categorical.values)
            ? dataView.categorical.values.find((v: any) => v.source.roles.measure)
            : undefined;

        // A stable key to detect measure changes across updates
        const newMeasureKey = measureColumn?.source?.queryName || measureColumn?.source?.displayName || "";

        // If the measure changed, reset cached scale
        if (this.measureKey !== newMeasureKey) {
            this.measureKey = newMeasureKey;
            this.scaleMin = null;
            this.scaleMax = null;
            this.logDebugMessage("Measure changed, resetting color scale cache");
        }

        // Color scale caching will be done after point counting to ensure correct values

        const filterCategoryValues = dataView && dataView.categorical && dataView.categorical.categories
            ? (dataView.categorical.categories.find(category => category.source.roles.filter)?.values || [])
            : [];
        this.filterValues = Array.from(new Set(
            filterCategoryValues
                .map(value => (value == null ? "" : value.toString().trim()))
                .filter(value => value.length > 0)
        ));
        
        // Debug logging for filtering (only in development)
        if (this.filterValues.length > 0) {
            this.logDebugMessage(`Filter values: ${this.filterValues.join(", ")}`);
            this.logDebugMessage(`Using cached scale: min=${this.scaleMin}, max=${this.scaleMax}`);
        }

        // Auto-detect map type from statcode if category values are available
        if (categoryValues && categoryValues.length > 0) {
            const detectedMapType = this.detectMapTypeFromStatcode(categoryValues);
            if (detectedMapType) {
                const currentMapValue = this.settings.mapSettings.mapSelectSettingsGroup.selectedMapSlice.value.value;
                if (currentMapValue !== detectedMapType) {
                    // Find the matching map option
                    const mapOptions = [
                        { displayName: "Provincies", value: "provincies" },
                        { displayName: "Wijken", value: "wijken" },
                        { displayName: "Gemeenten 2023", value: "gemeenten 2023" },
                        { displayName: "Gemeenten 2021", value: "gemeenten 2021" },
                        { displayName: "Gemeenten 2020", value: "gemeenten 2020" },
                        { displayName: "Gemeenten 2019", value: "gemeenten 2019" },
                        { displayName: "Gemeenten 2016", value: "gemeenten 2016" },
                        { displayName: "Gemeenten 2012", value: "gemeenten 2012" },
                        { displayName: "GGD Regios", value: "ggdregios" },
                        { displayName: "Veiligheidsregios", value: "veiligheidsregios" },
                        { displayName: "Politieregios", value: "politieregios" }
                    ];
                    const matchingOption = mapOptions.find(opt => opt.value === detectedMapType);
                    if (matchingOption) {
                        this.settings.mapSettings.mapSelectSettingsGroup.selectedMapSlice.value = matchingOption;
                        this.logDebugMessage(`Auto-detected map type: ${detectedMapType} from statcode`);
                    }
                }
            }
        }

        let geojson;
        const mapKey = this.settings.mapSettings.mapSelectSettingsGroup.selectedMapSlice.value.value as string;
        // Converting TopoJSON to GeoJSON decodes every arc (up to 3 MB for gemeenten), so do it
        // once per map and hand out shallow per-update copies: update() mutates value /
        // tooltipData / selectionId on each feature, but never geometry or properties.
        const cachedBaseGeojson = this.baseGeojsonCache.get(mapKey);
        if (cachedBaseGeojson) {
            geojson = {
                type: cachedBaseGeojson.type,
                features: cachedBaseGeojson.features.map((f: any) => ({ ...f }))
            };
        } else {
        // if (this.settings.mapSettings.mapSelectSettingsGroup.customTopojson.value) {
        //     //try {
        //         // Parse the TopoJSON string
        //         const topojsonData = topojson.parse(this.settings.mapSettings.mapSelectSettingsGroup.customTopojson.value.trim());
                
        //         // Convert TopoJSON to GeoJSON
        //         const geojson = feature(topojsonData, topojsonData.objects.gj);
        
        //         if (geojson?.features) {
        //             geojson.features.forEach((feature) => {
        //                 const originalProps = feature.properties;
        //                 feature.properties = {
        //                     naam: originalProps?.statnaam || "Unknown",
        //                     code: originalProps?.statcode || "N/A",
        //                     rubriek: originalProps?.rubriek || ""
        //                 };
        //             });
        //         }
        //     // } catch (error) {
        //     //     console.error("Error parsing custom TopoJSON:", error);
        //     // }
        // } else {
            switch (this.settings.mapSettings.mapSelectSettingsGroup.selectedMapSlice.value.value) {
                case "provincies":
                    geojson = feature(provinciesTopojson, provinciesTopojson.objects.gj);
                geojson.features.forEach((feature) => {
                    const originalProps = feature.properties;
                    feature.properties = {
                        naam: originalProps?.statnaam || "Unknown",
                        code: originalProps?.statcode || "N/A",
                        rubriek: originalProps?.rubriek || ""
                    };
                });
                break;
            case "gemeenten 2023":
                geojson = feature(gemeentenTopojson_2023, gemeentenTopojson_2023.objects.gj);
                geojson.features.forEach((feature) => {
                    const originalProps = feature.properties;
                    feature.properties = {
                        naam: originalProps?.statnaam || "Unknown",
                        code: originalProps?.statcode || "N/A",
                        rubriek: originalProps?.rubriek || ""
                    };
                });
                break;
            case "gemeenten 2021":
                geojson = feature(gemeentenTopojson_2021, gemeentenTopojson_2021.objects.gj);
                geojson.features.forEach((feature) => {
                    const originalProps = feature.properties;
                    feature.properties = {
                        naam: originalProps?.statnaam || "Unknown",
                        code: originalProps?.statcode || "N/A",
                        rubriek: originalProps?.rubriek || ""
                    };
                });
                break;
            case "gemeenten 2020":
                geojson = feature(gemeentenTopojson_2020, gemeentenTopojson_2020.objects.gj);
                geojson.features.forEach((feature) => {
                    const originalProps = feature.properties;
                    feature.properties = {
                        naam: originalProps?.statnaam || "Unknown",
                        code: originalProps?.statcode || "N/A",
                        rubriek: originalProps?.rubriek || ""
                    };
                });
                break;
            case "gemeenten 2019":
                geojson = feature(gemeentenTopojson_2019, gemeentenTopojson_2019.objects.gj);
                geojson.features.forEach((feature) => {
                    const originalProps = feature.properties;
                    feature.properties = {
                        naam: originalProps?.statnaam || "Unknown",
                        code: originalProps?.statcode || "N/A",
                        rubriek: originalProps?.rubriek || ""
                    };
                });
                break;
            case "gemeenten 2016":
                geojson = feature(gemeentenTopojson_2016, gemeentenTopojson_2016.objects.gj);
                geojson.features.forEach((feature) => {
                    const originalProps = feature.properties;
                    feature.properties = {
                        naam: originalProps?.statnaam || "Unknown",
                        code: originalProps?.statcode || "N/A",
                        rubriek: originalProps?.rubriek || ""
                    };
                });
                break;
            case "gemeenten 2012":
                geojson = feature(gemeentenTopojson_2012, gemeentenTopojson_2012.objects.gj);
                geojson.features.forEach((feature) => {
                    const originalProps = feature.properties;
                    feature.properties = {
                        naam: originalProps?.statnaam || "Unknown",
                        code: originalProps?.statcode || "N/A",
                        rubriek: originalProps?.rubriek || ""
                    };
                });
                break;
            case "wijken":
                geojson = feature(wijkenTopojson, wijkenTopojson.objects.gj);
                geojson.features.forEach((feature) => {
                    const originalProps = feature.properties;
                    feature.properties = {
                        naam: originalProps?.wijknaam || "Unknown",
                        code: originalProps?.wijkcode || "N/A",
                        rubriek: originalProps?.rubriek || ""
                    };
                });
                break;
            case "ggdregios":
                geojson = feature(ggdregiosTopojson, ggdregiosTopojson.objects.gj);
                geojson.features.forEach((feature) => {
                    const originalProps = feature.properties;
                    feature.properties = {
                        naam: originalProps?.statnaam || "Unknown",
                        code: originalProps?.statcode || "N/A",
                        rubriek: originalProps?.rubriek || ""
                    };
                });
                break;
            // case "werkgebieden":
            //     geojson = feature(KONwerkGebiedenTopojson, KONwerkGebiedenTopojson.objects.gj);
            //     geojson.features.forEach((feature) => {
            //         const originalProps = feature.properties;
            //         feature.properties = {
            //             naam: originalProps?.naam || "Unknown",
            //             code: originalProps?.code.toString() || "N/A",
            //             rubriek: originalProps?.rubriek || ""
            //         };
            //     });
            //     break;
            case "politieregios":
                geojson = feature(politieregiosTopojson, politieregiosTopojson.objects.gj);
                geojson.features.forEach((feature) => {
                    const originalProps = feature.properties;
                    feature.properties = {
                        naam: originalProps?.statnaam || "Unknown",
                        code: originalProps?.statcode || "N/A",
                        rubriek: originalProps?.rubriek || ""
                    };
                });
                break;
            case "veiligheidsregios":
                geojson = feature(veiligheidsregiosTopojson, veiligheidsregiosTopojson.objects.gj);
                geojson.features.forEach((feature) => {
                    const originalProps = feature.properties;
                    feature.properties = {
                        naam: originalProps?.statnaam || "Unknown",
                        code: originalProps?.statcode || "N/A",
                        rubriek: originalProps?.rubriek || ""
                    };
                });
                break;
            }
            this.baseGeojsonCache.set(mapKey, {
                type: geojson.type,
                features: geojson.features.map((f: any) => ({ ...f }))
            });
        }
        // }

        // Apply filter if a filter value is set
        if (this.filterValues.length > 0) {
            this.logDebugMessage(`Applying filter with values: ${this.filterValues.join(", ")}`);
            const originalCount = geojson.features.length;
            const filterLookup = this.buildRegionLookup(this.filterValues);
            geojson.features = geojson.features.filter(feature =>
                this.findRegionIndex(filterLookup, feature.properties?.naam, feature.properties?.code) !== -1
            );
            this.logDebugMessage(`Filtered from ${originalCount} to ${geojson.features.length} features`);
        }

        // Extract regions from the loaded map to use as category values
        // This ensures we use the actual map regions instead of user-provided 'Regio' field values
        const mapRegionValues = geojson.features.map(feature => 
            feature.properties.code || feature.properties.naam
        );
        this.logDebugMessage(`Extracted ${mapRegionValues.length} regions from map`);

        // Clear projection cache when the projection itself changes.
        // fitSize depends on viewport AND on which features remain (filtering zooms the map),
        // so cached pixel coordinates are invalid after a filter or map change.
        const currentFeatureCount = geojson.features.length;
        const currentViewportSize = {width, height};
        this.lastViewportSize = currentViewportSize;
        const projectionKey = [
            this.settings.mapSettings.mapSelectSettingsGroup.selectedMapSlice.value.value,
            currentFeatureCount,
            this.filterValues.join("|"),
            width,
            height
        ].join("~");
        if (this.lastProjectionKey !== projectionKey) {
            this.projectedPointsCache.clear();
            this.lastProjectionKey = projectionKey;
            this.cachedProjection = null;
        }

        // fitSize walks the whole geometry to compute bounds, so reuse it while the
        // map, viewport and visible region set are unchanged.
        if (!this.cachedProjection) {
            this.cachedProjection = geoMercator().fitSize([width, height], geojson);
        }
        const projection = this.cachedProjection;
        const path = geoPath().projection(projection);

        // Detect external highlight/selection early so points/clusters can be scoped
        const measureHighlights = categorical.values && categorical.values[0]
            ? categorical.values[0].highlights
            : undefined;
        const hasHighlights = !!(measureHighlights && measureHighlights.some((h: any) => h !== undefined && h !== null));

        // Scope points to Filter well / highlighted / hard-filtered regions.
        // One pass over the geometry yields both the in-scope points and the per-region
        // counts, so filtering and counting no longer walk every polygon separately.
        const pointScope = this.getPointScopeGeojson(geojson, hasHighlights, measureHighlights, categoryValues as any[]);
        let assignment = this.assignPointsToRegions(pointData, pointScope.restrict ? pointScope.geojson : geojson);

        let regionFilteredPointData = pointData;
        if (pointScope.restrict) {
            // Never blank out the map: if nothing matched (e.g. coordinates just outside
            // the region geometry), keep the unscoped points instead of showing none.
            if (assignment.inside.length > 0 || pointData.length === 0) {
                regionFilteredPointData = assignment.inside;
            } else {
                this.logDebugMessage("Region scoping matched no points, falling back to all points");
                assignment = this.assignPointsToRegions(pointData, geojson);
            }
            this.logDebugMessage(`Scoped points from ${pointData.length} to ${regionFilteredPointData.length} (highlight/filter)`);
        }

        // Apply data limiting for very large datasets
        const limitedPointData = this.limitPointData(regionFilteredPointData);
        
        // Count points per region if point data is available (using optimized method)
        const pointCountsPerRegion = new Map<string, number>();
        if (limitedPointData.length > 0) {
            this.logDebugMessage(`Processing ${limitedPointData.length} points for region counting`);
            
            assignment.counts.forEach((count, region) => pointCountsPerRegion.set(region, count));
            
            this.logDebugMessage(() => `Point counts per region: ${Array.from(pointCountsPerRegion.entries()).map(([region, count]) => `${region}=${count}`).join(", ")}`);
            
            // When coordinates are present but no "Waarde" field, treat point counts as "Waarde"
            // Use map regions as category values (not user-provided 'Regio' field)
            if (measureValues.length === 0 && mapRegionValues.length > 0) {
                // Map map regions to point counts to maintain relationship
                // This ensures measureValues has the same length as mapRegionValues for proper relationship
                this.measureValues = mapRegionValues.map((regionValue: string) => {
                    const regionStr = String(regionValue).trim();
                    
                    // Try to find matching region in pointCountsPerRegion by code or name
                    // First try exact match
                    if (pointCountsPerRegion.has(regionStr)) {
                        return pointCountsPerRegion.get(regionStr) || 0;
                    }
                    
                    // Try case-insensitive match
                    const regionLower = regionStr.toLowerCase();
                    for (const [regionKey, count] of pointCountsPerRegion.entries()) {
                        const regionKeyLower = regionKey.toLowerCase();
                        if (regionLower === regionKeyLower) {
                            return count;
                        }
                    }
                    
                    // PO1 matches PO01, but not PO10
                    for (const [regionKey, count] of pointCountsPerRegion.entries()) {
                        if (this.regionCodesMatch(regionLower, regionKey.toLowerCase())) {
                            return count;
                        }
                    }
                    
                    // If no match found, return 0
                    return 0;
                });
                
                // Ensure measureValues length matches mapRegionValues length for proper relationship
                if (this.measureValues.length !== mapRegionValues.length) {
                    this.logDebugMessage(`Warning: measureValues length (${this.measureValues.length}) doesn't match mapRegionValues length (${mapRegionValues.length})`);
                    // Pad or truncate to match
                    while (this.measureValues.length < mapRegionValues.length) {
                        this.measureValues.push(0);
                    }
                    this.measureValues = this.measureValues.slice(0, mapRegionValues.length);
                }
                
                this.logDebugMessage(`Using point counts as measure values (Waarde). Mapped ${this.measureValues.length} values from ${pointCountsPerRegion.size} regions.`);
                this.logDebugMessage(`Map regions sample: ${mapRegionValues.slice(0, 5).join(", ")}`);
                this.logDebugMessage(`Region keys sample: ${Array.from(pointCountsPerRegion.keys()).slice(0, 5).join(", ")}`);
                this.logDebugMessage(`Measure values sample: ${this.measureValues.slice(0, 5).join(", ")}`);
                this.isUsingPointCounts = true;
            } else if (measureValues.length === 0 && mapRegionValues.length === 0) {
                // No measure values and no map regions - can't create relationship
                this.logDebugMessage("Cannot use point counts as measure values: No regions in map");
                this.measureValues = [];
                this.isUsingPointCounts = false;
            } else {
                this.isUsingPointCounts = false;
            }
        } else {
            // No point data available
            this.isUsingPointCounts = false;
        }

        // Now do color scale caching with the correct values (after point counting)
        const currentValues = this.measureValues.filter(v => v != null && !isNaN(v));
        this.unfilteredMeasureValues = currentValues;

        // Initialize / expand cached scale when locked
        if (this.scaleLocked && currentValues.length > 0) {
            const curMin = Math.min(...currentValues);
            const curMax = Math.max(...currentValues);

            if (this.scaleMin === null || this.scaleMax === null) {
                // first seen domain
                this.scaleMin = curMin;
                this.scaleMax = curMax;
                this.logDebugMessage(`Initialized color scale cache: min=${this.scaleMin}, max=${this.scaleMax}`);
            } else {
                // expand-only to never shrink when filters are applied
                const oldMin = this.scaleMin;
                const oldMax = this.scaleMax;
                this.scaleMin = Math.min(this.scaleMin, curMin);
                this.scaleMax = Math.max(this.scaleMax, curMax);
                if (oldMin !== this.scaleMin || oldMax !== this.scaleMax) {
                    this.logDebugMessage(`Expanded color scale cache: min=${this.scaleMin}, max=${this.scaleMax}`);
                }
            }
        }

        // Map filter values to their original and lowercase versions
        const filterValueMap = this.filterValues.map(value => ({
            original: value,
            lower: value.toLowerCase()
        }));
        // Create sets for existing names and codes
        const existingNames = new Set(geojson.features.map(feature => feature.properties.naam.toLowerCase()));
        const existingCodes = new Set(geojson.features.map(feature => feature.properties.code.toLowerCase()));

        // Identify missing features based on filter values
        const missingFeatures = filterValueMap.filter(filterValue => 
            !existingNames.has(filterValue.lower) && !existingCodes.has(filterValue.lower)
        );

        if (this.settings.mapSettings.mapSelectSettingsGroup.showExtraFeatures.value.valueOf()) {

            // Create extra features for missing items
            // const svgWidth = this.svg.node().getBoundingClientRect().width;
            // const svgHeight = this.svg.node().getBoundingClientRect().height;
            const extraFeatures = missingFeatures.map((missingValue, index) => 
                this.createExtraFeature(missingValue.original, `EXTRA_${index}`, index, width, height)
            );

        // Append extra features (ensuring they are not projected)
            geojson.features.push(...extraFeatures.map(f => ({
                type: "Feature",
                properties: f.properties,
                geometry: f.geometry
            })));
        }

        
        // Check if there are any highlights from external selections (computed earlier for point scoping)
        
        this.svg.on("click", (event) => {
            if (event.target === this.svg.node()) {
                this.clearLegendHighlight();
            }
            // Only clear if clicking on empty space (not on a region)
            if (event.target === this.svg.node() && !hasHighlights) {
                this.selectionManager.clear().then(() => {
                    this.setSelectedPointKeys([]);
                    this.featuresGroup.selectAll<SVGPathElement, GeoJsonFeature>("path")
                        .attr("opacity", 1)
                        .style("stroke-width", (this.settings.mapSettings.regionBorderSettingsGroup.showRegionBorder.value ? this.settings.mapSettings.regionBorderSettingsGroup.regionBorderThickness.value : 0) + "px");
                });
            }
        });

        // Data and selection both need this region's position in the 'Regio' field,
        // so resolve it once per feature against a prebuilt index
        const categoryLookup = categoryValues.length > 0 ? this.buildRegionLookup(categoryValues) : null;
        const categoryColumnForSelection = (categorical.categories || []).find((cat: any) => cat.source.roles.category);

        // Set values and tooltips for all features before rendering
        geojson.features.forEach((d: GeoJsonFeature, index: number) => {
            // Map region key (code or name) for matching
            const regionKey = d.properties.code || d.properties.naam;
            
            const matchedCategoryIndex = categoryLookup
                ? this.findRegionIndex(categoryLookup, d.properties.naam, d.properties.code)
                : -1;
            const categoryIndexForData = matchedCategoryIndex;
            
            // Use the category index for data access, or fall back to map region index if no 'Regio' field
            d.categoryIndex = categoryIndexForData !== -1 ? categoryIndexForData : (index >= 0 && index < mapRegionValues.length ? index : -1);

            // Set value: prioritize measure values from dataView, then point counts, then index-based measureValues
            const rawVal = (categorical.values && categorical.values[0] && d.categoryIndex !== -1 && d.categoryIndex < categorical.values[0].values.length)
                ? categorical.values[0].values[d.categoryIndex]
                : undefined;
            
            if (rawVal !== null && rawVal !== undefined && !isNaN(Number(rawVal))) {
                // Use measure values from dataView when available
                d.value = Number(rawVal);
            } else if (this.isUsingPointCounts && this.measureValues.length > index) {
                // Use point counts mapped to map regions
                d.value = this.measureValues[index] || 0;
            } else if (pointCountsPerRegion.size > 0) {
                // Fallback to direct point counts lookup
                d.value = pointCountsPerRegion.get(regionKey) || 0;
            } else {
                // No data available
                d.value = NaN;
            }

            // Provisional selection id. matchedCategoryIndex indexes the 'Regio' column,
            // so this is only valid when that is also the column we select against; the
            // block at the end of update() reassigns every feature authoritatively.
            const canSelectOnCategory = !!(this.regionCategoryColumn
                && this.regionCategoryColumn === categoryColumnForSelection
                && matchedCategoryIndex !== -1);

            d.selectionId = canSelectOnCategory
                ? this.host.createSelectionIdBuilder()
                    .withCategory(this.regionCategoryColumn, matchedCategoryIndex)
                    .createSelectionId()
                : null as any;

            // Find the category column (Regio field) - the one with category role
            const categoryColumn = categorical.categories && categorical.categories.length > 0
                ? categorical.categories.find((cat: any) => cat.source.roles.category)
                : undefined;
            // Use dynamic region type label based on selected map
            const selectedMap = this.settings.mapSettings.mapSelectSettingsGroup.selectedMapSlice.value;
            const categoryName = this.getRegionTypeLabel(selectedMap);

            // Find the measure column (assuming only one)
            const measureColumn = (categorical.values && categorical.values.find) ? categorical.values.find((v: any) => v.source.roles.measure) : undefined;
            const measureName = measureColumn?.source?.displayName ?? "Waarde";

            // Create tooltip data
            d.tooltipData = [];

            // Always add region name from topojson
            d.tooltipData.push({ displayName: categoryName, value: d.properties.naam });

            // Add measure value or point count with formatting (prioritize measure values)
            if (categorical.values && categorical.values[0] && d.categoryIndex !== -1 && d.categoryIndex < categorical.values[0].values.length) {
                const measureValue = categorical.values[0];
                const rawValue = measureValue.values[d.categoryIndex];
                const formatString = (measureValue.source && measureValue.source.format) || "General";

                if (rawValue !== null && rawValue !== undefined && !isNaN(Number(rawValue))) {
                    // Show measure value when available
                    const formattedValue = formatString === "General" ? rawValue.toString() : valueFormatter.format(rawValue, formatString);
                    d.tooltipData.push({ displayName: measureName, value: formattedValue });
                } else if (pointCountsPerRegion.size > 0) {
                    // Show point count as "Waarde" when no measure values
                    const pointCount = pointCountsPerRegion.get(regionKey) || 0;
                    d.tooltipData.push({ displayName: "Waarde", value: pointCount.toString() });
                }
            } else if (pointCountsPerRegion.size > 0) {
                // Show point count as "Waarde" when no measure values
                const pointCount = pointCountsPerRegion.get(regionKey) || 0;
                d.tooltipData.push({ displayName: "Waarde", value: pointCount.toString() });
            }
            
            // First check if we have any values
            // Add additional tooltip fields and check for custom label
            // Tooltip fields can be in categorical.values (if kind: "Measure") or categorical.categories (if kind: "Grouping")
            // IMPORTANT: When checking categories, exclude the 'Regio' field (category column) by comparing column identity
            // This works regardless of field name (statcode, GemeenteCode, etc.)
            
            // Check categorical.values for tooltip fields (if they're measures)
            // Include tooltip fields from measures (tooltip role),
            // but skip the primary measure already shown above
            if (categorical.values && Array.isArray(categorical.values)) {
                categorical.values.forEach((value: any) => {
                    if (value.source.roles.tooltip) {
                        // Avoid duplicating the main measure row
                        const isPrimaryMeasure = measureColumn && (value.source.queryName === measureColumn.source.queryName);
                        if (isPrimaryMeasure) return;

                        const rawValue = (d.categoryIndex !== -1 && d.categoryIndex < value.values.length)
                            ? value.values[d.categoryIndex]
                            : undefined;
                        const formatString = value.source.format || ""; // Get the format string from the metadata

                        // Check if the value is null or undefined
                        const formattedValue = rawValue !== null && rawValue !== undefined
                            ? (formatString === "General" || formatString === "" ? rawValue.toString() : valueFormatter.format(rawValue, formatString))
                            : "N/A";

                        d.tooltipData.push({
                            displayName: value.source.displayName,
                            value: formattedValue
                        });
                    }
                });
            }
            
            // Check categorical.categories for tooltip fields (if they're groupings)
            // Skip only the first tooltip category encountered, include the rest
            if (categorical.categories && Array.isArray(categorical.categories)) {
                let skippedFirstTooltipCategory = false;
                categorical.categories.forEach((category: any) => {
                    // Only include fields with tooltip role
                    if (category.source.roles.tooltip) {
                        // If this is the first tooltip category we encounter, skip it once
                        if (!skippedFirstTooltipCategory) {
                            if (category.source.queryName === categoryColumn.source.queryName){
                                skippedFirstTooltipCategory = true;
                                return;
                            }
                        }

                        // Use categoryIndex to get the correct value from the category array
                        const rawValue = (d.categoryIndex !== -1 && d.categoryIndex < category.values.length)
                            ? category.values[d.categoryIndex]
                            : undefined;
                        const formatString = category.source.format || ""; // Get the format string from the metadata

                        // Check if the value is null or undefined
                        const formattedValue = rawValue !== null && rawValue !== undefined
                            ? (formatString === "General" || formatString === "" ? rawValue.toString() : valueFormatter.format(rawValue, formatString))
                            : "N/A";

                        d.tooltipData.push({
                            displayName: category.source.displayName,
                            value: formattedValue
                        });
                    }
                });
            }
            
            // Custom label measure (optional) — format with field format / numOfDigits
            const labelMeasureColumn = categorical.values && Array.isArray(categorical.values)
                ? categorical.values.find((value: any) => value.source?.roles?.labelMeasure)
                : undefined;
            if (labelMeasureColumn && d.categoryIndex !== -1) {
                const raw = labelMeasureColumn.values[d.categoryIndex];
                d.customLabel = this.formatCustomLabelValue(raw, labelMeasureColumn.source?.format);
            } else {
                d.customLabel = "";
            }
           
            // Check if this region is highlighted from external selection
            d.isHighlighted = !hasHighlights || 
            (   categorical.values && 
                categorical.values[0] && 
                categorical.values[0].highlights && 
                d.categoryIndex !== -1 && // Make sure we have a valid category index
                categorical.values[0].highlights[d.categoryIndex] !== undefined && // Check if this specific region is highlighted
                categorical.values[0].highlights[d.categoryIndex] !== null); // Also check it's not null
        });

        this.createLegend();
        this.createPointLegend();
        
        const regions = this.featuresGroup.selectAll<SVGPathElement, GeoJsonFeature>("path")
        .data((geojson as GeoJsonData).features, d => d.properties.code || d.properties.naam);


        const regionsEnter = regions.enter()
            .append("path")
            .classed("feature", true);

        const buildPathData = (d: GeoJsonFeature) => {
            if (d.properties.rubriek === "custom") {
                // For extra features, use the unprojected coordinates
                return d3.geoPath()(d.geometry as GeoPermissibleObjects);
            } else {
                // For main map features, use the projected path
                return path(d.geometry as GeoPermissibleObjects);
            }
        };

        // Serialising path data is the most expensive part of a render on the detailed
        // maps, and the result only changes when the projection does. Cross-filter
        // updates (which only restyle) reuse the existing "d" attributes.
        const projectionChanged = this.lastRenderedProjectionKey !== projectionKey;
        this.lastRenderedProjectionKey = projectionKey;
        if (projectionChanged) {
            regionsEnter.merge(regions).attr("d", buildPathData);
        } else {
            regionsEnter.attr("d", buildPathData);
            // Extra features are laid out against the live zoom level, not the projection
            regionsEnter.merge(regions)
                .filter(d => d.properties.rubriek === "custom")
                .attr("d", buildPathData);
        }

        regionsEnter
            .merge(regions)
            .attr("fill", d => (this.unfilteredMeasureValues && this.unfilteredMeasureValues.length > 0 && !isNaN(d.value)) ? this.getColor(d.value) : "#e0e0e0")
            .attr("opacity", (d) => {
                if (hasHighlights) {
                    return d.isHighlighted ? 1 : 0.5;
                }
                return 1;
            })
            .style("stroke", (d) => {
                if (hasHighlights) {
                    return d.isHighlighted ? this.settings.mapSettings.regionBorderSettingsGroup.regionBorderColor.value.value : 
                           (this.settings.mapSettings.regionBorderSettingsGroup.showRegionBorder.value ? this.settings.mapSettings.regionBorderSettingsGroup.regionBorderColor.value.value : "none");
                }
                return this.settings.mapSettings.regionBorderSettingsGroup.showRegionBorder.value ? this.settings.mapSettings.regionBorderSettingsGroup.regionBorderColor.value.value : "none";
            })
            .style("stroke-width", (d) => {
                const thickness = this.settings.mapSettings.regionBorderSettingsGroup.regionBorderThickness.value || 5;
                const showBorder = this.settings.mapSettings.regionBorderSettingsGroup.showRegionBorder.value;
                
                if (hasHighlights) {
                    return (d.isHighlighted ? Math.max(2, thickness) : (showBorder ? thickness : 0)) + "px";
                }
                return (showBorder ? thickness : 0) + "px";
            })

            // Bind tooltip to regions
            .on("mouseover", (event, d) => {
                if (!this.settings.tooltipSettings.tooltipSettingsGroup.showTooltip.value) {
                    return;  // Exit if region tooltips disabled
                }
                if (d.tooltipData) {
                        this.visualHostTooltipService.show({
                            coordinates: [event.pageX, event.pageY],
                            isTouchEvent: false,
                            dataItems: d.tooltipData,
                            identities: d.selectionId ? [d.selectionId] : [],
                        });
                    }        
            })
            .on("mousemove", (event, d) => {
                if (this.settings.tooltipSettings.tooltipSettingsGroup.showTooltip.value) {
                    this.visualHostTooltipService.move({
                        coordinates: [event.pageX, event.pageY],
                        isTouchEvent: false,
                        identities: d.selectionId ? [d.selectionId] : []
                    });
                }
            })
            .on("mouseout", (event) => {
                // Hide tooltip only when leaving the entire feature
                if (this.settings.tooltipSettings.tooltipSettingsGroup.showTooltip.value) {
                    this.visualHostTooltipService.hide({
                        immediately: true,
                        isTouchEvent: false,
                    });
                }
            })
            .on("click", (event, d) => {
                this.logDebugMessage(`Region clicked: ${d.properties.naam} (${d.properties.code})`);
                if (this.selectRegionFeature(event, d)) {
                    event.stopPropagation();
                }
            })
            
        regions.exit().remove(); // Remove old regions
        
        // Drop stale point dimming if the selection was cleared elsewhere (e.g. by
        // another visual or the report's clear-filters button)
        if (this.selectedPointKeys && !this.selectionManager.hasSelection()) {
            this.selectedPointKeys = null;
        }

        // Render point data on the map (individual points or clusters)
        if (this.settings.pointSettings.pointSettingsGroup.showPoints.value && limitedPointData.length > 0) {
            if (this.settings.pointSettings.clusterSettingsGroup.enableClustering.value) {
                // Zoom-based geographic clustering (merges out / splits in)
                this.renderClusters(limitedPointData, projection, pointTooltipValues, legendFieldName);
            } else {
                this.clusterProjectedPoints = null;
                // Render individual points
                this.renderIndividualPoints(limitedPointData, projection, pointTooltipValues, legendFieldName);
            }
        } else {
            // Remove all points/clusters if showPoints is false or no data
            this.clusterProjectedPoints = null;
            this.pointsGroup.selectAll("*").remove();
        }
        
        //const filteredFeatures = geojson.features;

        const labels = this.labelsGroup.selectAll<SVGGElement, GeoJsonFeature>("g.label")
            .data((geojson as GeoJsonData).features, d => d.properties.code || d.properties.naam); // Use a key function if possible

        // Enter: Add new labels
        const labelsEnter = labels.enter()
            .append("g")
            .classed("label", true);

        labelsEnter.append("rect");
        labelsEnter.append("text");

        // Merge: Update existing and newly added labels
        labelsEnter.merge(labels)
            .each((d: GeoJsonFeature, i, nodes) => {
                const group = d3.select(nodes[i]);
                const fontSize = this.settings.labelSettings.labelSettingsGroup.fontSize.value;
                const centroid = d.properties.rubriek === "custom"
                    ? [
                        (d.geometry.coordinates[0][0][0] + d.geometry.coordinates[0][2][0]) / 2,
                        d.geometry.coordinates[0][2][1] + (fontSize * 1.5 )/ d3.zoomTransform(this.svg.node()).k
                    ]
                    : path.centroid(d as GeoPermissibleObjects);
                
                
                let labelText: string;
                if (this.settings.labelSettings.labelSettingsGroup.useCustomLabel.value.valueOf()) {
                    labelText = d.customLabel && d.customLabel.length > 0
                        ? d.customLabel
                        : "---";
                } else {
                    // Prepare label text
                    const labelParts: string[] = [];
                    if (this.localShowRegionName && d.properties.rubriek !== "custom") {
                        labelParts.push(d.properties.naam);
                    }
                    if (this.settings.mapSettings.mapSelectSettingsGroup.showExtraFeatures.value.valueOf() && d.properties.rubriek === "custom") {
                        labelParts.push(d.properties.naam);
                    }
                    // Add value with proper formatting
                if (this.localShowValues && this.measureValues && this.measureValues.length > 0) {
                    if (this.isUsingPointCounts) {
                        // For point counts, show as integer
                        labelParts.push(d.value.toString());
                    } else {
                        // For regular measure values, use formatting
                        const measureValue = this.dataViews && this.dataViews[0] && this.dataViews[0].categorical && this.dataViews[0].categorical.values
                            ? this.dataViews[0].categorical.values[0] : undefined;
                        const formatString = (measureValue && measureValue.source && measureValue.source.format) || "General";
                        const formattedValue = formatString === "General" 
                            ? d.value.toFixed(this.settings.labelSettings.labelSettingsGroup.numOfDigits.value)
                            : valueFormatter.format(d.value, formatString);
                        labelParts.push(formattedValue);
                    }
                }

                    // Join parts
                    if (this.settings.mapSettings.mapSelectSettingsGroup.showExtraFeatures.value.valueOf() && d.properties.rubriek === "custom") {
                        labelText = labelParts.join(": ");
                    } else {
                        labelText = labelParts.join("\n");
                    }
                }
                
                // Check if there's actually any content to display
                const hasContent = labelText && labelText.trim().length > 0;
                
                // Update text with visibility check
                const text = group.select("text")
                    .attr("x", centroid[0])
                    .attr("y", centroid[1])
                    .attr("text-anchor", "middle")
                    .attr("font-size", fontSize)
                    .attr("fill", "black")
                    .style("display", (this.localShowLabels && hasContent) ? null : "none");

                // Clear and re-add tspans for multiline text
                text.selectAll("tspan").remove();
                if (hasContent) {
                labelText.split("\n").forEach((line, index) => {
                    text.append("tspan")
                        .attr("x", centroid[0])
                        .attr("dy", index === 0 ? "0" : "1.2em")
                        .text(line);
                });
                }

                // Update background rectangle - only call getBBox when necessary
                const rect = group.select("rect");
                if (hasContent && this.localShowLabels) {
                    const textNode = text.node() as SVGTextElement;
                    if (textNode) {
                        try {
                            const textBBox = textNode.getBBox();
                            // Get current zoom transform to scale properties correctly
                            const currentTransform = d3.zoomTransform(this.svg.node());
                            const padding = 3 / currentTransform.k;

                            rect.attr("x", textBBox.x - padding)
                                .attr("y", textBBox.y - padding)
                                .attr("rx", 5 / currentTransform.k)
                                .attr("ry", 5 / currentTransform.k)
                                .attr("width", textBBox.width + 2 * padding)
                                .attr("height", textBBox.height + 2 * padding)
                                .attr("fill", this.settings.labelSettings.labelSettingsGroup.showLabelBackground.value.valueOf() ? this.settings.labelSettings.labelSettingsGroup.labelBackgroundColor.value.value : "none")
                                .attr("opacity", this.settings.labelSettings.labelSettingsGroup.labelBackgroundOpacity.value / 100)
                                .attr("stroke", this.settings.labelSettings.labelSettingsGroup.showLabelBorder.value.valueOf() ? (this.settings.labelSettings.labelSettingsGroup as any).labelBorderColor?.value?.value : "none")
                                .style("stroke-width", (this.settings.labelSettings.labelSettingsGroup.showLabelBorder.value.valueOf() ? ((this.settings.labelSettings.labelSettingsGroup as any).labelBorderThickness?.value ?? 1) / currentTransform.k : 0) + "px");
                        } catch (error) {
                            // If getBBox fails, skip background rectangle update
                            console.warn("Failed to get text bounding box for label:", error);
                        }
                    }
                }
                rect.style("display", (this.localShowLabels && hasContent) ? null : "none");
            });
        // Exit: Remove old labels
        labels.exit().remove();

        // Ensure selectionId is assigned to each feature
        // Match map regions to categoryValues (from 'Regio' field) for selection
        const selectionCategoryColumn = this.regionCategoryColumn || null;
        const selectionLookup = selectionCategoryColumn
            ? this.buildRegionLookup(selectionCategoryColumn.values as any[])
            : null;

        geojson.features.forEach((feature) => {
            if (!selectionCategoryColumn || !selectionLookup) {
                feature.selectionId = null as any;
                return;
            }
            const categoryColumn = selectionCategoryColumn;
            const categoryIndex = this.findRegionIndex(selectionLookup, feature.properties.naam, feature.properties.code);
            if (categoryIndex !== -1) {
                const selectionId = this.host.createSelectionIdBuilder()
                    .withCategory(categoryColumn, categoryIndex)
                    .createSelectionId();
                feature.selectionId = selectionId;
            } else {
                feature.selectionId = null as any;
            }
        });
        // Get current zoom transform
        // const currentTransform = d3.zoomTransform(this.svg.node());
        
        // this.updateLabelScaling(currentTransform);
        
    } 

    private updateSelection(ids: powerbi.visuals.ISelectionId[]) {
        this.logDebugMessage("Starting updateSelection");
        const regions = this.featuresGroup.selectAll<SVGPathElement, GeoJsonFeature>("path");
        this.logDebugMessage(`Found ${regions.size()} regions`);


        regions.each((d, i, nodes) => {
            this.logDebugMessage(`Checking feature ${d.properties.naam}(${d.properties.code}) with selectionId: ${d.selectionId?.getKey()}`);
            const isSelected = ids.some(id => id.equals(d.selectionId));
            this.logDebugMessage(`Feature ${d.properties.naam}(${d.properties.code}) is selected: ${isSelected}`);
            d3.select(nodes[i])
                .attr("opacity", isSelected ? 1 : 0.5)
                .style("stroke-width", (isSelected ? Math.max(2, this.settings.mapSettings.regionBorderSettingsGroup.regionBorderThickness.value) : 
                       (this.settings.mapSettings.regionBorderSettingsGroup.showRegionBorder.value ? this.settings.mapSettings.regionBorderSettingsGroup.regionBorderThickness.value : 0)) + "px");
        });
    }

    
    private getColor(value: number): string {
        if (value === undefined || value === null || isNaN(value)) return "#e0e0e0";

        const minColor = this.settings.mapSettings.mapColorSettingsGroup.minColorSlice.value.value;
        const midColor = this.settings.mapSettings.mapColorSettingsGroup.midColorSlice.value.value;
        const maxColor = this.settings.mapSettings.mapColorSettingsGroup.maxColorSlice.value.value;
        const useMidColor = this.settings.mapSettings.mapColorSettingsGroup.useMidColorSlice.value;
        const useAbsoluteValues = ((this.settings.mapSettings.mapColorSettingsGroup as any).useAbsoluteValues?.value) ?? false;

        const minPos = this.settings.mapSettings.mapColorSettingsGroup.minPosition.value;
        const midPos = this.settings.mapSettings.mapColorSettingsGroup.midPosition.value;
        const maxPos = this.settings.mapSettings.mapColorSettingsGroup.maxPosition.value;

        // Prefer cached global scale when locked
        const minValue = (this.scaleLocked && this.scaleMin != null) ? this.scaleMin :
                         (this.unfilteredMeasureValues.length ? Math.min(...this.unfilteredMeasureValues) : NaN);
        const maxValue = (this.scaleLocked && this.scaleMax != null) ? this.scaleMax :
                         (this.unfilteredMeasureValues.length ? Math.max(...this.unfilteredMeasureValues) : NaN);

        if (!isFinite(minValue) || !isFinite(maxValue)) return "#e0e0e0";
        if (maxValue === minValue) return useMidColor ? midColor : maxColor;

        // Convert positions to percentages if using absolute values
        let minPosPercent, midPosPercent, maxPosPercent;
        let valuePercent;
        
        if (useAbsoluteValues && maxPos !== minPos) {
            // When using absolute values, use the minPos-maxPos range as the color scale
            // This matches the legend bar positioning
            // Gradient offset: 0% = bottom (minPos), 100% = top (maxPos)
            const range = maxPos - minPos;
            minPosPercent = 0;   // minPos at bottom (offset 0%)
            maxPosPercent = 100; // maxPos at top (offset 100%)
            
            if (useMidColor && range !== 0) {
                // Calculate where midPos falls proportionally between minPos and maxPos
                midPosPercent = ((midPos - minPos) / range) * 100;
            } else {
                midPosPercent = 50; // Default middle if not used
            }
            
            // Convert the actual value to a percentage within the minPos-maxPos range
            // 0% = minPos (bottom), 100% = maxPos (top)
            if (range !== 0) {
                valuePercent = ((value - minPos) / range) * 100;
            } else {
                valuePercent = 50; // Default if range is 0
            }
        } else if (useAbsoluteValues) {
            // Edge case: minPos === maxPos
            minPosPercent = 0;
            maxPosPercent = 100;
            midPosPercent = 50;
            valuePercent = 50;
        } else {
            // Use positions as percentages (current behavior)
            minPosPercent = minPos;
            midPosPercent = midPos;
            maxPosPercent = maxPos;
            // Convert value to percentage based on data range
            valuePercent = ((value - minValue) / (maxValue - minValue)) * 100;
        }

        const colorScale = d3.scaleLinear<string>()
            .domain(useMidColor ? [minPosPercent, midPosPercent, maxPosPercent] : [minPosPercent, maxPosPercent])
            .range(useMidColor ? [minColor, midColor, maxColor] : [minColor, maxColor]);

        return colorScale(valuePercent);
    }

    private renderIndividualPoints(pointData: PointDatum[], projection: any, pointTooltipValues?: any[], legendFieldName?: string) {
        // Remove any existing cluster groups when switching to individual points
        this.pointsGroup.selectAll<SVGGElement, any>("g.cluster").remove();
        
        // Pre-compute projections for all points
        const projectedPoints = this.precomputeProjections(pointData, projection);
        
        const points = this.pointsGroup.selectAll<SVGCircleElement, ProjectedPointDatum>("circle")
            .data(projectedPoints);
        
        const defaultPointColor = this.settings.pointSettings.pointSettingsGroup.pointColor.value.value;
        const pointSize = this.settings.pointSettings.pointSettingsGroup.pointSize.value;
        const pointOpacity = this.settings.pointSettings.pointSettingsGroup.pointOpacity.value / 100;
        const pointBorderColor = this.settings.pointSettings.pointSettingsGroup.pointBorderColor.value.value;
        const pointBorderThickness = this.settings.pointSettings.pointSettingsGroup.pointBorderThickness.value;
        const useLegendColors = this.settings.pointSettings.pointSettingsGroup.useLegendColors.value;
        
        points.enter()
            .append("circle")
            .merge(points)
            .attr("cx", d => d.projectedX)  // Use pre-computed projection
            .attr("cy", d => d.projectedY)  // Use pre-computed projection
            .attr("data-base-r", pointSize)
            .attr("data-base-stroke", pointBorderThickness)
            .attr("r", pointSize)
            .attr("fill", d => {
                if (useLegendColors && d.legend) {
                    return this.getLegendColor(d.legend);
                }
                return defaultPointColor;
            })
            .attr("data-base-opacity", pointOpacity)
            .attr("opacity", pointOpacity)
            .attr("stroke", pointBorderColor)
            .attr("stroke-width", pointBorderThickness)
            // Add tooltip functionality for points
            .on("mouseover", (event, d) => {
                if (!this.settings.tooltipSettings.tooltipSettingsGroup.showPointTooltip.value) {
                    return;  // Exit if point tooltips disabled
                }
                
                const tooltipData = this.createPointTooltipData(d, projectedPoints.indexOf(d), pointTooltipValues, legendFieldName);
                if (tooltipData.length > 0) {
                    this.visualHostTooltipService.show({
                        coordinates: [event.pageX, event.pageY],
                        isTouchEvent: false,
                        dataItems: tooltipData,
                        identities: [] // Points don't have selection IDs
                    });
                }
            })
            .on("mousemove", (event, d) => {
                if (this.settings.tooltipSettings.tooltipSettingsGroup.showPointTooltip.value) {
                    this.visualHostTooltipService.move({
                        coordinates: [event.pageX, event.pageY],
                        isTouchEvent: false,
                        identities: []
                    });
                }
            })
            .on("mouseout", (event) => {
                if (this.settings.tooltipSettings.tooltipSettingsGroup.showPointTooltip.value) {
                    this.visualHostTooltipService.hide({
                        immediately: true,
                        isTouchEvent: false,
                    });
                }
            })
            .on("click", (event, d) => {
                this.selectPointIdentities(event, [d.selectionId]);
            });
        
        points.exit().remove();
        this.applyPointHighlightStyles();
        this.updatePointScaling(d3.zoomTransform(this.svg.node()));
    }

    /**
     * Category colours live in fixed slots. The slot a category owns is remembered in
     * the report (legendColors.categoryOrder) rather than derived from the categories
     * currently in view — otherwise filtering out one category shifts every later
     * category onto a different colour.
     */
    private getLegendColorSlot(legendValue: string): string | null {
        const index = this.legendSlotOrder.indexOf(legendValue);
        if (index < 0 || index >= this.maxLegendColorSlots) {
            return null;
        }
        return `color${index + 1}`;
    }

    private getPersistedCategoryOrder(): string[] {
        const legendColors = this.dataView?.metadata?.objects?.["legendColors"] as any;
        const raw = legendColors?.categoryOrder;
        if (typeof raw !== "string" || !raw) {
            return [];
        }
        try {
            const parsed = JSON.parse(raw);
            return Array.isArray(parsed) ? parsed.filter((v: any) => typeof v === "string") : [];
        } catch {
            return [];
        }
    }

    /**
     * Give every category a permanent slot: keep the order already saved in the report
     * and append categories the report has not seen before. Only writes back when the
     * list actually grew, so this cannot loop on itself.
     */
    private resolveLegendSlots() {
        const order = this.getPersistedCategoryOrder().slice(0, this.maxLegendColorSlots);
        let grew = false;

        this.legendCategories.forEach(category => {
            if (order.indexOf(category) < 0 && order.length < this.maxLegendColorSlots) {
                order.push(category);
                grew = true;
            }
        });

        this.legendSlotOrder = order;

        if (grew) {
            try {
                this.host.persistProperties({
                    merge: [{
                        objectName: "legendColors",
                        selector: null,
                        properties: { categoryOrder: JSON.stringify(order) }
                    }]
                } as any);
            } catch (error) {
                this.logDebugMessage("Could not persist legend category order");
            }
        }
    }

    private extractFillColor(value: any): string | null {
        if (!value) {
            return null;
        }
        if (typeof value === "string") {
            return value;
        }
        if (value.solid && value.solid.color) {
            return value.solid.color;
        }
        if (value.value) {
            if (typeof value.value === "string") {
                return value.value;
            }
            if (value.value.value) {
                return value.value.value;
            }
        }
        return null;
    }

    private getPersistedLegendColor(legendValue: string): string | null {
        const slot = this.getLegendColorSlot(legendValue);
        if (!slot) {
            return null;
        }
        const legendColors = this.dataView?.metadata?.objects?.["legendColors"] as any;
        return legendColors ? this.extractFillColor(legendColors[slot]) : null;
    }

    private getLegendColor(legendValue: string): string {
        // A colour picked in the formatting pane always wins
        const persisted = this.getPersistedLegendColor(legendValue);
        if (persisted) {
            return persisted;
        }

        // Colours saved by older versions were bound to row identities
        if (this.legendCategoryColumn && this.legendCategoryColumn.objects) {
            // Create a stable mapping by checking all objects for this legend value
            // This ensures colors are preserved even when data is filtered
            for (let i = 0; i < this.legendCategoryColumn.values.length; i++) {
                if (String(this.legendCategoryColumn.values[i]) === legendValue) {
                    const obj = this.legendCategoryColumn.objects[i] as any;
                    if (obj && obj.legendColors && obj.legendColors.fill) {
                        const colorObj = obj.legendColors.fill;
                        
                        // Try different possible structures
                        if (typeof colorObj === 'string') {
                            return colorObj;
                        }
                        if (colorObj.solid && colorObj.solid.color) {
                            return colorObj.solid.color;
                        }
                        if (colorObj.value) {
                            if (typeof colorObj.value === 'string') {
                                return colorObj.value;
                            }
                            if (colorObj.value.value) {
                                return colorObj.value.value;
                            }
                        }
                    }
                }
            }
        }
        
        const modelColor = this.getLegendColorFromModel(legendValue);
        if (modelColor) {
            return modelColor;
        }

        // No persisted color yet: use the same default palette the formatting pane
        // shows for this category. Keyed off the stable slot so the default a category
        // falls back to does not change when other categories are filtered away.
        const slotIndex = this.legendSlotOrder.indexOf(legendValue);
        if (slotIndex >= 0) {
            return this.getDefaultColorForIndex(slotIndex);
        }

        // Default color if not set
        return this.settings.pointSettings.pointSettingsGroup.pointColor.value.value;
    }

    private renderClusters(
        pointData: PointDatum[],
        projection: any,
        pointTooltipValues?: any[],
        legendFieldName?: string
    ) {
        // Clear existing point/cluster marks
        this.pointsGroup.selectAll("*").remove();

        const projectedPoints = this.precomputeProjections(pointData, projection);
        this.clusterProjectedPoints = projectedPoints;
        this.clusterPointTooltipValues = pointTooltipValues;
        this.clusterLegendFieldName = legendFieldName;
        this.lastClusterZoomBucket = null;

        const transform = d3.zoomTransform(this.svg.node());
        this.drawZoomClusters(transform, true);
    }

    /**
     * Screen-space grid clustering: zoomed out → fewer larger clusters;
     * zoomed in → more/smaller clusters (and singles).
     */
    private buildZoomClusterData(
        transform: d3.ZoomTransform
    ): Array<ClusterDatum> {
        const points = this.clusterProjectedPoints || [];
        if (!points.length) {
            return [];
        }

        const k = transform.k || 1;
        const cell = this.clusterScreenCellSize;
        const buckets = new Map<string, {x: number, y: number, legendCounts: Map<string, number>, points: ProjectedPointDatum[]}>();

        points.forEach(point => {
            const screenX = point.projectedX * k + transform.x;
            const screenY = point.projectedY * k + transform.y;
            const key = `${Math.floor(screenX / cell)},${Math.floor(screenY / cell)}`;
            let bucket = buckets.get(key);
            if (bucket) {
                bucket.x += point.projectedX;
                bucket.y += point.projectedY;
            } else {
                bucket = {
                    x: point.projectedX,
                    y: point.projectedY,
                    legendCounts: new Map<string, number>(),
                    points: []
                };
                buckets.set(key, bucket);
            }
            bucket.points.push(point);
            if (point.legend) {
                bucket.legendCounts.set(point.legend, (bucket.legendCounts.get(point.legend) || 0) + 1);
            }
        });

        const clusterData: Array<ClusterDatum> = [];
        buckets.forEach((bucket, key) => {
            const count = bucket.points.length;
            clusterData.push({
                id: key,
                count,
                centroid: [bucket.x / count, bucket.y / count],
                legend: count === 1 ? bucket.points[0].legend : undefined,
                legendCounts: bucket.legendCounts,
                points: bucket.points
            });
        });
        return clusterData;
    }

    /**
     * How many rows in this cluster are currently in focus — matching the selected
     * legend categories and part of any active row selection.
     */
    private countMatchingInCluster(cluster: ClusterDatum): number {
        if (!this.isLegendHighlightActive() && !this.selectedPointKeys) {
            return cluster.count;
        }
        return cluster.points.filter(point =>
            this.legendMatches(point.legend) && this.isPointSelected(point)
        ).length;
    }

    private drawZoomClusters(transform: d3.ZoomTransform, force: boolean = false) {
        if (!this.clusterProjectedPoints || !this.clusterProjectedPoints.length) {
            return;
        }

        const bucket = this.getClusterZoomBucket(transform.k || 1);
        if (!force && this.lastClusterZoomBucket === bucket) {
            // Same aggregation level — only keep bubble screen size stable
            this.updatePointScaling(transform);
            return;
        }
        this.lastClusterZoomBucket = bucket;

        const clusterData = this.buildZoomClusterData(transform);
        const maxCount = Math.max(1, ...clusterData.map(c => c.count));
        const minSize = this.settings.pointSettings.clusterSettingsGroup.minClusterSize.value;
        const maxSize = this.settings.pointSettings.clusterSettingsGroup.maxClusterSize.value;
        const clusterColor = this.settings.pointSettings.clusterSettingsGroup.clusterColor.value.value;
        const clusterOpacity = this.settings.pointSettings.clusterSettingsGroup.clusterOpacity.value / 100;
        const clusterBorderColor = this.settings.pointSettings.clusterSettingsGroup.clusterBorderColor.value.value;
        const clusterBorderThickness = this.settings.pointSettings.clusterSettingsGroup.clusterBorderThickness.value;
        const showCount = this.settings.pointSettings.clusterSettingsGroup.showClusterCount.value;
        const countColor = this.settings.pointSettings.clusterSettingsGroup.clusterCountColor.value.value;
        const countSize = this.settings.pointSettings.clusterSettingsGroup.clusterCountSize.value;
        const pointSize = this.settings.pointSettings.pointSettingsGroup.pointSize.value;
        const useLegendColors = this.settings.pointSettings.pointSettingsGroup.useLegendColors.value;
        const defaultPointColor = this.settings.pointSettings.pointSettingsGroup.pointColor.value.value;

        const highlightActive = this.isLegendHighlightActive() || !!this.selectedPointKeys;
        // With one category selected we can paint the highlight in that category's own colour
        const singleSelection = this.selectedLegendCategories.size === 1
            ? Array.from(this.selectedLegendCategories)[0]
            : null;
        const highlightFill = (singleSelection && useLegendColors)
            ? this.getLegendColor(singleSelection)
            : clusterColor;
        const radiusFor = (count: number) => count <= 1
            ? pointSize
            : minSize + (maxSize - minSize) * Math.sqrt(count / maxCount);

        const clusters = this.pointsGroup.selectAll<SVGGElement, any>("g.cluster")
            .data(clusterData, (d: any) => d.id);

        const clusterEnter = clusters.enter()
            .append("g")
            .classed("cluster", true);

        clusterEnter.append("circle").classed("cluster-base", true);
        clusterEnter.append("circle").classed("cluster-highlight", true);
        clusterEnter.append("text");

        const merged = clusterEnter.merge(clusters as any);

        merged.each((d, i, nodes) => {
            const group = d3.select(nodes[i]);
            const circle = group.select("circle.cluster-base");
            const highlightCircle = group.select("circle.cluster-highlight");
            const text = group.select("text");

            const isSingle = d.count === 1;
            const clusterSize = radiusFor(d.count);
            const fill = (isSingle && useLegendColors && d.legend)
                ? this.getLegendColor(d.legend)
                : (isSingle ? defaultPointColor : clusterColor);
            const matchCount = this.countMatchingInCluster(d);
            const baseOpacity = isSingle
                ? this.settings.pointSettings.pointSettingsGroup.pointOpacity.value / 100
                : clusterOpacity;

            circle
                .attr("cx", d.centroid[0])
                .attr("cy", d.centroid[1])
                .attr("data-base-r", clusterSize)
                .attr("data-base-stroke", clusterBorderThickness)
                .attr("r", clusterSize)
                .attr("fill", fill)
                .attr("data-base-opacity", baseOpacity)
                .attr("opacity", highlightActive ? baseOpacity * this.legendDimOpacity : baseOpacity)
                .attr("stroke", isSingle
                    ? this.settings.pointSettings.pointSettingsGroup.pointBorderColor.value.value
                    : clusterBorderColor)
                .attr("stroke-width", clusterBorderThickness);

            // Solid inner bubble showing the selected category's share of this cluster
            if (highlightActive && matchCount > 0) {
                const highlightSize = Math.min(clusterSize, radiusFor(matchCount));
                highlightCircle
                    .style("display", null)
                    .attr("cx", d.centroid[0])
                    .attr("cy", d.centroid[1])
                    .attr("data-base-r", highlightSize)
                    .attr("data-base-stroke", clusterBorderThickness)
                    .attr("r", highlightSize)
                    .attr("fill", matchCount === 1 && useLegendColors && d.legendCounts.size === 1
                        ? this.getLegendColor(Array.from(d.legendCounts.keys())[0] as string)
                        : highlightFill)
                    .attr("data-base-opacity", baseOpacity)
                    .attr("opacity", baseOpacity)
                    .attr("stroke", clusterBorderColor)
                    .attr("stroke-width", clusterBorderThickness);
            } else {
                highlightCircle.style("display", "none");
            }

            const labelCount = highlightActive ? matchCount : d.count;
            if (showCount && labelCount > 1) {
                text
                    .style("display", null)
                    .attr("x", d.centroid[0])
                    .attr("y", d.centroid[1])
                    .attr("text-anchor", "middle")
                    .attr("dy", "0.35em")
                    .attr("fill", countColor)
                    .attr("data-base-font", countSize)
                    .attr("font-size", countSize)
                    .attr("font-weight", "bold")
                    .text(labelCount.toString());
            } else {
                text.style("display", "none").text("");
            }
        });

        merged
            .on("mouseover", (event, d) => {
                if (!this.settings.tooltipSettings.tooltipSettingsGroup.showPointTooltip.value) {
                    return;
                }

                const tooltipData: VisualTooltipDataItem[] = [];
                if (d.count > 1) {
                    tooltipData.push({
                        displayName: "Cluster",
                        value: `${d.count} punten`
                    });
                } else if (d.legend && this.settings.tooltipSettings.tooltipSettingsGroup.showPointLegend.value) {
                    tooltipData.push({
                        displayName: this.clusterLegendFieldName || "Categorie",
                        value: d.legend
                    });
                }

                if (this.isLegendHighlightActive()) {
                    this.selectedLegendCategories.forEach(category => {
                        tooltipData.push({
                            displayName: category,
                            value: `${d.legendCounts.get(category) || 0} punten`
                        });
                    });
                }

                if (this.settings.tooltipSettings.tooltipSettingsGroup.showPointCoordinates.value) {
                    tooltipData.push({
                        displayName: "Centrum X",
                        value: d.centroid[0].toFixed(1)
                    });
                    tooltipData.push({
                        displayName: "Centrum Y",
                        value: d.centroid[1].toFixed(1)
                    });
                }

                if (tooltipData.length > 0) {
                    this.visualHostTooltipService.show({
                        coordinates: [event.pageX, event.pageY],
                        isTouchEvent: false,
                        dataItems: tooltipData,
                        identities: []
                    });
                }
            })
            .on("mousemove", (event) => {
                if (this.settings.tooltipSettings.tooltipSettingsGroup.showPointTooltip.value) {
                    this.visualHostTooltipService.move({
                        coordinates: [event.pageX, event.pageY],
                        isTouchEvent: false,
                        identities: []
                    });
                }
            })
            .on("mouseout", () => {
                if (this.settings.tooltipSettings.tooltipSettingsGroup.showPointTooltip.value) {
                    this.visualHostTooltipService.hide({
                        immediately: true,
                        isTouchEvent: false,
                    });
                }
            })
            .on("click", (event, d: ClusterDatum) => {
                this.selectPointIdentities(event, d.points.map(point => point.selectionId));
            });

        clusters.exit().remove();
        this.updatePointScaling(transform);
    }

    private createExtraFeature(name: string, code: string, index: number, width: number, height: number): GeoJsonFeature {
        const zoomScale = d3.zoomTransform(this.svg.node()).k;

        const extraFeatureWidth = 60 / zoomScale; // width * 0.08;
        const extraFeatureHeight = 30 / zoomScale; // height * 0.04;
        const ySpacing = 60 / zoomScale; // height * 0.02;
    
        const xOffset = width - extraFeatureWidth - 2 / zoomScale; // 10 pixels from the right edge
        const yOffset = height - ((index + 1) * (extraFeatureHeight + ySpacing)) - 2 / zoomScale; // 5 pixels from the bottom
    
        return {
            type: "Feature",
            properties: {
                naam: name,
                code: code,
                rubriek: "custom"
            },
            geometry: {
                type: "Polygon",
                coordinates: [[
                    [xOffset, yOffset],
                    [xOffset + extraFeatureWidth, yOffset],
                    [xOffset + extraFeatureWidth, yOffset + extraFeatureHeight],
                    [xOffset, yOffset + extraFeatureHeight],
                    [xOffset, yOffset]
                ]]
            },
            value: 0,
            customLabel: name,
            tooltipData: [{displayName: "Name", value: name}],
            selectionId: null
        };
    }
    
    
    
    /**
     * Returns properties pane formatting model content hierarchies, properties and latest formatting values.
     */
    public getFormattingModel(): powerbi.visuals.FormattingModel {
        const formattingModel = this.formattingSettingsService.buildFormattingModel(this.settings);
        
        // Apply conditional formatting: disable settings based on cockpit toggle states
        try {
            const modelAny = formattingModel as any;
            if (modelAny && modelAny.cards) {
                // Find cockpit card
                const cockpitCard = modelAny.cards.find((card: any) => 
                    card.displayName === "Cockpit instellingen" || card.name === "cockpitSettingsGroup"
                );
                
                // Find label settings card
                const labelCard = modelAny.cards.find((card: any) => 
                    card.displayName === "Labelinstellingen" || card.name === "labelSettingsGroup"
                );
                
                // Disable region names toggle in cockpit when labels toggle is off
                if (cockpitCard && cockpitCard.groups && cockpitCard.groups.length > 0) {
                    const cockpitGroup = cockpitCard.groups[0];
                    if (cockpitGroup && cockpitGroup.slices) {
                        const regionNamesSlice = cockpitGroup.slices.find((slice: any) => 
                            slice.displayName === "Regionamen toggle"
                        );
                        
                        if (regionNamesSlice) {
                            const labelsToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showLabelsToggle.value;
                            regionNamesSlice.enabled = labelsToggleEnabled;
                        }
                    }
                }
                
                // Disable main label settings when corresponding cockpit toggles are disabled
                if (labelCard && labelCard.groups && labelCard.groups.length > 0) {
                    const labelGroup = labelCard.groups[0];
                    if (labelGroup && labelGroup.slices) {
                        // Disable "Labels weergeven" when labels toggle is disabled in cockpit
                        const showLabelsSlice = labelGroup.slices.find((slice: any) => 
                            slice.displayName === "Labels weergeven"
                        );
                        if (showLabelsSlice) {
                            const labelsToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showLabelsToggle.value;
                            showLabelsSlice.enabled = labelsToggleEnabled;
                        }
                        
                        // Disable "Toon regiobenaming" when region names toggle is disabled in cockpit
                        const showRegionNameSlice = labelGroup.slices.find((slice: any) => 
                            slice.displayName === "Toon regiobenaming"
                        );
                        if (showRegionNameSlice) {
                            const regionNamesToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showRegionNamesToggle.value;
                            const labelsToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showLabelsToggle.value;
                            showRegionNameSlice.enabled = regionNamesToggleEnabled && labelsToggleEnabled;
                        }
                        
                        // Disable "Waardes weergeven" when values toggle is disabled in cockpit
                        const showValuesSlice = labelGroup.slices.find((slice: any) => 
                            slice.displayName === "Waardes weergeven"
                        );
                        if (showValuesSlice) {
                            const valuesToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showValuesToggle.value;
                            const labelsToggleEnabled = this.settings.cockpitSettings.cockpitSettingsGroup.showLabelsToggle.value;
                            showValuesSlice.enabled = valuesToggleEnabled && labelsToggleEnabled;
                        }
                    }
                }
            }
        } catch (error) {
            // If conditional formatting fails, just return the model as-is
            console.warn("Could not apply conditional formatting:", error);
        }

        // Add dynamic legend color pickers if legend categories exist (always show when legend field is populated)
        if (this.legendSlotOrder.length > 0) {
            try {
                // One colour picker per category, bound to the slot that category owns.
                // Every known category is listed, not just the ones currently in view, so
                // a filtered-out category keeps its colour and stays editable.
                const colorSlices: any[] = [];

                this.legendSlotOrder.forEach((category, slotIndex) => {
                    if (slotIndex >= this.maxLegendColorSlots) {
                        return;
                    }

                    const inView = this.legendCategories.indexOf(category) >= 0;

                    colorSlices.push({
                        uid: `legendColor_${category}_slice_${slotIndex}`,
                        displayName: inView ? category : `${category} (niet in selectie)`,
                        control: {
                            type: powerbi.visuals.FormattingComponent.ColorPicker,
                            properties: {
                                descriptor: {
                                    objectName: "legendColors",
                                    propertyName: `color${slotIndex + 1}`
                                },
                                value: {
                                    value: this.getLegendColor(category)
                                }
                            }
                        }
                    });
                });

                // Create point legend settings slices
                const pointLegendSettingsSlices: any[] = [
                    {
                        uid: "showPointLegend_slice",
                        displayName: "Puntlegenda weergeven",
                        control: {
                            type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                            properties: {
                                descriptor: {
                                    objectName: "pointSettings",
                                    propertyName: "showPointLegend"
                                },
                                value: this.settings.pointSettings.pointSettingsGroup.showPointLegend.value
                            }
                        }
                    },
                    {
                        uid: "pointLegendPosition_slice",
                        displayName: "Puntlegenda positie",
                        control: {
                            type: powerbi.visuals.FormattingComponent.Dropdown,
                            properties: {
                                descriptor: {
                                    objectName: "pointSettings",
                                    propertyName: "pointLegendPosition"
                                },
                                value: this.settings.pointSettings.pointSettingsGroup.pointLegendPosition.value
                            }
                        }
                    }
                ];

                // Create a single card with one group containing point legend settings and color pickers
                const legendColorsCard: any = {
                    displayName: "Puntlegenda",
                    uid: "legendColors_card",
                    groups: [
                        {
                            displayName: "Instellingen",
                            uid: "pointLegendSettings_group",
                            slices: pointLegendSettingsSlices
                        },
                        {
                            displayName: "Categorieën",
                            uid: "legendColors_group",
                            slices: colorSlices
                        }
                    ]
                };

                // Add the card to the formatting model
                (formattingModel as any).cards.push(legendColorsCard);
            } catch (error) {
                // Silently fail if unable to add legend colors
            }
        }
        
        return formattingModel;
    }

    private getLegendColorFromModel(legendValue: string): string | null {
        try {
            const legendColorsObj = this.dataView?.metadata?.objects?.["legendColors"];
            if (legendColorsObj && legendColorsObj[legendValue]) {
                const colorObj = legendColorsObj[legendValue] as any;
                
                // Try different possible structures
                if (colorObj) {
                    // Direct color value
                    if (typeof colorObj === 'string') {
                        return colorObj;
                    }
                    // { solid: { color: "#..." } }
                    if (colorObj.solid && colorObj.solid.color) {
                        return colorObj.solid.color;
                    }
                    // { value: "#..." }
                    if (colorObj.value) {
                        if (typeof colorObj.value === 'string') {
                            return colorObj.value;
                        }
                        if (colorObj.value.value) {
                            return colorObj.value.value;
                        }
                    }
                    // { fill: { solid: { color: "#..." } } }
                    if (colorObj.fill && colorObj.fill.solid && colorObj.fill.solid.color) {
                        return colorObj.fill.solid.color;
                    }
                }
            }
        } catch (error) {
            console.warn("Error getting legend color:", error);
        }
        return null;
    }

    private getDefaultColorForIndex(index: number): string {
        // Default color palette
        const colors = [
            "#FF6B6B", "#4ECDC4", "#45B7D1", "#FFA07A", "#98D8C8",
            "#F7DC6F", "#BB8FCE", "#85C1E2", "#F8B4D9", "#ABEBC6"
        ];
        return colors[index % colors.length];
    }


}