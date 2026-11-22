/*
 *  Power BI Visualizations
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

import { dataViewObjectsParser } from "powerbi-visuals-utils-dataviewutils";
import DataViewObjectsParser = dataViewObjectsParser.DataViewObjectsParser;
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import { dataViewWildcard } from "powerbi-visuals-utils-dataviewutils";
import powerbi from "powerbi-visuals-api";
import ISQExpr = powerbi.data.ISQExpr;

import FormattingSettingsCompositeCard = formattingSettings.SimpleCard;
import FormattingSettingsGroup = formattingSettings.Group;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

export class VisualSettings extends formattingSettings.Model {
  // public labelSettings: LabelSettings = new LabelSettings();
  // Create formatting settings model formatting cards
  public mapSettings : mapSettingsCard = new mapSettingsCard();
  public labelSettings : labelSettingsCard = new labelSettingsCard();
  public tooltipSettings : tooltipSettingsCard = new tooltipSettingsCard();
  public cockpitSettings : cockpitSettingsCard = new cockpitSettingsCard();
  public legendSettings : legendSettingsCard = new legendSettingsCard();
  public pointSettings : pointSettingsCard = new pointSettingsCard();
  // labelSettingsCard : FormattingSettingsCompositeCard = new LabelSettings();
  public cards: Array<formattingSettings.SimpleCard> = [this.mapSettings, this.labelSettings, this.tooltipSettings, this.cockpitSettings, this.legendSettings, this.pointSettings];

}


/**
 * ColorSettings Formatting Card #1
 */

const mapOptions : powerbi.IEnumMember[] = [
  { displayName: "Provincies", value: "provincies" },
  { displayName: "Wijken", value: "wijken" },
  { displayName: "Gemeenten 2023", value: "gemeenten 2023" },
  { displayName: "Gemeenten 2021", value: "gemeenten 2021" },
  { displayName: "Gemeenten 2020", value: "gemeenten 2020" },
  { displayName: "Gemeenten 2019", value: "gemeenten 2019" },
  { displayName: "Gemeenten 2016", value: "gemeenten 2016" },
  { displayName: "Gemeenten 2012", value: "gemeenten 2012" },
  { displayName: "GGD Regios", value: "ggdregios" },
  // { displayName: "Werkgebieden", value: "werkgebieden" },
  { displayName: "Veiligheidsregios", value: "veiligheidsregios" },
  { displayName: "Politieregios", value: "politieregios" }
]

class mapColorSettingsGroup extends formattingSettings.SimpleCard {
  public minColorSlice: formattingSettings.ColorPicker = new formattingSettings.ColorPicker({
      name: "minColorSlice",
      displayName: "Min kleur",
      value: { value: "#000000" }
  });

  public minPosition: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
      name: "minPosition",
      displayName: "Min positie",
      value: 0
  });

  public useMidColorSlice: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "useMidColorSlice",
      displayName: "Gebruik middenkleur",
      value: true // Default value
  });

  public midColorSlice: formattingSettings.ColorPicker = new formattingSettings.ColorPicker({
      name: "midColorSlice",
      displayName: "Midden kleur",
      value: { value: "#FFFFFF" } // Default color
  });

  public midPosition: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
      name: "midPosition",
      displayName: "Midden positie",
      value: 50 // Default position
  });

  public maxColorSlice: formattingSettings.ColorPicker = new formattingSettings.ColorPicker({
      name: "maxColorSlice",
      displayName: "Max kleur",
      value: { value: "#FF0000" } // Default color
  });

  public maxPosition: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
      name: "maxPosition",
      displayName: "Max positie",
      value: 100
  });

  public useAbsoluteValues: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "useAbsoluteValues",
      displayName: "Gebruik absolute waarden",
      value: false
  });

  name: string = "mapColorSettingsGroup";
  displayName: string = "Kaartkleuren";
  slices: formattingSettings.Slice[] = [
      this.minColorSlice, this.minPosition,
      this.useMidColorSlice, this.midColorSlice, this.midPosition,
      this.maxColorSlice, this.maxPosition,
      this.useAbsoluteValues
  ];
}
class mapSelectSettingsGroup extends formattingSettings.SimpleCard {
  public selectedMapSlice: formattingSettings.ItemDropdown = new formattingSettings.ItemDropdown({
      name: "selectedMapSlice",
      displayName: "Geselecteerde kaart",
      items: mapOptions,
      value: mapOptions[0]
  });
  /* public customTopojson: formattingSettings.TextArea = new formattingSettings.TextArea({
      name: "customTopojson",
      displayName: "Aangepaste topojson",
      value: "",
      placeholder: "Voer de jsontekst van je kaart in"
  });
  public customNameFieldName: formattingSettings.TextInput = new formattingSettings.TextInput({
      name: "customNameFieldName",
      displayName: "Aangepaste naam veld",
      value: "",
      placeholder: "Voer de naam van het naamveld in"
  });
  public customCodeFieldName: formattingSettings.TextInput = new formattingSettings.TextInput({
      name: "customCodeFieldName",
      displayName: "Aangepaste code veld",
      value: "",
      placeholder: "Voer de naam van het codeveld in"
  });*/
  public showExtraFeatures: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showExtraFeatures",
      displayName: "Toon extra features",
      value: true
  });
  name: string = "mapSelectSettingsGroup";
  displayName: string = "Kaartselectie";
  slices: formattingSettings.Slice[] = [this.selectedMapSlice, this.showExtraFeatures];
}

class cockpitSettingsGroup extends formattingSettings.SimpleCard {
  public enableCockpit: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "enableCockpit",
      displayName: "Cockpit inschakelen",
      value: false
  });

  public showLabelsToggle: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showLabelsToggle",
      displayName: "Labels toggle",
      value: true
  });
  public showRegionNamesToggle: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showRegionNamesToggle",
      displayName: "Regionamen toggle",
      value: true
  });
    public showValuesToggle: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showValuesToggle",
      displayName: "Waarden toggle",
      value: true
  });

  public showLegendToggle: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showLegendToggle",
      displayName: "Legenda toggle",
      value: true
  });

  public showPointLegendToggle: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showPointLegendToggle",
      displayName: "Puntlegenda toggle",
      value: true
  });

  public startCockpitCollapsed: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "startCockpitCollapsed",
      displayName: "Start ingeklapt",
      value: false
  });

  name: string = "cockpitSettingsGroup";
  displayName: string = "Cockpit instellingen";
  slices: formattingSettings.Slice[] = [
      this.enableCockpit,
      this.showLabelsToggle,
      this.showRegionNamesToggle,
      this.showValuesToggle,
      this.showLegendToggle,
      this.showPointLegendToggle,
      this.startCockpitCollapsed
  ];
}

class mapSettingsCard extends formattingSettings.CompositeCard {
  public mapSelectSettingsGroup: mapSelectSettingsGroup = new mapSelectSettingsGroup();
  public mapColorSettingsGroup: mapColorSettingsGroup = new mapColorSettingsGroup();
  public regionBorderSettingsGroup: regionBorderSettingsGroup = new regionBorderSettingsGroup();
  name: string = "mapSettings";
  displayName: string = "Kaartinstellingen";
  groups: Array<formattingSettings.Group> = [this.mapSelectSettingsGroup, this.mapColorSettingsGroup, this.regionBorderSettingsGroup];
}

class labelSettingsCard extends formattingSettings.SimpleCard {
  public labelSettingsGroup: labelSettingsGroup = new labelSettingsGroup();
  name: string = "labelSettings";
  displayName: string = "Labelinstellingen";
  slices: formattingSettings.Slice[] = [
    this.labelSettingsGroup.showLabels,
    this.labelSettingsGroup.showLabelBorder,
    this.labelSettingsGroup.labelBorderColor,
    this.labelSettingsGroup.labelBorderThickness,
    this.labelSettingsGroup.showValues,
    this.labelSettingsGroup.useCustomLabel,
    this.labelSettingsGroup.showRegionName,
    this.labelSettingsGroup.fontSize,
    this.labelSettingsGroup.numOfDigits,
    this.labelSettingsGroup.showLabelBackground,
    this.labelSettingsGroup.labelBackgroundColor,
    this.labelSettingsGroup.labelBackgroundOpacity
  ];
}

class tooltipSettingsCard extends formattingSettings.SimpleCard {
  public tooltipSettingsGroup: tooltipSettingsGroup = new tooltipSettingsGroup();
  name: string = "tooltipSettings";
  displayName: string = "Tooltipinstellingen";
  slices: formattingSettings.Slice[] = [
    this.tooltipSettingsGroup.showTooltip,
    this.tooltipSettingsGroup.showPointTooltip,
    this.tooltipSettingsGroup.showPointCategory,
    this.tooltipSettingsGroup.showPointCoordinates,
    this.tooltipSettingsGroup.showPointLegend
  ];
}

class cockpitSettingsCard extends formattingSettings.SimpleCard {
  public cockpitSettingsGroup: cockpitSettingsGroup = new cockpitSettingsGroup();
  name: string = "cockpitSettings";
  displayName: string = "Cockpitinstellingen";
  slices: formattingSettings.Slice[] = [
    this.cockpitSettingsGroup.enableCockpit,
    this.cockpitSettingsGroup.showLabelsToggle,
    this.cockpitSettingsGroup.showRegionNamesToggle,
    this.cockpitSettingsGroup.showValuesToggle,
    this.cockpitSettingsGroup.showLegendToggle,
    this.cockpitSettingsGroup.showPointLegendToggle,
    this.cockpitSettingsGroup.startCockpitCollapsed
  ];
}

class legendSettingsGroup extends formattingSettings.SimpleCard {
  public showLegend: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showLegend",
      displayName: "Legenda weergeven",
      value: true
  });

  public legendPosition: formattingSettings.ItemDropdown = new formattingSettings.ItemDropdown({
      name: "legendPosition",
      displayName: "Legenda positie",
      items: [
          { displayName: "Links onder", value: "bottom-left" },
          { displayName: "Rechts onder", value: "bottom-right" },
          { displayName: "Links boven", value: "top-left" },
          { displayName: "Rechts boven", value: "top-right" }
      ],
      value: { displayName: "Links onder", value: "bottom-left" }
  });

  name: string = "legendSettingsGroup";
  displayName: string = "Legenda instellingen";
  slices: formattingSettings.Slice[] = [
      this.showLegend,
      this.legendPosition
  ];
}


class legendSettingsCard extends formattingSettings.SimpleCard {
  public legendSettingsGroup: legendSettingsGroup = new legendSettingsGroup();
  name: string = "legendSettings";
  displayName: string = "Legenda instellingen";
  slices: formattingSettings.Slice[] = [
    this.legendSettingsGroup.showLegend,
    this.legendSettingsGroup.legendPosition
  ];
}

class labelSettingsGroup extends formattingSettings.SimpleCard {
  public showLabels: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showLabels",
      displayName: "Labels weergeven",
      value: true
  });
  public showLabelBorder: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showLabelBorder",
      displayName: "Labelrand",
      value: true
  });
  public labelBorderColor: formattingSettings.ColorPicker = new formattingSettings.ColorPicker({
      name: "labelBorderColor",
      displayName: "Randkleur",
      value: { value: "#000000" }
  });
  public labelBorderThickness: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
      name: "labelBorderThickness",
      displayName: "Randdikte",
      value: 1
  });
  public showValues: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showValues",
      displayName: "Waardes weergeven",
      value: true
  });
  public useCustomLabel: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "useCustomLabel",
      displayName: "Gebruik aangepast label",
      value: false
  });
  
  public showRegionName: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showRegionName",
      displayName: "Toon regiobenaming",
      value: true
  });

  public fontSize: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
      name: "fontSize",
      displayName: "Lettergrootte",
      value: 12
  });

  public numOfDigits: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
      name: "numOfDigits",
      displayName: "Aantal decimalen",
      value: 1
  });

  public showLabelBackground: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showLabelBackground",
      displayName: "Toon achtergrond",
      value: true
  });

  public labelBackgroundColor: formattingSettings.ColorPicker = new formattingSettings.ColorPicker({
      name: "labelBackgroundColor",
      displayName: "Achtergrondkleur",
      value: { value: "#ffffff" }
  });

  // Deprecated in favor of labelBorderColor/labelBorderThickness

  public labelBackgroundOpacity: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
      name: "labelBackgroundOpacity",
      displayName: "Doorzichtigheid",
      value: 100
  });

name: string = "labelSettingsGroup";
  displayName: string = "Labelinstellingen";
  //groupLevelSlice: formattingSettings.ToggleSwitch[] = [this.showLabels]
  //topLevelSlices: formattingSettings.ToggleSwitch[] = [this.showLabels]
  slices: formattingSettings.Slice[] = [
      this.showLabels,
      this.showLabelBorder,
      this.labelBorderColor,
      this.labelBorderThickness,
      this.showValues,
      this.useCustomLabel,
      this.showRegionName,
      this.fontSize,
      this.numOfDigits,
      this.showLabelBackground,
      this.labelBackgroundColor,
      this.labelBackgroundOpacity
  ];
}

class tooltipSettingsGroup extends formattingSettings.SimpleCard {
  public showTooltip: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showTooltip",
      displayName: "Regio tooltip weergeven",
      value: true
  });

  public showPointTooltip: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showPointTooltip",
      displayName: "Punt tooltip weergeven",
      value: true
  });

  public showPointCategory: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showPointCategory",
      displayName: "Toon punt categorie",
      value: true
  });

  public showPointCoordinates: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showPointCoordinates",
      displayName: "Toon coördinaten",
      value: false
  });

  public showPointLegend: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showPointLegend",
      displayName: "Toon punt legenda",
      value: true
  });

  name: string = "tooltipSettingsGroup";
  displayName: string = "Tooltipinstellingen";
  slices: formattingSettings.Slice[] = [
    this.showTooltip,
    this.showPointTooltip,
    this.showPointCategory,
    this.showPointCoordinates,
    this.showPointLegend
  ];
}

class regionBorderSettingsGroup extends formattingSettings.SimpleCard {
  public showRegionBorder: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showRegionBorder",
      displayName: "Toon rand",
      value: true
  });

  public regionBorderColor: formattingSettings.ColorPicker = new formattingSettings.ColorPicker({
      name: "regionBorderColor",
      displayName: "Randkleur",
      value: { value: "#000000" }
  });

  public regionBorderThickness: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
      name: "regionBorderThickness",
      displayName: "Randdikte",
      value: 5
  });

  name: string = "borderSettingsGroup";
  displayName: string = "Randinstellingen";
  slices: formattingSettings.Slice[] = [
      this.showRegionBorder,
      this.regionBorderColor,
      this.regionBorderThickness
  ];
}

class pointSettingsGroup extends formattingSettings.SimpleCard {
  public showPoints: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showPoints",
      displayName: "Toon punten",
      value: true
  });

  public pointColor: formattingSettings.ColorPicker = new formattingSettings.ColorPicker({
      name: "pointColor",
      displayName: "Puntkleur",
      value: { value: "#FF0000" }
  });

  public pointSize: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
      name: "pointSize",
      displayName: "Puntgrootte",
      value: 3
  });

  public pointBorderColor: formattingSettings.ColorPicker = new formattingSettings.ColorPicker({
      name: "pointBorderColor",
      displayName: "Puntrandkleur",
      value: { value: "#FFFFFF" }
  });

  public pointBorderThickness: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
      name: "pointBorderThickness",
      displayName: "Puntranddikte",
      value: 1
  });

  public pointOpacity: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
      name: "pointOpacity",
      displayName: "Doorzichtigheid",
      value: 60
  });

  public useLegendColors: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "useLegendColors",
      displayName: "Gebruik legenda",
      value: false
  });

  public showPointLegend: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
      name: "showPointLegend",
      displayName: "Puntlegenda weergeven",
      value: true
  });

  public pointLegendPosition: formattingSettings.ItemDropdown = new formattingSettings.ItemDropdown({
      name: "pointLegendPosition",
      displayName: "Puntlegenda positie",
      items: [
          { displayName: "Links onder", value: "bottom-left" },
          { displayName: "Rechts onder", value: "bottom-right" },
          { displayName: "Links boven", value: "top-left" },
          { displayName: "Rechts boven", value: "top-right" }
      ],
      value: { displayName: "Rechts onder", value: "bottom-right" }
  });

  name: string = "pointSettingsGroup";
  displayName: string = "Puntinstellingen";
  slices: formattingSettings.Slice[] = [
      this.showPoints,
      this.pointColor,
      this.pointSize,
      this.pointBorderColor,
      this.pointBorderThickness,
      this.pointOpacity,
      this.useLegendColors,
      this.showPointLegend,
      this.pointLegendPosition
  ];
}

class clusterSettingsGroup extends formattingSettings.SimpleCard {
  public enableClustering: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
    name: "enableClustering",
        displayName: "Cluster punten",
        value: false
    });

    public clusterColor: formattingSettings.ColorPicker = new formattingSettings.ColorPicker({
        name: "clusterColor",
        displayName: "Cluster kleur",
        value: { value: "#FF6B6B" }
    });

    public clusterBorderColor: formattingSettings.ColorPicker = new formattingSettings.ColorPicker({
        name: "clusterBorderColor",
        displayName: "Clusterrandkleur",
        value: { value: "#FFFFFF" }
    });

    public clusterOpacity: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
        name: "clusterOpacity",
        displayName: "Cluster doorzichtigheid",
        value: 80
    });

    public clusterBorderThickness: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
        name: "clusterBorderThickness",
        displayName: "Clusterranddikte",
        value: 1
    });

    public minClusterSize: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
        name: "minClusterSize",
        displayName: "Min cluster grootte",
        value: 10
    });

    public maxClusterSize: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
        name: "maxClusterSize",
        displayName: "Max cluster grootte",
        value: 50
    });

    public showClusterCount: formattingSettings.ToggleSwitch = new formattingSettings.ToggleSwitch({
        name: "showClusterCount",
        displayName: "Toon cluster aantal",
        value: true
    });

    public clusterCountColor: formattingSettings.ColorPicker = new formattingSettings.ColorPicker({
        name: "clusterCountColor",
        displayName: "Aantal tekst kleur",
        value: { value: "#FFFFFF" }
    });

    public clusterCountSize: formattingSettings.NumUpDown = new formattingSettings.NumUpDown({
        name: "clusterCountSize",
        displayName: "Aantal tekst grootte",
        value: 12
    });

  name: string = "clusterSettingsGroup";
  displayName: string = "Clusterinstellingen";
  slices: formattingSettings.Slice[] = [
    this.enableClustering,
    this.clusterColor,
    this.clusterBorderColor,
    this.clusterBorderThickness,
    this.clusterOpacity,
    this.minClusterSize,
    this.maxClusterSize,
    this.showClusterCount,
    this.clusterCountColor,
    this.clusterCountSize
  ];
}


class pointSettingsCard extends formattingSettings.CompositeCard {
  public pointSettingsGroup: pointSettingsGroup = new pointSettingsGroup();
  public clusterSettingsGroup: clusterSettingsGroup = new clusterSettingsGroup();
  name: string = "pointSettings";
  displayName: string = "Puntinstellingen";
  
  groups: Array<formattingSettings.Group> = [this.pointSettingsGroup, this.clusterSettingsGroup];
}