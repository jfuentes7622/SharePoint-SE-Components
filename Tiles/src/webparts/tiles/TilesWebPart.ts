import * as React from 'react';
import * as ReactDom from 'react-dom';
//import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
} from '@microsoft/sp-webpart-base';

import {
  PropertyFieldCustomList,
  CustomListFieldType
} from 'sp-client-custom-fields/lib/PropertyFieldCustomList';

import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

import * as strings from 'TilesWebPartStrings';
import { ITilesProps } from './components/ITilesProps';
import { ITileInfo } from './ITileInfo';
import { Tiles } from './components/Tiles';

const packageSolutionConfig: any = require('../../../config/package-solution.json');

export type TileShape = 'round' | 'rounded' | 'squared';
export type TileFontStyle = 'normal' | 'italic' | 'oblique';
export type TileHoverTransition = 'solid' | 'fade' | 'slide-top' | 'slide-bottom' | 'slide-left' | 'slide-right';

export interface ITilesWebPartProps {
  collectionData: ITileInfo[];
  title: string;

  // --- Appearance ---
  tileShape: TileShape;
  backgroundColor: string;
  textColor: string;
  hoverColor: string;
  hoverSameAsBackground: boolean;
  hoverTransition: TileHoverTransition;

  // --- Typography ---
  fontFamily: string;
  fontStyle: TileFontStyle;
  fontBold: boolean;
  fontSize: number;

  // --- Size ---
  tileWidth: number;
  tileHeight: number;

  // legacy
  tileEffect: string;

  enableDiagnostics: boolean;
}

export default class TilesWebPart extends BaseClientSideWebPart<ITilesWebPartProps> {

  public render(): void {
    this.logDiagnostic('render() called. tile count=' + String(this.properties.collectionData ? this.properties.collectionData.length : 0));
    const element: React.ReactElement<ITilesProps> = React.createElement(
      Tiles,
      {
        title: this.properties.title,
        collectionData: this.properties.collectionData,
        displayMode: this.displayMode,
        // appearance
        tileShape: this.properties.tileShape || 'rounded',
        backgroundColor: this.properties.backgroundColor || '#8A1717',
        textColor: this.properties.textColor || '#FAFAFA',
        hoverColor: this.properties.hoverColor || '#6a1010',
        hoverSameAsBackground: this.properties.hoverSameAsBackground !== false,
        hoverTransition: this.properties.hoverTransition || 'fade',
        // typography
        fontFamily: this.properties.fontFamily || 'inherit',
        fontStyle: this.properties.fontStyle || 'normal',
        fontBold: this.properties.fontBold,
        fontSize: this.properties.fontSize || 14,
        // size
        tileWidth: this.properties.tileWidth || 140,
        tileHeight: this.properties.tileHeight || 140,
        // legacy
        tileEffect: this.properties.tileEffect,
        enableDiagnostics: this.properties.enableDiagnostics !== false,
        fUpdateProperty: (value: string) => {
          this.properties.title = value;
        },
        fPropertyPaneOpen: this.context.propertyPane.open
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this.logDiagnostic('onInit started. tile count=' + String(this.properties.collectionData ? this.properties.collectionData.length : 0));
    return Promise.resolve();
  }

  protected onDispose(): void {
    this.logDiagnostic('onDispose called.');
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this.logDiagnostic('Property changed: ' + propertyPath + ', old=' + String(oldValue) + ', new=' + String(newValue));
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'hoverSameAsBackground') {
      this.context.propertyPane.refresh();
    }
  }

  private logDiagnostic(message: string): void {
    if (this.properties.enableDiagnostics === false) {
      return;
    }

    console.log('[TilesWebPart] ' + message);
  }

  private getWebPartVersion(): string {
    const solutionVersion = packageSolutionConfig && packageSolutionConfig.solution
      ? String(packageSolutionConfig.solution.version || '')
      : '';
    if (solutionVersion) {
      return solutionVersion;
    }

    const manifestVersion = this.context && this.context.manifest ? String(this.context.manifest.version || '') : '';
    if (manifestVersion && manifestVersion !== '*') {
      return manifestVersion;
    }

    return 'Unknown';
  }


  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const fontFamilyOptions: IPropertyPaneDropdownOption[] = [
      { key: 'inherit', text: strings.fontFamilyInherit },
      { key: 'Segoe UI, sans-serif', text: 'Segoe UI' },
      { key: 'Arial, sans-serif', text: 'Arial' },
      { key: 'Calibri, sans-serif', text: 'Calibri' },
      { key: 'Georgia, serif', text: 'Georgia' },
      { key: 'Tahoma, sans-serif', text: 'Tahoma' },
      { key: 'Trebuchet MS, sans-serif', text: 'Trebuchet MS' },
      { key: 'Verdana, sans-serif', text: 'Verdana' }
    ];
    const fontStyleOptions: IPropertyPaneDropdownOption[] = [
      { key: 'normal', text: strings.fontStyleNormal },
      { key: 'italic', text: strings.fontStyleItalic },
      { key: 'oblique', text: strings.fontStyleOblique }
    ];
    const tileShapeOptions: IPropertyPaneDropdownOption[] = [
      { key: 'squared', text: strings.tileShapeSquared },
      { key: 'rounded', text: strings.tileShapeRounded },
      { key: 'round', text: strings.tileShapeRound }
    ];
    const hoverTransitionOptions: IPropertyPaneDropdownOption[] = [
      { key: 'solid', text: strings.hoverTransitionSolid },
      { key: 'fade', text: strings.hoverTransitionFade },
      { key: 'slide-top', text: strings.hoverTransitionSlideTop },
      { key: 'slide-bottom', text: strings.hoverTransitionSlideBottom },
      { key: 'slide-left', text: strings.hoverTransitionSlideLeft },
      { key: 'slide-right', text: strings.hoverTransitionSlideRight }
    ];

    const fullVersionLabel = 'Version: ' + this.getWebPartVersion();

    return {
      pages: [
        {
          header: {
            description: ''
          },
          groups: [
            {
              groupName: fullVersionLabel,
              groupFields: [
                PropertyPaneLabel('propertyPaneVersionInfo', {
                  text: ' '
                })
              ]
            },
            // ── Tile Data ────────────────────────────────────────────
            {
              groupName: strings.tileDataGroup,
              groupFields: [
                PropertyFieldCustomList('collectionData', {
                  key: 'collectionData',
                  label: strings.tilesDataLabel,
                  headerText: strings.tilesPanelHeader,
                  value: this.properties.collectionData || [],
                  context: this.context,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  render: this.render.bind(this),
                  properties: this.properties,
                  fields: [
                    {
                      id: 'title',
                      title: strings.titleField,
                      type: CustomListFieldType.string,
                      required: true
                    },
                    {
                      id: 'description',
                      title: strings.descriptionField,
                      type: CustomListFieldType.string
                    },
                    {
                      id: 'url',
                      title: strings.urlField,
                      type: CustomListFieldType.string,
                      required: true
                    },
                    {
                      id: 'imageUrl',
                      title: strings.imageUrlField,
                      type: CustomListFieldType.string
                    },
                    {
                      id: 'imageOnly',
                      title: strings.imageOnlyField,
                      type: CustomListFieldType.boolean
                    },
                    {
                      id: 'imagePosition',
                      title: strings.imagePositionField,
                      type: CustomListFieldType.string
                    },
                    {
                      id: 'hoverImageUrl',
                      title: strings.hoverImageUrlField,
                      type: CustomListFieldType.string
                    },
                    {
                      id: 'hoverImageOnly',
                      title: strings.hoverImageOnlyField,
                      type: CustomListFieldType.boolean
                    },
                    {
                      id: 'hoverImagePosition',
                      title: strings.hoverImagePositionField,
                      type: CustomListFieldType.string
                    },
                    {
                      id: 'target',
                      title: strings.targetField,
                      type: CustomListFieldType.string,
                      required: true
                    },
                    {
                      id: 'color',
                      title: strings.tileColorOverrideField,
                      type: CustomListFieldType.color
                    }
                  ],
                  disableReactivePropertyChanges: false
                })
              ]
            },
            // ── Appearance ───────────────────────────────────────────
            {
              groupName: strings.appearanceGroup,
              groupFields: [
                PropertyPaneDropdown('tileShape', {
                  label: strings.tileShapeLabel,
                  options: tileShapeOptions,
                  selectedKey: this.properties.tileShape || 'rounded'
                }),
                PropertyFieldColorPicker('backgroundColor', {
                  label: strings.backgroundColorLabel,
                  selectedColor: this.properties.backgroundColor || '#8A1717',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'backgroundColorPicker'
                }),
                PropertyFieldColorPicker('textColor', {
                  label: strings.textColorLabel,
                  selectedColor: this.properties.textColor || '#FAFAFA',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'textColorPicker'
                }),
                PropertyPaneToggle('hoverSameAsBackground', {
                  label: strings.hoverSameAsBackgroundLabel,
                  checked: this.properties.hoverSameAsBackground !== false
                })
              ].concat(this.properties.hoverSameAsBackground === false ? [
                PropertyFieldColorPicker('hoverColor', {
                  label: strings.hoverColorLabel,
                  selectedColor: this.properties.hoverColor || '#6a1010',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'hoverColorPicker'
                })
              ] : []).concat([
                PropertyPaneDropdown('hoverTransition', {
                  label: strings.hoverTransitionLabel,
                  options: hoverTransitionOptions,
                  selectedKey: this.properties.hoverTransition || 'fade'
                })
              ])
            },
            // ── Typography ───────────────────────────────────────────
            {
              groupName: strings.typographyGroup,
              groupFields: [
                PropertyPaneDropdown('fontFamily', {
                  label: strings.fontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.fontFamily || 'inherit'
                }),
                PropertyPaneDropdown('fontStyle', {
                  label: strings.fontStyleLabel,
                  options: fontStyleOptions,
                  selectedKey: this.properties.fontStyle || 'normal'
                }),
                PropertyPaneToggle('fontBold', {
                  label: strings.fontBoldLabel,
                  checked: !!this.properties.fontBold
                }),
                PropertyPaneSlider('fontSize', {
                  label: strings.fontSizeLabel,
                  min: 10,
                  max: 28,
                  step: 1,
                  value: this.properties.fontSize || 14
                })
              ]
            },
            // ── Size ─────────────────────────────────────────────────
            {
              groupName: strings.sizeGroup,
              groupFields: [
                PropertyPaneSlider('tileWidth', {
                  label: strings.tileWidthLabel,
                  min: 80,
                  max: 400,
                  step: 4,
                  value: this.properties.tileWidth || 140
                }),
                PropertyPaneSlider('tileHeight', {
                  label: strings.tileHeightLabel,
                  min: 80,
                  max: 400,
                  step: 4,
                  value: this.properties.tileHeight || 140
                })
              ]
            },
            // ── Diagnostics ──────────────────────────────────────────
            {
              groupName: strings.diagnosticsGroup,
              groupFields: [
                PropertyPaneCheckbox('enableDiagnostics', {
                  text: strings.enableDiagnosticsLabel,
                  checked: this.properties.enableDiagnostics !== false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

