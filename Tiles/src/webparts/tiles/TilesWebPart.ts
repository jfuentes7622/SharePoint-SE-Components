import * as React from 'react';
import * as ReactDom from 'react-dom';
//import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
} from '@microsoft/sp-webpart-base';

import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

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
const picturePickerModule: any = require('sp-client-custom-fields/lib/PropertyFieldPicturePickerHost');
const PropertyFieldPicturePickerHost: any = picturePickerModule.default || picturePickerModule;

export type TileShape = 'round' | 'rounded' | 'squared';
export type TileFontStyle = 'normal' | 'italic' | 'oblique';
export type TileHoverTransition = 'solid' | 'fade' | 'slide-top' | 'slide-bottom' | 'slide-left' | 'slide-right';
export type TileLayout = 'stacked' | 'columns' | 'flow';

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
  tileLayout: TileLayout;
  columnCount: number;
  tileGap: number;

  // --- Web part surface ---
  webPartBackgroundColor: string;
  webPartBorderColor: string;
  webPartBorderWidth: number;
  webPartBorderStyle: string;
  webPartCornerRadius: number;
  webPartPadding: number;

  // --- Web part title ---
  showTitle: boolean;
  titleTextColor: string;
  titleFontFamily: string;
  titleFontSize: number;
  titleFontStyle: string;
  titleFontBold: boolean;
  titleAlignment: string;
  titleBottomSpacing: number;

  // legacy
  tileEffect: string;

  enableDiagnostics: boolean;
}

export default class TilesWebPart extends BaseClientSideWebPart<ITilesWebPartProps> {

  public constructor() {
    super();
    this.onPropertyPaneFieldChanged = this.onPropertyPaneFieldChanged.bind(this);
  }

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
        tileLayout: this.properties.tileLayout || 'columns',
        columnCount: this.properties.columnCount || 3,
        tileGap: typeof this.properties.tileGap === 'number' ? this.properties.tileGap : 16,
        webPartBackgroundColor: this.properties.webPartBackgroundColor || '#ffffff',
        webPartBorderColor: this.properties.webPartBorderColor || '#d2d0ce',
        webPartBorderWidth: typeof this.properties.webPartBorderWidth === 'number' ? this.properties.webPartBorderWidth : 0,
        webPartBorderStyle: this.properties.webPartBorderStyle || 'solid',
        webPartCornerRadius: typeof this.properties.webPartCornerRadius === 'number' ? this.properties.webPartCornerRadius : 0,
        webPartPadding: typeof this.properties.webPartPadding === 'number' ? this.properties.webPartPadding : 0,
        showTitle: this.properties.showTitle === true,
        titleTextColor: this.properties.titleTextColor || '#323130',
        titleFontFamily: this.properties.titleFontFamily || 'inherit',
        titleFontSize: typeof this.properties.titleFontSize === 'number' ? this.properties.titleFontSize : 20,
        titleFontStyle: this.properties.titleFontStyle || 'normal',
        titleFontBold: this.properties.titleFontBold !== false,
        titleAlignment: this.properties.titleAlignment || 'center',
        titleBottomSpacing: typeof this.properties.titleBottomSpacing === 'number' ? this.properties.titleBottomSpacing : 0,
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
    if (propertyPath === 'hoverSameAsBackground' || propertyPath === 'tileLayout') {
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

  private renderTileImagePicker(field: any, value: any,
    onUpdate: (fieldId: string, fieldValue: any) => void): React.ReactElement<any> {
    const properties: any = {};
    properties[field.id] = String(value || '');
    return React.createElement(PropertyFieldPicturePickerHost, {
      key: 'tileImagePicker-' + field.id + '-' + String(value || ''),
      targetProperty: field.id,
      label: '',
      initialValue: String(value || ''),
      context: this.context,
      previewImage: true,
      allowedFileExtensions: '.gif,.jpg,.jpeg,.bmp,.png,.svg,.webp',
      readOnly: false,
      disabled: false,
      properties: properties,
      disableReactivePropertyChanges: true,
      deferredValidationTime: 0,
      onGetErrorMessage: undefined,
      onPropertyChange: (propertyPath: string, oldValue: any, newValue: any): void => {
        onUpdate(field.id, String(newValue || ''));
      },
      render: (): void => undefined,
      onRender: undefined,
      onDispose: undefined
    });
  }

  private renderTileColorPicker(field: any, value: any,
    onUpdate: (fieldId: string, fieldValue: any) => void): React.ReactElement<any> {
    const colorValue = /^#[0-9a-f]{6}$/i.test(String(value || '')) ? String(value) : '#8A1717';
    return React.createElement('div', { style: { display: 'flex', alignItems: 'center' } },
      React.createElement('input', {
        type: 'color',
        value: colorValue,
        title: strings.tileColorOverrideField,
        style: { width: '40px', height: '32px', padding: '0', marginRight: '8px' },
        onChange: (event: React.FormEvent<HTMLInputElement>): void => {
          onUpdate(field.id, event.currentTarget.value);
        }
      }),
      React.createElement('input', {
        type: 'text',
        value: String(value || ''),
        placeholder: '#8A1717',
        style: { width: '90px' },
        onChange: (event: React.FormEvent<HTMLInputElement>): void => {
          onUpdate(field.id, event.currentTarget.value);
        }
      })
    );
  }

  private renderImageOpacity(field: any, value: any,
    onUpdate: (fieldId: string, fieldValue: any) => void): React.ReactElement<any> {
    const numericValue = Number(value);
    const opacity = isFinite(numericValue) ? Math.max(0, Math.min(100, numericValue)) : 100;
    return React.createElement('input', {
      type: 'number',
      min: 0,
      max: 100,
      step: 1,
      value: opacity,
      title: strings.imageOpacityField,
      style: { width: '80px' },
      onChange: (event: React.FormEvent<HTMLInputElement>): void => {
        const nextValue = Number(event.currentTarget.value);
        onUpdate(field.id, isFinite(nextValue) ? Math.max(0, Math.min(100, nextValue)) : 100);
      }
    });
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
              groupName: strings.webPartAppearanceGroup,
              groupFields: [
                PropertyFieldColorPicker('webPartBackgroundColor', {
                  label: strings.webPartBackgroundColorLabel,
                  selectedColor: this.properties.webPartBackgroundColor || '#ffffff',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'webPartBackgroundColorPicker'
                }),
                PropertyFieldColorPicker('webPartBorderColor', {
                  label: strings.webPartBorderColorLabel,
                  selectedColor: this.properties.webPartBorderColor || '#d2d0ce',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'webPartBorderColorPicker'
                }),
                PropertyPaneSlider('webPartBorderWidth', {
                  label: strings.webPartBorderWidthLabel,
                  min: 0,
                  max: 20,
                  step: 1,
                  value: typeof this.properties.webPartBorderWidth === 'number' ? this.properties.webPartBorderWidth : 0
                }),
                PropertyPaneDropdown('webPartBorderStyle', {
                  label: strings.webPartBorderStyleLabel,
                  options: [
                    { key: 'none', text: 'None' },
                    { key: 'solid', text: 'Solid' },
                    { key: 'dashed', text: 'Dashed' },
                    { key: 'dotted', text: 'Dotted' },
                    { key: 'double', text: 'Double' }
                  ],
                  selectedKey: this.properties.webPartBorderStyle || 'solid'
                }),
                PropertyPaneSlider('webPartCornerRadius', {
                  label: strings.webPartCornerRadiusLabel,
                  min: 0,
                  max: 40,
                  step: 1,
                  value: typeof this.properties.webPartCornerRadius === 'number' ? this.properties.webPartCornerRadius : 0
                }),
                PropertyPaneSlider('webPartPadding', {
                  label: strings.webPartPaddingLabel,
                  min: 0,
                  max: 100,
                  step: 1,
                  value: typeof this.properties.webPartPadding === 'number' ? this.properties.webPartPadding : 0
                })
              ]
            },
            {
              groupName: strings.titleAppearanceGroup,
              groupFields: [
                PropertyPaneCheckbox('showTitle', {
                  text: strings.showTitleLabel,
                  checked: this.properties.showTitle === true
                }),
                PropertyFieldColorPicker('titleTextColor', {
                  label: strings.titleTextColorLabel,
                  selectedColor: this.properties.titleTextColor || '#323130',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'titleTextColorPicker'
                }),
                PropertyPaneDropdown('titleFontFamily', {
                  label: strings.titleFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.titleFontFamily || 'inherit'
                }),
                PropertyPaneSlider('titleFontSize', {
                  label: strings.titleFontSizeLabel,
                  min: 10,
                  max: 72,
                  step: 1,
                  value: typeof this.properties.titleFontSize === 'number' ? this.properties.titleFontSize : 20
                }),
                PropertyPaneDropdown('titleFontStyle', {
                  label: strings.titleFontStyleLabel,
                  options: fontStyleOptions,
                  selectedKey: this.properties.titleFontStyle || 'normal'
                }),
                PropertyPaneCheckbox('titleFontBold', {
                  text: strings.titleFontBoldLabel,
                  checked: this.properties.titleFontBold !== false
                }),
                PropertyPaneDropdown('titleAlignment', {
                  label: strings.titleAlignmentLabel,
                  options: [
                    { key: 'left', text: strings.alignmentLeft },
                    { key: 'center', text: strings.alignmentCenter },
                    { key: 'right', text: strings.alignmentRight }
                  ],
                  selectedKey: this.properties.titleAlignment || 'center'
                }),
                PropertyPaneSlider('titleBottomSpacing', {
                  label: strings.titleBottomSpacingLabel,
                  min: 0,
                  max: 50,
                  step: 1,
                  value: typeof this.properties.titleBottomSpacing === 'number'
                    ? this.properties.titleBottomSpacing : 0
                })
              ]
            },
            {
              groupName: strings.tileDataGroup,
              groupFields: [
                PropertyFieldCollectionData('collectionData', {
                  key: 'collectionDataField',
                  label: strings.tilesDataLabel,
                  panelHeader: strings.tilesPanelHeader,
                  manageBtnLabel: strings.tilesManageBtn,
                  value: this.properties.collectionData || [],
                  enableSorting: true,
                  fields: [
                    {
                      id: 'title',
                      title: strings.titleField,
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'description',
                      title: strings.descriptionField,
                      type: CustomCollectionFieldType.string
                    },
                    {
                      id: 'url',
                      title: strings.urlField,
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'imageUrl',
                      title: strings.imageUrlField,
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: this.renderTileImagePicker.bind(this)
                    },
                    {
                      id: 'imageOpacity',
                      title: strings.imageOpacityField,
                      type: CustomCollectionFieldType.custom,
                      defaultValue: 100,
                      onCustomRender: this.renderImageOpacity.bind(this)
                    },
                    {
                      id: 'imageAsBackground',
                      title: strings.imageAsBackgroundField,
                      type: CustomCollectionFieldType.boolean,
                      defaultValue: false
                    },
                    {
                      id: 'imageOnly',
                      title: strings.imageOnlyField,
                      type: CustomCollectionFieldType.boolean,
                      defaultValue: false
                    },
                    {
                      id: 'imagePosition',
                      title: strings.imagePositionField,
                      type: CustomCollectionFieldType.dropdown,
                      defaultValue: 'top',
                      options: [
                        { key: 'top', text: 'Top' },
                        { key: 'bottom', text: 'Bottom' },
                        { key: 'left', text: 'Left' },
                        { key: 'right', text: 'Right' }
                      ]
                    },
                    {
                      id: 'hoverImageUrl',
                      title: strings.hoverImageUrlField,
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: this.renderTileImagePicker.bind(this)
                    },
                    {
                      id: 'hoverImageAsBackground',
                      title: strings.hoverImageAsBackgroundField,
                      type: CustomCollectionFieldType.boolean,
                      defaultValue: false
                    },
                    {
                      id: 'hoverImageOnly',
                      title: strings.hoverImageOnlyField,
                      type: CustomCollectionFieldType.boolean,
                      defaultValue: false
                    },
                    {
                      id: 'hoverImagePosition',
                      title: strings.hoverImagePositionField,
                      type: CustomCollectionFieldType.dropdown,
                      defaultValue: 'top',
                      options: [
                        { key: 'top', text: 'Top' },
                        { key: 'bottom', text: 'Bottom' },
                        { key: 'left', text: 'Left' },
                        { key: 'right', text: 'Right' }
                      ]
                    },
                    {
                      id: 'target',
                      title: strings.targetField,
                      type: CustomCollectionFieldType.dropdown,
                      defaultValue: '',
                      options: [
                        { key: '', text: strings.targetCurrent },
                        { key: '_blank', text: strings.targetNew }
                      ]
                    },
                    {
                      id: 'color',
                      title: strings.tileColorOverrideField,
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: this.renderTileColorPicker.bind(this)
                    }
                  ]
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
              groupFields: ([
                PropertyPaneDropdown('tileLayout', {
                  label: strings.tileLayoutLabel,
                  options: [
                    { key: 'stacked', text: strings.tileLayoutStacked },
                    { key: 'columns', text: strings.tileLayoutColumns },
                    { key: 'flow', text: strings.tileLayoutFlow }
                  ],
                  selectedKey: this.properties.tileLayout || 'columns'
                })
              ] as any[]).concat((this.properties.tileLayout || 'columns') === 'columns' ? [
                PropertyPaneSlider('columnCount', {
                  label: strings.columnCountLabel,
                  min: 1,
                  max: 12,
                  step: 1,
                  value: this.properties.columnCount || 3
                })
              ] : []).concat([
                PropertyPaneSlider('tileGap', {
                  label: strings.tileGapLabel,
                  min: 0,
                  max: 100,
                  step: 1,
                  value: typeof this.properties.tileGap === 'number' ? this.properties.tileGap : 16
                }),
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
              ])
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

