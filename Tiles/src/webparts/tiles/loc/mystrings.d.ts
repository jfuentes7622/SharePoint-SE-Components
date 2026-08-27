declare interface ITilesWebPartStrings {
  PropertyPaneDescription: string;
  TilesListDescription: string;
  TileHeight: string;
  iconInformation: string;

  // Properties
  tilesDataLabel: string;
  tilesPanelHeader: string;
  tilesManageBtn: string;

  // Tile fields
  titleField: string;
  descriptionField: string;
  urlField: string;
  iconField: string;
  imageUrlField: string;
  imagePositionField: string;
  hoverImageUrlField: string;
  hoverImageAsBackgroundField: string;
  hoverImageOnlyField: string;
  hoverImagePositionField: string;
  targetField: string;

  targetCurrent: string;
  targetNew: string;

  // Per-tile collection fields
  imageOpacityField: string;
  imageAsBackgroundField: string;
  imageOnlyField: string;
  tileColorOverrideField: string;

  // Property pane groups
  webPartAppearanceGroup: string;
  titleAppearanceGroup: string;
  tileDataGroup: string;
  appearanceGroup: string;
  typographyGroup: string;
  sizeGroup: string;
  tileLayoutLabel: string;
  tileLayoutStacked: string;
  tileLayoutColumns: string;
  tileLayoutFlow: string;
  columnCountLabel: string;
  tileGapLabel: string;
  diagnosticsGroup: string;

  // Web part appearance
  webPartBackgroundColorLabel: string;
  webPartBorderColorLabel: string;
  webPartBorderWidthLabel: string;
  webPartBorderStyleLabel: string;
  webPartCornerRadiusLabel: string;
  webPartPaddingLabel: string;
  showTitleLabel: string;
  titleTextColorLabel: string;
  titleFontFamilyLabel: string;
  titleFontSizeLabel: string;
  titleFontStyleLabel: string;
  titleFontBoldLabel: string;
  titleAlignmentLabel: string;
  titleBottomSpacingLabel: string;
  alignmentLeft: string;
  alignmentCenter: string;
  alignmentRight: string;

  // Appearance
  tileShapeLabel: string;
  tileShapeSquared: string;
  tileShapeRounded: string;
  tileShapeRound: string;
  backgroundColorLabel: string;
  textColorLabel: string;
  hoverColorLabel: string;
  hoverSameAsBackgroundLabel: string;
  hoverTransitionLabel: string;
  hoverTransitionSolid: string;
  hoverTransitionFade: string;
  hoverTransitionSlideTop: string;
  hoverTransitionSlideBottom: string;
  hoverTransitionSlideLeft: string;
  hoverTransitionSlideRight: string;

  // Typography
  fontFamilyLabel: string;
  fontFamilyInherit: string;
  fontStyleLabel: string;
  fontStyleNormal: string;
  fontStyleItalic: string;
  fontStyleOblique: string;
  fontBoldLabel: string;
  fontSizeLabel: string;

  // Size
  tileWidthLabel: string;
  tileHeightLabel: string;

  // Diagnostics
  enableDiagnosticsLabel: string;

  // Component
  noTilesIconText: string;
  noTilesConfigured: string;
  noTilesBtn: string;
}

declare module 'TilesWebPartStrings' {
  const strings: ITilesWebPartStrings;
  export = strings;
}
