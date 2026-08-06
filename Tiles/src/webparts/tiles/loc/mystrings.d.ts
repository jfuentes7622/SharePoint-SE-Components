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
  hoverImageOnlyField: string;
  hoverImagePositionField: string;
  targetField: string;

  targetCurrent: string;
  targetNew: string;

  // Per-tile collection fields
  imageOnlyField: string;
  tileColorOverrideField: string;

  // Property pane groups
  tileDataGroup: string;
  appearanceGroup: string;
  typographyGroup: string;
  sizeGroup: string;
  diagnosticsGroup: string;

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
