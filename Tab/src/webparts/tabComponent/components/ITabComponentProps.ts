export interface ITabVisualSettings {
  textPosition?: 'left' | 'center' | 'right';
  imageUrl?: string;
  imagePosition?: 'left' | 'right';
  onlyImage?: boolean;
}

export interface ITabGlobalFontSettings {
  fontFamily?: 'Segoe UI' | 'Arial' | 'Calibri' | 'Verdana' | 'Tahoma' | 'Trebuchet MS' | 'Georgia' | 'Times New Roman' | 'Courier New';
  fontStyle?: 'normal' | 'italic' | 'oblique';
  isBold?: boolean;
  fontSize?: number;
  tabHeight?: number;
  tabWidth?: number;
  autoTabWidth?: boolean;
  tabShape?: 'rounded' | 'square';
  tabTextColor?: string;
  tabBorderColor?: string;
  tabBorderWidth?: number;
  tabBorderStyle?: 'none' | 'solid' | 'dashed' | 'dotted' | 'double';
  tabCornerRadius?: number;
  tabLineColor?: string;
  tabLineWidth?: number;
  webPartBorderColor?: string;
  webPartBorderWidth?: number;
  webPartBorderStyle?: 'none' | 'solid' | 'dashed' | 'dotted' | 'double';
  webPartCornerStyle?: 'square' | 'rounded';
  webPartCornerRadius?: number;
  inactiveImageFade?: number;
  contentPaddingTop?: number;
  contentPaddingBottom?: number;
}

export interface ITabComponentProps {
  TabHeaders: Array<string>;
  ControlZones: Array<HTMLElement>;
  PageInEditMode: boolean;
  TabConfigs: ITabVisualSettings[];
  GlobalFontSettings: ITabGlobalFontSettings;
  EnableDiagnostics?: boolean;
}

export interface ITabControlState {
  SelectedTab: number;
}