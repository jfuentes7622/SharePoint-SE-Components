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
}

export interface ITabComponentProps {
  TabHeaders: Array<string>;
  ControlZones: Array<HTMLElement>;
  PageInEditMode: boolean;
  TabConfigs: ITabVisualSettings[];
  GlobalFontSettings: ITabGlobalFontSettings;
}

export interface ITabControlState {
  SelectedTab: number;
}