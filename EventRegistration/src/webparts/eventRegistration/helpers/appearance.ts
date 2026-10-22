export interface IEventRegistrationAppearance {
  appearanceMode?: string;
  surfaceBackgroundColor?: string;
  surfaceTextColor?: string;
  accentColor?: string;
  buttonTextColor?: string;
  activeButtonColor?: string;
  borderColor?: string;
  fontFamily?: string;
  fontSize?: number;
  borderWidth?: number;
  cornerRadius?: number;
}

export interface IEventRegistrationTheme {
  semanticColors?: { [name: string]: string };
  palette?: { [name: string]: string };
  fonts?: { medium?: { fontFamily?: string } };
}

export function getAppearanceMode(value: string): string {
  return value === 'siteTheme' || value === 'siteThemeOverrides' ? value : 'manual';
}

export function applyAppearance(
  style: { setProperty(name: string, value: string): void },
  properties: IEventRegistrationAppearance,
  theme: IEventRegistrationTheme
): void {
  var mode = getAppearanceMode(properties.appearanceMode);
  var siteTheme = theme || {};
  function resolve(manualValue: string, semanticName: string, paletteName: string, fallback: string): string {
    var manual = String(manualValue || '').trim();
    if (mode === 'manual' || (mode === 'siteThemeOverrides' && manual)) {
      return manual || fallback;
    }
    return String(siteTheme.semanticColors && siteTheme.semanticColors[semanticName]
      || siteTheme.palette && siteTheme.palette[paletteName] || fallback);
  }

  style.setProperty('--er-background', resolve(properties.surfaceBackgroundColor, 'bodyBackground', 'white', '#ffffff'));
  style.setProperty('--er-text', resolve(properties.surfaceTextColor, 'bodyText', 'neutralPrimary', '#323130'));
  style.setProperty('--er-accent', resolve(properties.accentColor, 'primaryButtonBackground', 'themePrimary', '#0078d4'));
  style.setProperty('--er-button-text', resolve(properties.buttonTextColor, 'primaryButtonText', 'white', '#ffffff'));
  style.setProperty('--er-active', resolve(properties.activeButtonColor, 'primaryButtonBackgroundPressed', 'themeDark', '#004578'));
  style.setProperty('--er-border', resolve(properties.borderColor, 'inputBorder', 'neutralTertiary', '#c8c6c4'));
  style.setProperty('--er-muted', resolve('', 'bodySubtext', 'neutralSecondary', '#605e5c'));
  style.setProperty('--er-table-header', resolve('', 'bodyStandoutBackground', 'neutralLighter', '#f3f2f1'));
  var font = String(properties.fontFamily || '').trim();
  if (mode === 'siteTheme' || (mode === 'siteThemeOverrides' && !font)) {
    font = String(siteTheme.fonts && siteTheme.fonts.medium && siteTheme.fonts.medium.fontFamily || '');
  }
  style.setProperty('--er-font', font || '"Segoe UI", Arial, sans-serif');
  style.setProperty('--er-font-size', String(typeof properties.fontSize === 'number' ? properties.fontSize : 14) + 'px');
  style.setProperty('--er-border-width', String(typeof properties.borderWidth === 'number' ? properties.borderWidth : 0) + 'px');
  style.setProperty('--er-radius', String(typeof properties.cornerRadius === 'number' ? properties.cornerRadius : 0) + 'px');
}
