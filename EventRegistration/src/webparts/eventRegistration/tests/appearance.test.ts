import { assert } from 'chai';
import { applyAppearance, getAppearanceMode, IEventRegistrationAppearance, IEventRegistrationTheme } from '../helpers/appearance';

describe('Event Registration appearance', function(): void {
  var properties: IEventRegistrationAppearance;
  var theme: IEventRegistrationTheme;
  var values: { [name: string]: string };
  var style: { setProperty(name: string, value: string): void };

  beforeEach(function(): void {
    properties = {};
    theme = {};
    values = {};
    style = { setProperty: function(name: string, value: string): void { values[name] = value; } };
  });

  it('preserves manual defaults for existing instances', function(): void {
    assert.equal(getAppearanceMode(undefined), 'manual');
    assert.equal(getAppearanceMode('invalid'), 'manual');
    applyAppearance(style, properties, theme);
    assert.equal(values['--er-accent'], '#0078d4');
    assert.equal(values['--er-font-size'], '14px');
    assert.equal(values['--er-border-width'], '0px');
    assert.equal(values['--er-radius'], '0px');
  });

  it('applies manual colors and dimensions without using the site theme', function(): void {
    theme = { semanticColors: { bodyText: '#eeeeee' } };
    properties = { surfaceTextColor: '#123456', fontFamily: 'Arial', borderWidth: 3, cornerRadius: 8, fontSize: 18 };
    applyAppearance(style, properties, theme);
    assert.equal(values['--er-text'], '#123456');
    assert.equal(values['--er-font'], 'Arial');
    assert.equal(values['--er-border-width'], '3px');
    assert.equal(values['--er-radius'], '8px');
    assert.equal(values['--er-font-size'], '18px');
  });

  it('uses theme colors and fonts without deleting saved overrides in theme-only mode', function(): void {
    properties = { appearanceMode: 'siteTheme', surfaceTextColor: '#123456', fontFamily: 'Arial' };
    theme = {
      semanticColors: { bodyText: '#eeeeee' },
      palette: { themePrimary: '#abcdef' },
      fonts: { medium: { fontFamily: 'Verdana' } }
    };
    applyAppearance(style, properties, theme);
    assert.equal(values['--er-text'], '#eeeeee');
    assert.equal(values['--er-accent'], '#abcdef');
    assert.equal(values['--er-font'], 'Verdana');
    assert.equal(properties.surfaceTextColor, '#123456');
    assert.equal(properties.fontFamily, 'Arial');
  });

  it('uses explicit overrides and theme values for unset properties', function(): void {
    properties = { appearanceMode: 'siteThemeOverrides', surfaceTextColor: '#123456' };
    theme = { semanticColors: { bodyText: '#eeeeee', bodyBackground: '#111111' } };
    applyAppearance(style, properties, theme);
    assert.equal(values['--er-text'], '#123456');
    assert.equal(values['--er-background'], '#111111');
  });

  it('restores saved manual styling when switching back from theme-only mode', function(): void {
    properties = { appearanceMode: 'siteTheme', accentColor: '#123456', fontFamily: 'Arial' };
    theme = { palette: { themePrimary: '#abcdef' }, fonts: { medium: { fontFamily: 'Verdana' } } };
    applyAppearance(style, properties, theme);
    assert.equal(values['--er-accent'], '#abcdef');
    properties.appearanceMode = 'manual';
    applyAppearance(style, properties, theme);
    assert.equal(values['--er-accent'], '#123456');
    assert.equal(values['--er-font'], 'Arial');
  });
});
