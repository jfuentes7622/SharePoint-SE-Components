import * as React from 'react';

import styles from '../components/TabComponent.module.scss';
import { ITabGlobalFontSettings, ITabVisualSettings } from './ITabComponentProps';

export interface ITabProps {
    TabIndex: number;
    TabText: string;
    TabSelected: boolean;
  TabSettings?: ITabVisualSettings;
    GlobalFontSettings: ITabGlobalFontSettings;
    TabClickCallback: (tabIndex: number) => void;
    EnableDiagnostics?: boolean;
}

export default class Tab extends React.Component<ITabProps, {}> {
  private tabLabelRef: HTMLSpanElement | undefined;
  private tabContentRef: HTMLSpanElement | undefined;

  private logDiagnostic(message: string): void {
    if (this.props.EnableDiagnostics === false) {
      return;
    }

    console.log('[Tab] ' + message);
  }

  public componentDidMount(): void {
    this.logDiagnostic('Mounted tab "' + this.props.TabText + '" (index=' + String(this.props.TabIndex) + ').');
    this.applyDynamicStyles();
  }

  public componentDidUpdate(prevProps: ITabProps): void {
    if (prevProps.TabSelected !== this.props.TabSelected || prevProps.TabSettings !== this.props.TabSettings || prevProps.GlobalFontSettings !== this.props.GlobalFontSettings) {
      this.applyDynamicStyles();
    }
  }

    public render(): React.ReactElement<ITabProps> {
        const tabSettings: ITabVisualSettings = this.props.TabSettings || {};
        const hasImage: boolean = !!tabSettings.imageUrl;
        const showText: boolean = !(tabSettings.onlyImage === true);
        const showInlineImage: boolean = hasImage && showText;
        const imagePosition: string = tabSettings.imagePosition || 'left';
    
        return (
          <li onClick={this.callback.bind(this)} className={styles.tab} role='tab'>    
            <span
              ref={(el) => { this.tabLabelRef = el; }}
              className={`tabLabel ${this.props.TabSelected ? 'selected' : ''}`}
            >
              {
                hasImage && !showText &&
                <img src={tabSettings.imageUrl} alt='' role='presentation' aria-hidden={true} className='tabBackgroundImage' />
              }
              <span ref={(el) => { this.tabContentRef = el; }}
                className={'tabLabelContent ' + (showInlineImage ? 'hasTabImage tabImage-' + imagePosition : '')}>
                {
                  showInlineImage &&
                  <img src={tabSettings.imageUrl} alt={this.props.TabText} className='tabLabelImage' />
                }
                {
                  showText &&
                  <span className='tabLabelText'>{this.props.TabText}</span>
                }
              </span>
                </span>              
            </li>
        );
    }

    
    // Informs the tab-control that this tab was clicked to trigger a new render
    private callback(): void {
        this.props.TabClickCallback(this.props.TabIndex);
    }

    private applyDynamicStyles(): void {
      const tabLabel: HTMLSpanElement | undefined = this.tabLabelRef;
      const tabContent: HTMLSpanElement | undefined = this.tabContentRef;
      if (!tabLabel || !tabContent) {
        return;
      }

      const tabSettings: ITabVisualSettings = this.props.TabSettings || {};
      const globalFontSettings: ITabGlobalFontSettings = this.props.GlobalFontSettings || {};
      const isRounded: boolean = (globalFontSettings.tabShape || 'rounded') === 'rounded';
      const imageFirst: boolean = (tabSettings.imagePosition || 'left') === 'left';
      const textPosition: 'left' | 'center' | 'right' = tabSettings.textPosition || 'left';
      const justifyContent: string = textPosition === 'center' ? 'center' : (textPosition === 'right' ? 'flex-end' : 'flex-start');
      const textAlign: string = textPosition === 'center' ? 'center' : (textPosition === 'right' ? 'right' : 'left');
      const useAutoWidth: boolean = globalFontSettings.autoTabWidth === true;
      const inactiveImageFade: number = typeof globalFontSettings.inactiveImageFade === 'number'
        ? Math.max(0, Math.min(100, globalFontSettings.inactiveImageFade))
        : 45;
      const imageOpacity: number = this.props.TabSelected ? 1 : (100 - inactiveImageFade) / 100;

      tabLabel.style.setProperty('--tab-justify', justifyContent);
      tabLabel.style.setProperty('--tab-text-align', textAlign);
      tabLabel.style.setProperty('--tab-font-family', globalFontSettings.fontFamily || 'inherit');
      tabLabel.style.setProperty('--tab-font-style', globalFontSettings.fontStyle || 'normal');
      tabLabel.style.setProperty('--tab-font-weight', globalFontSettings.isBold ? '700' : '400');
      tabLabel.style.setProperty('--tab-font-size', `${globalFontSettings.fontSize || 14}px`);
      tabLabel.style.setProperty('--tab-height', `${globalFontSettings.tabHeight || 60}px`);
      tabLabel.style.setProperty('--tab-min-width', useAutoWidth ? '60px' : `${globalFontSettings.tabWidth || 120}px`);
      tabLabel.style.setProperty('--tab-width', useAutoWidth ? 'auto' : `${globalFontSettings.tabWidth || 120}px`);
      tabLabel.style.setProperty('color', globalFontSettings.tabTextColor || '#ffffff');
      tabLabel.style.setProperty('border-color', globalFontSettings.tabBorderColor || '#000000');
      tabLabel.style.setProperty('border-width', `${globalFontSettings.tabBorderWidth || 0}px`);
      tabLabel.style.setProperty('border-style', globalFontSettings.tabBorderStyle || 'solid');
      tabLabel.style.setProperty('--tab-padding-x', tabSettings.onlyImage === true || !!tabSettings.imageUrl ? '0px' : '14px');
      tabLabel.style.setProperty('--tab-top-left-radius', isRounded ? `${globalFontSettings.tabCornerRadius || 0}px` : '0px');
      tabLabel.style.setProperty('--tab-top-right-radius', isRounded ? `${globalFontSettings.tabCornerRadius || 0}px` : '0px');
      tabLabel.style.setProperty('--tab-image-opacity', String(imageOpacity));

      tabContent.style.setProperty('--tab-content-direction', imageFirst ? 'row' : 'row-reverse');
    }
}