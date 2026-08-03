import * as React from 'react';

import styles from '../components/TabComponent.module.scss';
import { ITabGlobalFontSettings, ITabVisualSettings } from './ITabComponentProps';

import { Logger, LogLevel } from 'sp-pnp-js';

export interface ITabProps {
    TabIndex: number;
    TabText: string;
    TabSelected: boolean;
  TabSettings?: ITabVisualSettings;
    GlobalFontSettings: ITabGlobalFontSettings;
    TabClickCallback: (tabIndex: number) => void;
}

const LOG_SOURCE: string = 'KM Tabs - ';
  
  declare global {
    interface Window { 
         iskmActiveLog: boolean;
    }
  }

export default class Tab extends React.Component<ITabProps, {}> {
  private tabLabelRef: HTMLSpanElement | undefined;
  private tabContentRef: HTMLSpanElement | undefined;

  public componentDidMount(): void {
    this.applyDynamicStyles();
  }

  public componentDidUpdate(prevProps: ITabProps): void {
    if (prevProps.TabSelected !== this.props.TabSelected || prevProps.TabSettings !== this.props.TabSettings || prevProps.GlobalFontSettings !== this.props.GlobalFontSettings) {
      this.applyDynamicStyles();
    }
  }

    public render(): React.ReactElement<ITabProps> {
        if (!window.iskmActiveLog || window.iskmActiveLog === undefined) {
            Logger.activeLogLevel = LogLevel.Error;
          }
          else {Logger.activeLogLevel=LogLevel.Info;
        }

        const tabSettings: ITabVisualSettings = this.props.TabSettings || {};
        const hasImage: boolean = !!tabSettings.imageUrl;
        const showText: boolean = !(tabSettings.onlyImage === true);
    
        Logger.write(LOG_SOURCE + 'In Tab.tx Render', LogLevel.Info);
        return (
          <li onClick={this.callback.bind(this)} className={styles.tab} role='tab'>    
            <span
              ref={(el) => { this.tabLabelRef = el; }}
              className={`tabLabel ${this.props.TabSelected ? 'selected' : ''}`}
            >
              <span ref={(el) => { this.tabContentRef = el; }} className='tabLabelContent'>
                {
                  hasImage &&
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
      const useAutoWidth: boolean = globalFontSettings.autoTabWidth === true || tabSettings.onlyImage === true;

      tabLabel.style.setProperty('--tab-justify', justifyContent);
      tabLabel.style.setProperty('--tab-text-align', textAlign);
      tabLabel.style.setProperty('--tab-font-family', globalFontSettings.fontFamily || 'inherit');
      tabLabel.style.setProperty('--tab-font-style', globalFontSettings.fontStyle || 'normal');
      tabLabel.style.setProperty('--tab-font-weight', globalFontSettings.isBold ? '700' : '400');
      tabLabel.style.setProperty('--tab-font-size', `${globalFontSettings.fontSize || 14}px`);
      tabLabel.style.setProperty('--tab-height', `${globalFontSettings.tabHeight || 60}px`);
      tabLabel.style.setProperty('--tab-min-width', useAutoWidth ? '60px' : `${globalFontSettings.tabWidth || 120}px`);
      tabLabel.style.setProperty('--tab-width', useAutoWidth ? 'auto' : `${globalFontSettings.tabWidth || 120}px`);
      tabLabel.style.setProperty('--tab-padding-x', tabSettings.onlyImage === true ? '0px' : '14px');
      tabLabel.style.setProperty('--tab-top-left-radius', isRounded ? '10px' : '0px');
      tabLabel.style.setProperty('--tab-top-right-radius', isRounded ? '10px' : '0px');

      tabContent.style.setProperty('--tab-content-direction', imageFirst ? 'row' : 'row-reverse');
    }
}