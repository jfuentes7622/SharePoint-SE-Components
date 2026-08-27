import * as React from 'react';
import styles from './TabComponent.module.scss';
import { ITabComponentProps,ITabControlState } from './ITabComponentProps';

import Tab from './Tab';

export default class TabControl extends React.Component<ITabComponentProps, ITabControlState> {
    private borderedZones: HTMLElement[] = [];

  constructor(props: ITabComponentProps, state: ITabControlState) {
      super(props);
      this.state = { SelectedTab: 0 };

      this.logDiagnostic('Constructor invoked. Tab count=' + String(this.props.TabHeaders.length));
  }

  private logDiagnostic(message: string): void {
    if (this.props.EnableDiagnostics === false) {
      return;
    }

    console.log('[TabComponent] ' + message);
  }


  
  /*  shouldComponentUpdate(nextProps:ITabControlProps, nextState:ITabControlState) {
   console.log('ShouldUpdate!')
   console.log("thisState:", this.state.SelectedTab);
   console.log("NextState:", nextState.SelectedTab);
    return this.state.SelectedTab!= nextState.SelectedTab; 
  } */
  public componentDidUpdate(prevProps: ITabComponentProps, prevState: ITabControlState): void {
      //if properties have changes bind it
      if (this.state.SelectedTab !== prevState.SelectedTab) {
        this.logDiagnostic('Selected tab changed from ' + String(prevState.SelectedTab) + ' to ' + String(this.state.SelectedTab) + '.');
                this.RefreshLists();
      }
    } 

    public componentWillUnmount(): void {
            this.clearContentBorders();
    }
  
  public render(): React.ReactElement<ITabComponentProps> {
      const settings = this.props.GlobalFontSettings || {};
      const borderWidth = String(settings.webPartBorderWidth || 0) + 'px';
      const cornerRadius = settings.webPartCornerStyle === 'rounded'
          ? String(settings.webPartCornerRadius || 0) + 'px'
          : '0px';
      const webPartStyle: any = {
          borderColor: settings.webPartBorderColor || '#cccccc',
          borderWidth: borderWidth,
          borderBottomWidth: this.props.ControlZones.length > 0 ? '0px' : borderWidth,
          borderStyle: settings.webPartBorderStyle || 'solid',
          borderTopLeftRadius: cornerRadius,
          borderTopRightRadius: cornerRadius,
          borderBottomLeftRadius: this.props.ControlZones.length > 0 ? '0px' : cornerRadius,
          borderBottomRightRadius: this.props.ControlZones.length > 0 ? '0px' : cornerRadius,
          boxSizing: 'border-box',
          overflow: settings.webPartCornerStyle === 'rounded' ? 'hidden' : 'visible'
      };
      const navigationStyle: any = {
          borderBottom: String(settings.tabLineWidth || 0) + 'px solid ' + (settings.tabLineColor || '#8A1717')
      };
      return (
          <div className={styles.tabControl} style={webPartStyle}>
              <div className='cd-tabs'>
                  <nav>
                      <ul className='cd-tabs-navigation' role='tablist' style={navigationStyle}>
                          {
                              this.props.ControlZones.map((d, i, e) => {
                                  this.applyContentBorder(d, i === this.state.SelectedTab && this.props.PageInEditMode === false);
                                  // Hide the current section if it's not associated with the current selected tab
                                  if (this.props.PageInEditMode === false && !(i === this.state.SelectedTab)) {
                                      d.classList.add('sectionHidden');
                                  } else {
                                      d.classList.remove('sectionHidden');
                                  }
                                  return (
                                      <Tab
                                          key={`tab-${i}`}
                                          TabIndex={i}
                                          TabText={this.props.TabHeaders[i]}
                                          TabSelected={i === this.state.SelectedTab}
                                          TabSettings={this.props.TabConfigs[i]}
                                          GlobalFontSettings={this.props.GlobalFontSettings}
                                          TabClickCallback={this.tabClicked.bind(this)}
                                          EnableDiagnostics={this.props.EnableDiagnostics}
                                      />
                                  );
                              })
                          }
                      </ul>
                  </nav>
                  {
                      this.props.PageInEditMode &&
                      <div>
                          <div className='p4'>
                              Note: When exiting edit-mode under certain conditions, you may experience
                              issues with tab-content not displaying correctly which will resolve once
                              the page is refreshed.
                          </div>
                          <hr />
                      </div>
                  }
              </div>
          </div>
           
      );
      
  }

    private findReact(dom:any): any {
                        if (!dom) {
                            return undefined;
                        }
            for (const key in dom) {
                if (key.startsWith("__reactInternalInstance$")) {
                                        const internalInstance = dom[key];
                                        const currentElement = internalInstance && internalInstance._currentElement;
                                        const owner = currentElement && currentElement._owner;
                                        return owner ? owner._instance : undefined;
                }
      }
  }

  public RefreshLists():void
  {
      //console.log("Refresh Lists if exists");
        //lets refresh the lists/libs when changing tabs
        this.props.ControlZones.forEach(dom => {
          let lists: NodeListOf<Element>;
          lists = dom.querySelectorAll("div[class^='detailsListContainer']");
         
          if (lists) {
              
              Array.from(lists).forEach(el => {
                  const rInstance = this.findReact(el.firstElementChild);
                  if (rInstance && typeof rInstance.forceUpdate === 'function') {
                      rInstance.forceUpdate();
                  }
              });
          }
      });
  }

  private applyContentBorder(zone: HTMLElement, isSelected: boolean): void {
      const className = 'spseTabOverallBorderContent';
      const borderTarget = this.getContentBorderTarget(zone);
      if (!isSelected) {
          borderTarget.classList.remove(className);
          this.clearContentBorderVariables(borderTarget);
          return;
      }

      const settings = this.props.GlobalFontSettings || {};
      borderTarget.classList.add(className);
      borderTarget.style.setProperty('--spse-tab-border-color', settings.webPartBorderColor || '#cccccc');
      borderTarget.style.setProperty('--spse-tab-border-width', String(settings.webPartBorderWidth || 0) + 'px');
      borderTarget.style.setProperty('--spse-tab-border-style', settings.webPartBorderStyle || 'solid');
      borderTarget.style.setProperty('--spse-tab-border-radius', settings.webPartCornerStyle === 'rounded'
          ? String(settings.webPartCornerRadius || 0) + 'px'
          : '0px');
      borderTarget.style.setProperty('--spse-tab-content-padding-top', String(settings.contentPaddingTop || 0) + 'px');
      borderTarget.style.setProperty('--spse-tab-content-padding-bottom', String(settings.contentPaddingBottom || 0) + 'px');
      if (this.borderedZones.indexOf(borderTarget) < 0) {
          this.borderedZones.push(borderTarget);
      }
  }

  private getContentBorderTarget(zone: HTMLElement): HTMLElement {
      const innerControlZone = zone.classList.contains('CanvasZone')
          ? zone.querySelector('.ControlZone') as HTMLElement
          : undefined;
      return innerControlZone || zone;
  }

  private clearContentBorders(): void {
      this.borderedZones.forEach((zone: HTMLElement) => {
          zone.classList.remove('spseTabOverallBorderContent');
          this.clearContentBorderVariables(zone);
      });
      this.borderedZones = [];
  }

  private clearContentBorderVariables(zone: HTMLElement): void {
      zone.style.removeProperty('--spse-tab-border-color');
      zone.style.removeProperty('--spse-tab-border-width');
      zone.style.removeProperty('--spse-tab-border-style');
      zone.style.removeProperty('--spse-tab-border-radius');
      zone.style.removeProperty('--spse-tab-content-padding-top');
      zone.style.removeProperty('--spse-tab-content-padding-bottom');
  }

  // Handler for when a tab component is clicked
  private tabClicked(tabIndex: number): void {
      //Set New tab
      if (this.state.SelectedTab !== tabIndex) {
          this.logDiagnostic('Tab clicked. Index=' + String(tabIndex) + ', Header="' + (this.props.TabHeaders[tabIndex] || '') + '".');
          this.setState({ SelectedTab: tabIndex });
            } else {
                    this.RefreshLists();
            }
  }
}
