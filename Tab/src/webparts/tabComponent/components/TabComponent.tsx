import * as React from 'react';
import styles from './TabComponent.module.scss';
import { ITabComponentProps,ITabControlState } from './ITabComponentProps';

import Tab from './Tab';
import { Logger, LogLevel } from 'sp-pnp-js';

  const LOG_SOURCE: string = 'KM Tabs - ';
  
  declare global {
    interface Window { 
         iskmActiveLog: boolean;
    }
  }

export default class TabControl extends React.Component<ITabComponentProps, ITabControlState> {
  constructor(props: ITabComponentProps, state: ITabControlState) {
      super(props);
      this.state = { SelectedTab: 0 };

            if (!window.iskmActiveLog || window.iskmActiveLog === undefined) {
            console.log("if");
          Logger.activeLogLevel = LogLevel.Error;
        }
        else {Logger.activeLogLevel=LogLevel.Info;
            console.log("else");
      }
      
      Logger.write(LOG_SOURCE + 'In TabComponent.tsx constructor', LogLevel.Info);
  }


  
  /*  shouldComponentUpdate(nextProps:ITabControlProps, nextState:ITabControlState) {
   console.log('ShouldUpdate!')
   console.log("thisState:", this.state.SelectedTab);
   console.log("NextState:", nextState.SelectedTab);
    return this.state.SelectedTab!= nextState.SelectedTab; 
  } */
  public componentDidUpdate(prevProps: ITabComponentProps, prevState: ITabControlState): void {
      //if properties have changes bind it
      Logger.write(LOG_SOURCE + 'ComponentDid Update!', LogLevel.Info);
      Logger.write(LOG_SOURCE + "thisState:" + this.state.SelectedTab.toString(), LogLevel.Info);
      Logger.write(LOG_SOURCE + "PrevState:" + prevState.SelectedTab.toString(), LogLevel.Info);
      if (this.state.SelectedTab !== prevState.SelectedTab) {
        this.forceUpdate();
      }
    } 
  
  public render(): React.ReactElement<ITabComponentProps> {
    /*
      const telemetry = PnPTelemetry.getInstance();
      telemetry.optOut();

      Logger.subscribe(ConsoleListener());
      if (!window.iskmActiveLog || window.iskmActiveLog==undefined) {
        console.log("if render");
          Logger.activeLogLevel = LogLevel.Error;
        }
        else {Logger.activeLogLevel=LogLevel.Info;
            console.log("else render");
      }
      */

      Logger.write(LOG_SOURCE + 'In TabComponent.tsx Render', LogLevel.Info);
      return (
          <div id='ISKMTabControl' className={styles.tabControl}>
              <div className='cd-tabs'>
                  <nav>
                      <ul className='cd-tabs-navigation' role='tablist'>
                          {
                              this.props.ControlZones.map((d, i, e) => {
                                  // Hide the current section if it's not associated with the current selected tab
                                  if (this.props.PageInEditMode === false && !(i === this.state.SelectedTab)) {
                                      d.classList.add('sectionHidden');
                                  } else {
                                      d.classList.remove('sectionHidden');
                                  }
                                  this.RefreshLists();
                                  return (
                                      <Tab
                                          key={`tab-${i}`}
                                          TabIndex={i}
                                          TabText={this.props.TabHeaders[i]}
                                          TabSelected={i === this.state.SelectedTab}
                                          TabSettings={this.props.TabConfigs[i]}
                                          GlobalFontSettings={this.props.GlobalFontSettings}
                                          TabClickCallback={this.tabClicked.bind(this)}
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
            for (const key in dom) {
                if (key.startsWith("__reactInternalInstance$")) {
                    return dom[key]._currentElement._owner._instance;
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
                  if (rInstance) {
                      rInstance.forceUpdate();
                  }
              });
          }
      });
  }

  // Handler for when a tab component is clicked
  private tabClicked(tabIndex: number): void {
      //Set New tab
      if (this.state.SelectedTab !== tabIndex) {
          this.setState({ SelectedTab: tabIndex });
          
      }      
    this.RefreshLists();
  }
}
