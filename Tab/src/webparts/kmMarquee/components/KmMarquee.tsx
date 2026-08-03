import * as React from 'react';
import styles from './KmMarquee.module.scss';
import { IKmMarqueeProps } from './IKmMarqueeProps';
import {DisplayMode} from '@microsoft/sp-core-library';
import * as strings from 'KmMarqueeWebPartStrings';
import './KmMarquee.css';
// import {
//   Logger,
//   ConsoleListener,
//   LogLevel
// } from "@pnp/logging";

import { Logger, ConsoleListener, LogLevel } from 'sp-pnp-js';

declare global {
  interface Window { 
       iskmActiveLog: boolean;
  }
}
//import PnPTelemetry from "@pnp/telemetry-js";
const LOG_SOURCE: string = 'SPS Marquee - ';

export default class KmMarquee extends React.Component<IKmMarqueeProps, {}> {

  constructor(props: IKmMarqueeProps) {
    super(props);
    this.state = { loading: false };
    this.handleLoad = this.handleLoad.bind(this);
   }

 componentWillMount() {
    Logger.write(LOG_SOURCE+'ComponentWillMount', LogLevel.Info);
    this.handleLoad();
  }

 componentDidMount() {
  Logger.write(LOG_SOURCE+'ComponentDidMount', LogLevel.Info);
  setTimeout(() => {window.addEventListener('load', this.handleLoad);},0);
} 

  componentDidUpdate(prevProps: any):void {
    Logger.write(LOG_SOURCE+'ComponentWillUpdate', LogLevel.Info);
    if ((prevProps.description!==this.props.description) || (prevProps.marqueeTextColor !==this.props.marqueeTextColor) || (prevProps.marqueeBackColor!==this.props.marqueeBackColor) || (prevProps.marqueeActive!==this.props.marqueeActive)) {
      this.handleLoad();
    }
  }


  componentWillUnmount() { 
    Logger.write(LOG_SOURCE+'ComponentWillUnmount', LogLevel. Info);
    window.removeEventListener('load', this.handleLoad);
    //this.removeDiv();
  }

  
  private handleLoad() {
    const topMarqueeParent: HTMLElement = this.getParentElement(strings.ParentElement);
    Logger.write(LOG_SOURCE+ 'Top Marquee Element:'+ topMarqueeParent, LogLevel.Info);
    Logger.write(LOG_SOURCE + `this,props,marqueeActive:${this.props.marqueeActive}`, LogLevel.Info);
    Logger.write(LOG_SOURCE + `this,props.description:${this.props.description}`, LogLevel.Info);
    if (this.props.marqueeActive) {
      this.insertDiv(topMarqueeParent); 
      this.setState({ loading: true }); }
    else {
      this.removeDiv();
      this.setState({ loading: false });
    }
   }
  
   
    private getParentElement(parentFallbackList: string): HTMLElement {
      // check for null or empty fallback list (only whitespace characters present)
      if (!parentFallbackList ||
          parentFallbackList.match(/^\s*$/)) {
          Logger.write(LOG_SOURCE + 'WARNING: empty fallback list supplied; defaulting to document body', LogLevel.Warning);
          return document.body;
      }
      // parse string for fallback list
      const fallbackList: string[] = parentFallbackList.split('|');
      // declare null parent element
      let parent: HTMLElement | undefined = undefined;
      // get the first valid element
      fallbackList.forEach(element=> {
          try {
              const queryElement = document.querySelector(element) as HTMLElement;
              parent = parent || queryElement;
          } catch (err) {
              Logger.write(LOG_SOURCE +  err, LogLevel.Error);
          }
      });
      // return the proper parent, or the body element if none of the fallbacks succeed
      return parent || document.body;
  }

  public removeDiv() {    
    Logger.write(LOG_SOURCE + `In removeDiv`, LogLevel.Info);
    const marqueeDiv: HTMLElement = document.getElementById('marqueeDiv') as HTMLElement;
    const marqueeSpan: HTMLElement=document.getElementById('marqueeSpan') as HTMLElement;
    Logger.write(LOG_SOURCE +  `marqueeDiv found:${marqueeDiv}`, LogLevel.Info);
    // check for existing marqueeDiv and span and remove
    if (marqueeDiv) {marqueeDiv.remove();}
    if(marqueeSpan) {marqueeSpan.remove();}
    } 
   
 public insertDiv(marqueeParent: HTMLElement) {    
    const innerHtml: string = this.props.description;
    Logger.write(LOG_SOURCE + `In insertDiv`, LogLevel.Info);
    let marqueeDiv: HTMLElement = document.getElementById('marqueeDiv') as HTMLElement;
    const marqueeSpan: HTMLElement= document.getElementById('marqueeSpan') as HTMLElement;
    // check for existing maarqueeDiv
    if (!marqueeDiv) {
      marqueeDiv = document.createElement('div');
      marqueeDiv.id = 'marqueeDiv';
      marqueeDiv.className = 'bar';
      marqueeParent.insertAdjacentElement('beforebegin', marqueeDiv);
      marqueeDiv.style.color= this.props.marqueeTextColor;
      marqueeDiv.style.backgroundColor=this.props.marqueeBackColor;
      const marqueeSpan2 = document.createElement('span');
      marqueeSpan2.className='bar_content';
      marqueeSpan2.id='marqueeSpan';
      marqueeSpan2.innerHTML = innerHtml;   
      (document.getElementById('marqueeDiv') as HTMLElement).appendChild(marqueeSpan2);
      this.forceUpdate();
    }
    else {   // in any case re-set the inner html and style
      marqueeSpan.innerHTML = innerHtml;
      marqueeDiv.style.color= this.props.marqueeTextColor;
      marqueeDiv.style.backgroundColor=this.props.marqueeBackColor;
      this.forceUpdate(); }
  }

  public openPane(): void {
    this.props.spfxContext.propertyPane.open();
  }

  public render(): React.ReactElement<IKmMarqueeProps> {
    Logger.subscribe(new ConsoleListener());
    if (!window.iskmActiveLog || window.iskmActiveLog===undefined) {
      Logger.activeLogLevel = LogLevel.Error;
    }
    else {Logger.activeLogLevel=LogLevel.Info;}

    // const telemetry = PnPTelemetry.getInstance();
    // telemetry.optOut();
  
    if (DisplayMode.Edit===this.props.dMode || this.props.description===undefined) {
      return (
        <div className={styles.kmMarquee}>
          <div className="col">
            <p className="$ms-color-themePrimary ms-font-xl">Marquee Web Part</p>
            <button className={styles.button} onClick={this.openPane.bind(this)}>Click here to configure</button>
          </div>
        </div>
      );
    }

    if (this.props.description.length>0 && DisplayMode.Read===this.props.dMode) {
    return( 
      <div className={styles.kmMarquee}>
      <div className="col">
       <button className={styles.buttonH}>Hidden button</button>
      </div>    
      </div >
    );
     }
     return (
      <div></div>
     );
  }
}
