import * as React from 'react';
import styles from './Marquee.module.scss';
import { IKmMarqueeProps } from './IMarqueeProps';
import {DisplayMode} from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import * as strings from 'KmMarqueeWebPartStrings';
import './Marquee.css';

export interface IKmMarqueeState {
  loading: boolean;
  messages: string[];
  currentIndex: number;
}

function escapeODataText(value: string): string {
  return value.replace(/'/g, "''");
}

export default class KmMarquee extends React.Component<IKmMarqueeProps, IKmMarqueeState> {
  private _cycleTimer: number;

  constructor(props: IKmMarqueeProps) {
    super(props);
    this.state = { loading: false, messages: [], currentIndex: 0 };
    this.handleLoad = this.handleLoad.bind(this);
   }

 componentWillMount() {
    this.logDiagnostic('componentWillMount');
    this.handleLoad();
  }

 componentDidMount() {
  this.logDiagnostic('componentDidMount');
  setTimeout(() => {window.addEventListener('load', this.handleLoad);},0);
  this.loadMessages();
}

  componentDidUpdate(prevProps: IKmMarqueeProps):void {
    this.logDiagnostic('componentDidUpdate');
    if (prevProps.listName !== this.props.listName || prevProps.messageField !== this.props.messageField) {
      this.loadMessages();
      return;
    }
    if (prevProps.messageDuration !== this.props.messageDuration) {
      this.startCycle();
    }
    if (
      prevProps.description !== this.props.description
      || prevProps.marqueeTextColor !== this.props.marqueeTextColor
      || prevProps.marqueeBackColor !== this.props.marqueeBackColor
      || prevProps.marqueeActive !== this.props.marqueeActive
      || prevProps.fontFamily !== this.props.fontFamily
      || prevProps.fontSize !== this.props.fontSize
      || prevProps.fontStyle !== this.props.fontStyle
      || prevProps.fontBold !== this.props.fontBold
      || prevProps.marqueeHeight !== this.props.marqueeHeight
      || prevProps.scrollSpeed !== this.props.scrollSpeed
      || prevProps.scrollDirection !== this.props.scrollDirection
    ) {
      this.handleLoad();
    }
  }


  componentWillUnmount() { 
    this.logDiagnostic('componentWillUnmount');
    window.removeEventListener('load', this.handleLoad);
    this.clearCycle();
    //this.removeDiv();
  }

  private logDiagnostic(message: string): void {
    if (this.props.enableDiagnostics === false) {
      return;
    }
    console.log('[KmMarquee] ' + message);
  }

  private clearCycle(): void {
    if (this._cycleTimer) {
      window.clearInterval(this._cycleTimer);
      this._cycleTimer = undefined;
    }
  }

  private startCycle(): void {
    this.clearCycle();
    if (this.state.messages.length <= 1) {
      return;
    }
    const durationMs = (typeof this.props.messageDuration === 'number' && this.props.messageDuration > 0 ? this.props.messageDuration : 8) * 1000;
    this._cycleTimer = window.setInterval(() => {
      this.setState((prevState: IKmMarqueeState) => {
        const nextIndex = prevState.messages.length > 0 ? (prevState.currentIndex + 1) % prevState.messages.length : 0;
        return { currentIndex: nextIndex };
      }, () => { this.handleLoad(); });
    }, durationMs);
  }

  private async loadMessages(): Promise<void> {
    if (!this.props.listName || !this.props.messageField) {
      this.clearCycle();
      this.setState({ messages: [], currentIndex: 0 }, () => { this.handleLoad(); });
      return;
    }

    try {
      this.logDiagnostic('Loading messages from list "' + this.props.listName + '", field "' + this.props.messageField + '"');
      const webUrl = this.props.spfxContext.pageContext.web.absoluteUrl.replace(/\/$/, '');
      const fieldName = escapeODataText(this.props.messageField);
      const url = webUrl + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/items"
        + "?$select=" + encodeURIComponent(this.props.messageField) + ",ID&$orderby=ID&$top=100";
      let response = await this.props.spfxContext.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) {
        response = await this.props.spfxContext.spHttpClient.get(url, SPHttpClient.configurations.v1, {
          headers: { Accept: 'application/json;odata=nometadata' }
        });
      }
      if (!response.ok) {
        throw new Error('Request failed. HTTP ' + String(response.status) + ' ' + response.statusText);
      }

      const data = await response.json();
      const items = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
      const messages = items
        .map((item: any) => { return item[fieldName] !== undefined && item[fieldName] !== null ? String(item[fieldName]) : ''; })
        .filter((message: string) => { return message.trim().length > 0; });

      this.logDiagnostic('Loaded messages successfully. Count=' + String(messages.length));
      this.setState({ messages: messages, currentIndex: 0 }, () => {
        this.startCycle();
        this.handleLoad();
      });
    } catch (error) {
      console.error('[KmMarquee] Failed to load messages from list "' + this.props.listName + '": ' + (error && error.message ? error.message : String(error)));
      this.setState({ messages: [], currentIndex: 0 }, () => { this.handleLoad(); });
    }
  }

  private getCurrentMessage(): string {
    if (this.state.messages.length > 0) {
      return this.state.messages[this.state.currentIndex] || '';
    }
    return this.props.description || '';
  }

  private handleLoad() {
    const topMarqueeParent: HTMLElement = this.getParentElement(strings.ParentElement);
    this.logDiagnostic('Top Marquee Element:' + topMarqueeParent);
    this.logDiagnostic('marqueeActive:' + this.props.marqueeActive);
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
          this.logDiagnostic('WARNING: empty fallback list supplied; defaulting to document body');
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
              console.error('[KmMarquee] ' + err);
          }
      });
      // return the proper parent, or the body element if none of the fallbacks succeed
      return parent || document.body;
  }

  public removeDiv() {    
    this.logDiagnostic('In removeDiv');
    const marqueeDiv: HTMLElement = document.getElementById('marqueeDiv') as HTMLElement;
    const marqueeSpan: HTMLElement=document.getElementById('marqueeSpan') as HTMLElement;
    this.logDiagnostic('marqueeDiv found:' + marqueeDiv);
    // check for existing marqueeDiv and span and remove
    if (marqueeDiv) {marqueeDiv.remove();}
    if(marqueeSpan) {marqueeSpan.remove();}
    } 
   
 public insertDiv(marqueeParent: HTMLElement) {    
    const displayText: string = this.getCurrentMessage();
    this.logDiagnostic('In insertDiv');
    let marqueeDiv: HTMLElement = document.getElementById('marqueeDiv') as HTMLElement;
    let marqueeSpan: HTMLElement= document.getElementById('marqueeSpan') as HTMLElement;
    // check for existing marqueeDiv
    if (!marqueeDiv) {
      marqueeDiv = document.createElement('div');
      marqueeDiv.id = 'marqueeDiv';
      marqueeDiv.className = 'bar';
      marqueeParent.insertAdjacentElement('beforebegin', marqueeDiv);
      const marqueeSpan2 = document.createElement('span');
      marqueeSpan2.className='bar_content';
      marqueeSpan2.id='marqueeSpan';
      marqueeDiv.appendChild(marqueeSpan2);
      marqueeSpan = marqueeSpan2;
      this.forceUpdate();
    }
    this.applyStyles(marqueeDiv, marqueeSpan);
    marqueeSpan.textContent = displayText;
  }

  private applyStyles(marqueeDiv: HTMLElement, marqueeSpan: HTMLElement): void {
    marqueeDiv.style.color = this.props.marqueeTextColor;
    marqueeDiv.style.backgroundColor = this.props.marqueeBackColor;
    marqueeDiv.style.minHeight = this.props.marqueeHeight || '32px';
    marqueeDiv.style.display = 'flex';
    marqueeDiv.style.alignItems = 'center';
    marqueeDiv.style.boxSizing = 'border-box';

    marqueeSpan.style.fontFamily = this.props.fontFamily || '"Segoe UI", sans-serif';
    marqueeSpan.style.fontSize = this.props.fontSize || '13px';
    marqueeSpan.style.fontStyle = this.props.fontStyle || 'normal';
    marqueeSpan.style.fontWeight = this.props.fontBold !== false ? 'bold' : 'normal';

    const speed = typeof this.props.scrollSpeed === 'number' && this.props.scrollSpeed > 0 ? this.props.scrollSpeed : 50;
    const isRightward = this.props.scrollDirection === 'right';
    marqueeSpan.style.animationDuration = speed + 's';
    marqueeSpan.style.animationTimingFunction = 'linear';
    marqueeSpan.style.animationIterationCount = 'infinite';
    marqueeSpan.style.animationName = isRightward ? 'moveRight' : 'moveLeft';
    marqueeSpan.style.transform = isRightward ? 'translateX(-100%)' : 'translateX(100%)';
  }

  public openPane(): void {
    this.props.spfxContext.propertyPane.open();
  }

  public render(): React.ReactElement<IKmMarqueeProps> {
    if (DisplayMode.Edit===this.props.dMode || (this.props.description===undefined && !this.props.listName)) {
      return (
        <div className={styles.kmMarquee}>
          <div className="col">
            <p className="$ms-color-themePrimary ms-font-xl">Marquee Web Part</p>
            <button className={styles.button} onClick={this.openPane.bind(this)}>Click here to configure</button>
          </div>
        </div>
      );
    }

    if (this.getCurrentMessage().length>0 && DisplayMode.Read===this.props.dMode) {
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
