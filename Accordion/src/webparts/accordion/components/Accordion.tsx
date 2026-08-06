import * as React from 'react';
import styles from './Accordion.module.scss';
import { IAccordionProps } from './IAccordionProps';
import { escape } from '@microsoft/sp-lodash-subset';

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

const LOG_SOURCE: string = 'Accordion - ';

export interface ASPList {
  aHeader: string;
  aContent: string;
  aId: string;
}

export interface IReactGetItemsState {
  items: ASPList[];
  itemN: string;
  itemC: string;
}


export default class AccordionReact extends React.Component<IAccordionProps, IReactGetItemsState> {

  public constructor(props: IAccordionProps, state: IReactGetItemsState) {
    super(props);
    this.state = {
      items: [],
      itemN: '',
      itemC: ''

    };
  }


  public componentDidUpdate(prevProps: IAccordionProps, prevState: IReactGetItemsState): void {
    //if properties have changes bind it
    this.logDiagnostic("componentDidUpdate() called");
    if (this.props.itemContent !== prevProps.itemContent || this.props.listName !== prevProps.listName || this.props.itemName !== prevProps.itemName || this.props.optionChoice !== prevProps.optionChoice) {
        //use this code if using the additional states itemN and itemC and getColumnRealName function.
     /*  this.getColumnRealName(this.props.listName, this.props.itemName).then(val => {
        console.log('valN from getColumnRealname:' + val);
        this.setState({ itemN: val });
      });

      this.getColumnRealName(this.props.listName, this.props.itemContent).then(val => {
        console.log('valC from getColumnRealname:' + val);
        this.setState({ itemC: val });
      });
      if (this.props.listName != undefined && this.state.itemC != undefined && this.state.itemN != undefined) {
        this.buildAccordion(this.props.listName, this.state.itemN, this.state.itemC);
      } */
      this.buildAccordion(this.props.listName, this.props.itemName, this.props.itemContent);
      this.render();
    }
    if (this.props.overrideCssUrl !== prevProps.overrideCssUrl) {
      this._applyOverrideStylesheet();
    }
  }

  public componentDidMount() {
   this.logDiagnostic('componentDidMount() called');
   this._applyOverrideStylesheet();

   /*  this.getColumnRealName(this.props.listName, this.props.itemName).then(val => {
      console.log('valN from getColumnRealname:' + val);
      this.setState({ itemN: val });
    });
    this.getColumnRealName(this.props.listName, this.props.itemContent).then(val => {
      console.log('valC from getColumnRealname:' + val);
      this.setState({ itemC: val });
    });
    if (this.props.listName != undefined && this.state.itemC != undefined && this.state.itemN != undefined) {
      this.buildAccordion(this.props.listName, this.state.itemN, this.state.itemC);
    } */
    if (this.props.listName!=='' && this.props.itemName!=='' && this.props.itemContent!=='' && this.props.optionChoice!=='') {
    this.buildAccordion(this.props.listName, this.props.itemName, this.props.itemContent);
    this.render();}
  }



  public fetchLists(url: string): Promise<any> {
    this.logDiagnostic('fetchLists() requesting url: ' + url);
    return this.props.spfxContext.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        }
        else {
          console.error(LOG_SOURCE+ "Failed to get url:" + url + ". Error=" + response.statusText);
          return null;
        }
      }
      );
  }

  public async getColumnRealName(listTitle: string, displName: string): Promise<string> {
    //let displayName=displName.replace(/ /g, "_x0020_");
    //this function not used - but retained for reference purposes
    let url = this.props.spfxContext.pageContext.web.absoluteUrl + "/_api/web/lists/GetById('" + listTitle + "')/fields?$select InternalName&$filter=(Title eq '" + displName + "') and (Hidden eq false)";
    this.logDiagnostic("getColumnRealName() url: " + url);
    return await this.fetchLists(url).then(val => {
      this.logDiagnostic("getColumnRealName() resolved InternalName: " + val.value[0].InternalName);
      return Promise.resolve(val.value[0].InternalName);
    });
  }


  public buildAccordion(listTitle: string, colum1: string, colum2: string): void {
    if (colum1 !== '' && colum2 !== '' && listTitle !== '') {
      let column1 = colum1.replace(/ /g, "_x0020_");
      let column2 = colum2.replace(/ /g, "_x0020_");
      let url = this.props.spfxContext.pageContext.web.absoluteUrl + "/_api/web/lists/GetById('" + listTitle + "')/Items?$select=" + column1 + "," + column2 + ",ID";
      this.logDiagnostic("buildAccordion() url: " + url);
      this.fetchLists(url).then((response) => {
        let itemsArray: Array<ASPList> = new Array<ASPList>();
        response.value.map((accContents: any) => {
          let wtag = this.props.spfxContext.webPartTag;
          let uniqId = wtag + "_" + accContents.ID;
          let accHead=accContents[column1];
          if (accHead==null) {
            accHead='';
          }         
          let accCont=accContents[column2];
          if (accCont==null) {
            accCont='';
          }
          else {
            accCont= accContents[column2].replace(/<\/?p[^>]*?>/gi, '');
          }
          this.logDiagnostic("buildAccordion() built item uniqueId=" + uniqId + ", header=" + accHead);
          itemsArray.push({ aHeader: accHead, aContent: accCont, aId: uniqId });
        });
        
        this.logDiagnostic("buildAccordion() completed. Count=" + itemsArray.length);
        this.setState({ items: itemsArray } as IReactGetItemsState);

      // this.setState({ items: itemsArray });
      });
    }
  }

  public openPane(): void {
    this.props.spfxContext.propertyPane.open();
  }

  private logDiagnostic(message: string): void {
    if (this.props.enableDiagnostics === false) {
      return;
    }

    console.log('[Accordion] ' + message);
  }

  private _applyOverrideStylesheet(): void {
    if (this.props.overrideCssUrl) {
      this.logDiagnostic('Applying override stylesheet: ' + this.props.overrideCssUrl);
      const existingLink: HTMLElement = document.getElementById('accordion-component-override-css');
      if (existingLink) {
        existingLink.remove();
      }
      const link: HTMLLinkElement = document.createElement('link');
      link.id = 'accordion-component-override-css';
      link.rel = 'stylesheet';
      link.href = this.props.overrideCssUrl;
      document.head.appendChild(link);
    }
  }

  private _getHeaderStyle(): React.CSSProperties {
    return {
      backgroundColor: this.props.headerBackgroundColor,
      color: this.props.headerTextColor,
      fontWeight: this.props.headerFontBold ? 'bold' : 'normal',
      fontFamily: this.props.fontFamily,
      fontStyle: this.props.fontStyle as React.CSSProperties['fontStyle']
    };
  }

  private _getContentStyle(): React.CSSProperties {
    return {
      backgroundColor: this.props.contentBackgroundColor,
      color: this.props.contentTextColor,
      fontWeight: this.props.contentFontBold ? 'bold' : 'normal',
      fontFamily: this.props.fontFamily,
      fontStyle: this.props.fontStyle as React.CSSProperties['fontStyle']
    };
  }

  public render(): React.ReactElement<IAccordionProps> {
    this.logDiagnostic("render() optionChoice=" + this.props.optionChoice + ", listName=" + this.props.listName + ", itemContent=" + this.props.itemContent + ", itemName=" + this.props.itemName);
   // this.buildAccordion(this.props.listName, this.props.itemName, this.props.itemContent);
 /*    if (this.props.itemContent=='' || this.props.optionChoice=='' || this.props.listName==''|| this.props.itemName=='') {
      return (
        <div className={styles.accordionReact}>
          <div className="col">
            <p className="$ms-color-themePrimary ms-font-xl">Accordion Web Part</p>
            <button className={styles.button} onClick={this.openPane.bind(this)}>Click here to configure</button>
          </div>
        </div>
      );
    }
 */
    if (this.props.optionChoice == 'single' &&  this.props.listName!=='' && this.props.itemName!=='' && this.props.itemContent!=='') {
      const headerStyle: React.CSSProperties = this._getHeaderStyle();
      const contentStyle: React.CSSProperties = this._getContentStyle();
      return (
        <div className={styles.accordion} >
          {this.state.items.map((item, key) => {
            return (<div className={styles.row} key={key}>
              <div className="row">
                <div className="col">
                  <div className="tabs">
                    <div className="tab"><input className="input" type="radio" id={item.aId} name="rd"></input><label className="tab-label" style={headerStyle} htmlFor={item.aId}>{item.aHeader}</label>
                      <div className="tab-content" style={contentStyle}><div dangerouslySetInnerHTML={{ __html: item.aContent }} /></div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
            );
          })
          }
          <div className="row">
            <div className="col">
              <div className="tabs">
                <div className="tab"><input className="input" type="radio" id="rd3" name="rd"></input>
                  <label htmlFor="rd3" className="tab-close">Close others &times;</label></div>
              </div>
            </div>
          </div>

        </div>);
    }
    else if (this.props.optionChoice == 'multiple' &&  this.props.listName!=='' && this.props.itemName!=='' && this.props.itemContent!=='') {
      const headerStyle: React.CSSProperties = this._getHeaderStyle();
      const contentStyle: React.CSSProperties = this._getContentStyle();
      return (
        <div className={styles.accordion} >
          {this.state.items.map( (item, key) => {
            return (<div className={styles.row} key={key}>
              <div className="row">
                <div className="col">
                  <div className="tabs">
                    <div className="tab">
                      <input className="input" type="checkbox" id={item.aId}></input><label className="tab-label" style={headerStyle} htmlFor={item.aId}>{item.aHeader}</label>
                      <div className="tab-content" style={contentStyle}><div dangerouslySetInnerHTML={{ __html: item.aContent }} /></div>
                    </div>
                  </div>
                </div>
              </div>
            </div>);
          })}
        </div>
      );
    }
     else {
      return (
        <div className={styles.accordion}>
          <div className="col">
            <p className="$ms-color-themePrimary ms-font-xl">Accordion Web Part</p>
            <button className={styles.button} onClick={this.openPane.bind(this)}>Click here to configure</button>
          </div>
        </div>
      );
     }
  }
} 

