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
    console.info(LOG_SOURCE + "Component Did Update");
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
  }

  public componentDidMount() {
   console.info(LOG_SOURCE+ 'Component Did Mount!');

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
    console.info(LOG_SOURCE+"getColumnName URL: " + url);
    return await this.fetchLists(url).then(val => {
      console.info(LOG_SOURCE+"value0:" + val.value[0].InternalName);
      return Promise.resolve(val.value[0].InternalName);
    });
  }


  public buildAccordion(listTitle: string, colum1: string, colum2: string): void {
    if (colum1 !== '' && colum2 !== '' && listTitle !== '') {
      let column1 = colum1.replace(/ /g, "_x0020_");
      let column2 = colum2.replace(/ /g, "_x0020_");
      let url = this.props.spfxContext.pageContext.web.absoluteUrl + "/_api/web/lists/GetById('" + listTitle + "')/Items?$select=" + column1 + "," + column2 + ",ID";
      console.info(LOG_SOURCE+"Accordian URL-Selected Items: " + url);
      this.fetchLists(url).then((response) => {
        let itemsArray: Array<ASPList> = new Array<ASPList>();
        response.value.map((accContents: any) => {
          console.info(LOG_SOURCE+"Found accContents.column1:" + accContents[column1]);
          let wtag = this.props.spfxContext.webPartTag;
          console.info(LOG_SOURCE+"WebPartTag:" + wtag);
          let uniqId = wtag + "_" + accContents.ID;
          console.info(LOG_SOURCE+"UniqueId:" + uniqId);
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
          itemsArray.push({ aHeader: accHead, aContent: accCont, aId: uniqId });
        });
        
        this.setState({ items: itemsArray } as IReactGetItemsState);

      // this.setState({ items: itemsArray });
      });
    }
  }

  public openPane(): void {
    this.props.spfxContext.propertyPane.open();
  }

  public render(): React.ReactElement<IAccordionProps> {
    console.info(LOG_SOURCE+ "optionsChoice:" + this.props.optionChoice);
    console.info(LOG_SOURCE+ "listName:" + this.props.listName);
    console.info(LOG_SOURCE+ "itemContent:" + this.props.itemContent);
    console.info(LOG_SOURCE+ "itemName:" + this.props.itemName);
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
      return (
        <div className={styles.accordion} >
          {this.state.items.map((item, key) => {
            return (<div className={styles.row} key={key}>
              <div className="row">
                <div className="col">
                  <div className="tabs">
                    <div className="tab"><input className="input" type="radio" id={item.aId} name="rd"></input><label className="tab-label" htmlFor={item.aId}>{item.aHeader}</label>
                      <div className="tab-content"><div dangerouslySetInnerHTML={{ __html: item.aContent }} /></div>
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
      return (
        <div className={styles.accordion} >
          {this.state.items.map( (item, key) => {
            return (<div className={styles.row} key={key}>
              <div className="row">
                <div className="col">
                  <div className="tabs">
                    <div className="tab">
                      <input className="input" type="checkbox" id={item.aId}></input><label className="tab-label" htmlFor={item.aId}>{item.aHeader}</label>
                      <div className="tab-content"><div dangerouslySetInnerHTML={{ __html: item.aContent }} /></div>
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

