import * as React from 'react';
import styles from './Contacts.module.scss';
import { IContactsProps } from './IContactsProps';
import PersonnelWidget from './PersonnelWidget';
import { PersonGroupModel } from './shared/PersonGroup';
import { Promise } from 'es6-promise';

//import PnPTelemetry from "@pnp/telemetry-js";
// import {
//   Logger,
//   ConsoleListener,
//   LogLevel
// } from "@pnp/logging";
import  { Logger, LogLevel } from 'sp-pnp-js';

const LOG_SOURCE: string = 'Contacts - ';
//const telemetry = PnPTelemetry.getInstance();

declare global {
  interface Window {
    iskmActiveLog: boolean;
  }
}

export interface IContactImageState {
  dataLoaded: boolean;
}

export default class ContactImages extends React.Component<IContactsProps, IContactImageState> {
  private _personnelList: Array<PersonGroupModel> = new Array<PersonGroupModel>();
  private _customList: Array<PersonGroupModel> = new Array<PersonGroupModel>();
  constructor(props: IContactsProps) {
    super(props);
    this.state = {
      dataLoaded: false
    };

    this.getListData();
    this.sortListData();

   // Logger.subscribe(ConsoleListener());
    if (!window.iskmActiveLog || window.iskmActiveLog === undefined) {
      Logger.activeLogLevel = LogLevel.Error;
    }
    else {
      Logger.activeLogLevel = LogLevel.Info;
    }

    //telemetry.optOut();

    Logger.write(LOG_SOURCE + 'In ContactImages.tsx constructor', LogLevel.Info);
  }

  private getListData(): void {
    //let PromisesResult: boolean = true;  
    //Logger.write('Dir:' + this.props.directorate,LogLevel.Info);
    //Logger.write('Div:' + this.props.division ,LogLevel.Info);
    Promise.all([
      this.props.recSvc.GetPersonnel(this.props.directorate, this.props.division)
        .then(p => { this._personnelList = p; return { list: 'GetPersonnel', data: p }; })
        .catch(e => { Logger.write('GetListData failed promises' + e, LogLevel.Error); })
    ])
      .then(result => {
        // result has the data returned from all of the promises...could be dumped to the console if needed
        // this._organizationInvalid = !this._orgData && this.state.dirDivBranch.directorate.length > 0;
        this.setState({
          dataLoaded: true,
        });
      })
      .catch(e => {
        Logger.write('GetListData failed promises' + e, LogLevel.Error);
      });
    //this.context;
  }

  private sortListData(): void {
    //const newList = Object.assign([], this.props.customList);
    const newList = (this.props.customList || []).slice();
    const data = this.props.recSvc.SortPersonnel(newList);
    //Logger.write('Data from SortListData:' + data,LogLevel.Info);
    this._customList = data;
  }


  public componentDidUpdate(prevProps: { directorate: string; division: string; imageWidth: number; }): void {
    Logger.write('ComponentDidUPdate', LogLevel.Info);
    //Logger.write('Custom List in ComponentDidUpdate:'+ this.props.customList,LogLevel.Info);
    //Logger.write('Prev Custom List in ComponentDidUpdate:'+ prevProps, LogLevel.Info);
    //Logger.write('propertyCheckBox:'+ this.props.propertyCheckbox,LogLevel.Info);
    //Logger.write('ImageWidth:'+ this.props.imageWidth,LogLevel.Info);
    if (((prevProps.directorate !== '' && prevProps.directorate !== this.props.directorate) || (prevProps.division !== this.props.division) || (prevProps.imageWidth !== this.props.imageWidth)) && this.props.propertyCheckbox === true) {
      this.getListData();
    }
    if (this.props.customList !== undefined && this.props.propertyCheckbox === false) {
      this.sortListData();
      this.render();
    }

  }

  public openPane(): void {
    this.props.spfxContext.propertyPane.open();
  }

  public render(): React.ReactElement<IContactsProps> {

    if ((this.props.directorate === undefined || this.props.directorate === '') && this.props.customList === undefined) {
      return (
        <div className={styles.ContactImages}>
          <div className="col">
            <p className="$ms-color-themePrimary ms-font-xl">Contacts Web Part</p>
            <button className={styles.button} onClick={this.openPane.bind(this)}>Click here to configure</button>
          </div>
        </div>
      );
    }
    if (this.props.directorate !== undefined && this.props.directorate !== '' && this.props.propertyCheckbox === true) {
      return (
        <div className={styles.landingBlock}>
          <div className={styles.landingRow}>
            <div className={(this._personnelList.length > 0) ? 'ms-Grid-col ms-sm12' : 'display:none'}>
              <PersonnelWidget PersonGroups={this._personnelList} />
            </div>
          </div></div>
      );
    }
    else {
      if (this.props.customList !== undefined && this.props.propertyCheckbox === false) {
        this.sortListData();
        //Logger.write('_customList lenght:' + this._customList.length,LogLevel.Info);
        //Logger.write('CustomList lenght:'+ this.props.customList.length,LogLevel.Info);
        return (
          <div className={styles.landingBlock}>
            <div className={styles.landingRow}>
              <div className={(this._customList.length > 0) ? 'ms-Grid-col ms-sm12' : 'display:none'}>
                <PersonnelWidget PersonGroups={this._customList} />
              </div>
            </div></div>
        );
      }
    }
    return (
      <div className={styles.ContactImages}>
        <div className="col">
          <p className="$ms-color-themePrimary ms-font-xl">Contacts Web Part</p>
          <button className={styles.button} onClick={this.openPane.bind(this)}>Click here to configure</button>
        </div>
        {}
      </div>
    );

  }
}
