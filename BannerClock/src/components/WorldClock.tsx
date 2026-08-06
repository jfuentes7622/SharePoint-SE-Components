import * as React from 'react';
import { Clock } from './Clock';
import { ISPClockList } from '../Interfaces/ISPClockListElement';
import { SPHttpClient } from '@microsoft/sp-http';
//import { SpfxFetch } from '../Services/SpfxFetch';
import { IWorldClockProps, IWorldClockState } from '../Interfaces/IWorldClock';
import { ISPClockListElement } from '../Interfaces/ISPClockListElement';

interface IclockItem {
  Title: string;
  Timezone: string;
}

export default class WorldClock extends React.Component<IWorldClockProps, IWorldClockState> {

  //private _spfxFetch!: SpfxFetch;
  private _invalidList: boolean;
  private _absoluteUrl: string;
  private _spHttpClinet: SPHttpClient;

  constructor(props: IWorldClockProps) {
    super(props);
    this.context = this.props.context;
    this._absoluteUrl = this.props.absoluteUrl;
    this._invalidList = false;
    this._spHttpClinet = this.props.spHttpClient;
    this.state = {
      spClockList: { value: [] }
    };
  }

  private logDiagnostic(message: string): void {
    if (this.props.enableDiagnostics === false) {
      return;
    }

    console.log('[WorldClock] ' + message);
  }
  

  componentDidMount(): void {
    this.logDiagnostic('componentDidMount. spList=' + String(this.props.spList || '(none)'));
    if (this.props.spList) {
      this.getDataFromSharepoint(this.props.spList).then((res) => {
        if (res.value.length > 0) {
          this.logDiagnostic('Loaded ' + String(res.value.length) + ' clock entries from list "' + this.props.spList + '".');
          this._updateState(res);
        } else {
          console.warn('[WorldClock] List "' + this.props.spList + '" returned no usable clock entries.');
          this._invalidList = true;
          this.setState({ spClockList: { value: [] } });
        }
      })
        .catch((error) => {
          console.error('Error fetching data from SharePoint:', error);
        });
    }
  }

  componentDidUpdate(prev: IWorldClockProps): void {
    if (this.props.absoluteUrl !== prev.absoluteUrl) {
      this._absoluteUrl = this.props.absoluteUrl;
      this.logDiagnostic('absoluteUrl changed to ' + this._absoluteUrl + '; reloading clock list.');
      if (this.props.spList !== null) {
        this.getDataFromSharepoint(this.props.spList).then((res) => {
          if (res.value.length > 0) {
            this._updateState(res);
          } else {
            this._invalidList = true;
            this.setState({ spClockList: { value: [] } });
          }
        })
          .catch((error) => {
            console.error('Error fetching data from SharePoint:', error);
          });
      }
    }
    if (this.props.spList !== prev.spList) {
      this.logDiagnostic('spList prop changed from "' + prev.spList + '" to "' + this.props.spList + '"; reloading clock list.');
      this.getDataFromSharepoint(this.props.spList).then((res) => {
        if (res.value.length > 0) {
          this._updateState(res);
        } else {
          this._invalidList = true;
          this.setState({ spClockList: { value: [] } });
        }
      })
        .catch((error) => {
          console.error('Error fetching data from SharePoint:', error);
        });
    }
  }

  private _isClockType(element: any): element is IclockItem {
    return (
      typeof element === 'object' &&
      element !== null &&
      'Title' in element &&
      'Timezone' in element
    );
  }

  private _updateState(spClockList: ISPClockList): void {
    if (spClockList && spClockList.value && (spClockList.value.length > 0)) {
      if (this._isClockType(spClockList.value[0])) {
        this._invalidList = false;
        this.setState({ spClockList: spClockList });
      } else {
        this._invalidList = true;
        this.setState({ spClockList: { value: [] } });
      }
    }
  }

  private  getDataFromSharepoint(listName: string): Promise<ISPClockList> {
  //const _spfxFetch: SpfxFetch = new SpfxFetch(this.context.spHttpClient);
  this.logDiagnostic('Fetching clock list "' + listName + '" from ' + this._absoluteUrl);
  const fetchResponse: any =  this._spHttpClinet.get(
    `${this._absoluteUrl}/_api/web/lists/GetbyTitle('${listName}')/items?$orderby=SortOrder asc`,
    SPHttpClient.configurations.v1
  );

  return fetchResponse
    .then(async (response): Promise<ISPClockList> => {
      try {
        const json: any = await response.json();
        const result: ISPClockListElement[] = [];

        if (json.error) {
          console.error(`Failed to get list: '${listName}' at ${this._absoluteUrl} ; ${json.error.message}`);
        } else if (Array.isArray(json.value)) {
          json.value.forEach((listItem: ISPClockListElement) => {
            if ('Title' in listItem && 'Timezone' in listItem) {
              result.push({ Title: listItem.Title, Timezone: listItem.Timezone });
            }
          });
        }

        return { value: result };
      } catch (error) {
        console.error(`Failed to parse JSON from list: '${listName}' at ${this._absoluteUrl} ; ${error}`);
        return { value: [] };
      }
    })
    .catch((error) => {
      console.error(`SpfxFetch failed: ${error.message}`);
      return { value: [] };
    });
}

  public render(): React.ReactElement<IWorldClockProps> {

    if (this._invalidList) {
      return (
        <div className='worldClockWebPart'>
          <div className='invalidList'>Invalid World Clock List!</div>
        </div>
      );
    }
    if (this.state.spClockList.value.length === 0) {
      return (
        <div className='worldClockWebPart'>
          <div className='ms-fontWeight-bold'>Please open the property pane and choose a list with locations and timezones.</div>
        </div>
      );
    }
    if (!this._invalidList) {
      return (
        <div className='worldClockWebPart'>
          <div className='clocksGrid'>
            {this.state.spClockList.value.map((mapping: ISPClockListElement) => (
              <div key={mapping.Title} className='clockComponent'>
                <div className='clockTitle'> {mapping.Title} </div>
                <Clock
                  timeZone={mapping.Timezone}
                  digitalOnly={this.props.digitalOnly}
                  hour12={this.props.hour12}
                  displayDay={this.props.displayDay}
                />
              </div>
            ))}
          </div>
        </div>
      );
    }
    return (
      <div></div>
    );
  }
}