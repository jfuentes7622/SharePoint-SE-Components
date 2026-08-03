import * as React from 'react';
import * as strings from 'TilesWebPartStrings';
import styles from './Tiles.module.scss';
import { Tile } from './Tile/Tile';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import { ITilesProps } from './ITilesProps';

/* import {
  Logger,
  ConsoleListener,
  LogLevel
} from "@pnp/logging"; */

import { Logger, ConsoleListener, LogLevel } from 'sp-pnp-js';

declare global {
  interface Window { 
       iskmActiveLog: boolean;
  }
}

const LOG_SOURCE: string = 'KM Tiles- ';
//import PnPTelemetry from "@pnp/telemetry-js";

export class Tiles extends React.Component<ITilesProps, {}> {

  public render(): React.ReactElement<ITilesProps> {
    Logger.subscribe(new ConsoleListener());
    if (!window.iskmActiveLog || window.iskmActiveLog===undefined) {
      Logger.activeLogLevel = LogLevel.Error;
    }
    else {Logger.activeLogLevel=LogLevel.Info;}

   // const telemetry = PnPTelemetry.getInstance();
    //telemetry.optOut();

    Logger.write(LOG_SOURCE + 'In Tiles Render', LogLevel.Info);
    Logger.write(LOG_SOURCE + 'this.props.title:' + this.props.title);
    return (
      <div className={styles.tiles}>
        <WebPartTitle displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.fUpdateProperty} />

        {
          this.props.collectionData && this.props.collectionData.length > 0 ? (
            <div className={styles.tilesList}>
              {
                this.props.collectionData.map((tile, idx) =>
                  <Tile key={idx} item={tile} height={this.props.tileHeight} tileEffect={this.props.tileEffect} />)
              }
            </div>
          ) : (
              <Placeholder
                iconName='Edit'
                iconText={strings.noTilesIconText}
                description={strings.noTilesConfigured}
                buttonLabel={strings.noTilesBtn}
                onConfigure={this.props.fPropertyPaneOpen} />
            )
        }
      </div>
    );
  }
}
