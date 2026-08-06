import * as React from 'react';
import * as strings from 'TilesWebPartStrings';
import styles from './Tiles.module.scss';
import { Tile } from './Tile/Tile';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import { ITilesProps } from './ITilesProps';

export class Tiles extends React.Component<ITilesProps, {}> {

  private logDiagnostic(message: string): void {
    if (this.props.enableDiagnostics === false) {
      return;
    }

    console.log('[Tiles] ' + message);
  }

  public render(): React.ReactElement<ITilesProps> {
    this.logDiagnostic('render() called. title=' + this.props.title + ', tile count=' + String(this.props.collectionData ? this.props.collectionData.length : 0));
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
                  <Tile key={idx}
                    item={tile}
                    tileEffect={this.props.tileEffect}
                    tileShape={this.props.tileShape}
                    backgroundColor={this.props.backgroundColor}
                    textColor={this.props.textColor}
                    hoverColor={this.props.hoverColor}
                    hoverSameAsBackground={this.props.hoverSameAsBackground}
                    hoverTransition={this.props.hoverTransition}
                    fontFamily={this.props.fontFamily}
                    fontStyle={this.props.fontStyle}
                    fontBold={this.props.fontBold}
                    fontSize={this.props.fontSize}
                    tileWidth={this.props.tileWidth}
                    tileHeight={this.props.tileHeight} />)
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
