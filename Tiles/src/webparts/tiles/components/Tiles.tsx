import * as React from 'react';
import * as strings from 'TilesWebPartStrings';
import styles from './Tiles.module.scss';
import { Tile } from './Tile/Tile';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import { DisplayMode } from '@microsoft/sp-core-library';
import { ITilesProps } from './ITilesProps';

export class Tiles extends React.Component<ITilesProps, {}> {

  private logDiagnostic(message: string): void {
    if (this.props.enableDiagnostics === false) {
      return;
    }

    console.log('[Tiles] ' + message);
  }

  private onTitleChanged(event: React.FormEvent<HTMLInputElement>): void {
    this.props.fUpdateProperty(event.currentTarget.value);
  }

  public render(): React.ReactElement<ITilesProps> {
    this.logDiagnostic('render() called. title=' + this.props.title + ', tile count=' + String(this.props.collectionData ? this.props.collectionData.length : 0));
    const tileLayout = this.props.tileLayout || 'columns';
    const columnCount = tileLayout === 'stacked'
      ? 1
      : Math.max(1, Math.min(12, Number(this.props.columnCount) || 3));
    const tileGap = Math.max(0, Math.min(100, Number(this.props.tileGap) || 0));
    const tileWidth = Math.max(80, Math.min(400, Number(this.props.tileWidth) || 140));
    const tilesListStyle: React.CSSProperties = {
      gridTemplateColumns: tileLayout === 'flow'
        ? `repeat(auto-fit, ${tileWidth}px)`
        : `repeat(${columnCount}, ${tileWidth}px)`,
      gridGap: `${tileGap}px`
    };
    const webPartStyle = {
      backgroundColor: this.props.webPartBackgroundColor || '#ffffff',
      borderColor: this.props.webPartBorderColor || '#d2d0ce',
      borderWidth: `${Math.max(0, Number(this.props.webPartBorderWidth) || 0)}px`,
      borderStyle: this.props.webPartBorderStyle || 'solid',
      borderRadius: `${Math.max(0, Number(this.props.webPartCornerRadius) || 0)}px`,
      padding: `${Math.max(0, Number(this.props.webPartPadding) || 0)}px`,
      boxSizing: 'border-box',
      '--tiles-title-color': this.props.titleTextColor || '#323130',
      '--tiles-title-font-family': this.props.titleFontFamily || 'inherit',
      '--tiles-title-font-size': `${Math.max(10, Number(this.props.titleFontSize) || 20)}px`,
      '--tiles-title-font-style': this.props.titleFontStyle || 'normal',
      '--tiles-title-font-weight': this.props.titleFontBold !== false ? 'bold' : 'normal',
      '--tiles-title-alignment': this.props.titleAlignment || 'center',
      '--tiles-title-bottom-spacing': `${Math.max(0, Number(this.props.titleBottomSpacing) || 0)}px`
    } as React.CSSProperties;
    return (
      <div className={styles.tiles} style={webPartStyle}>
        {this.props.showTitle === true && (
          <div className={styles.webPartTitle}>
            {this.props.displayMode === DisplayMode.Edit ? (
              <input type="text"
                aria-label="Web part title"
                placeholder="Add a title"
                value={this.props.title || ''}
                onChange={this.onTitleChanged.bind(this)} />
            ) : (
              <span>{this.props.title}</span>
            )}
          </div>
        )}

        {
          this.props.collectionData && this.props.collectionData.length > 0 ? (
            <div className={styles.tilesList} style={tilesListStyle}>
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
