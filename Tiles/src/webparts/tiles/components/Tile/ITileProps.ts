import { ITileInfo } from '../../ITileInfo';
import { TileShape, TileFontStyle, TileHoverTransition } from '../../TilesWebPart';

export interface ITileProps {
  item: ITileInfo;
  tileEffect: string;
  tileShape: TileShape;
  backgroundColor: string;
  textColor: string;
  hoverColor: string;
  hoverSameAsBackground: boolean;
  hoverTransition: TileHoverTransition;
  fontFamily: string;
  fontStyle: TileFontStyle;
  fontBold: boolean;
  fontSize: number;
  tileWidth: number;
  tileHeight: number;
}
