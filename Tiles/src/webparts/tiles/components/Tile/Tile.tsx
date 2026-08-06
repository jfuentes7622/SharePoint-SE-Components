import * as React from 'react';
import styles from './Tile.module.scss';
import { ITileProps } from './ITileProps';

export class Tile extends React.Component<ITileProps, {}> {
  // Map hover transition type to CSS class
  private getTransitionClass(transition?: string): string {
    const transitionMap: { [key: string]: string } = {
      'solid': styles.tileOverSolid,
      'fade': styles.tileOverFade,
      'slide-top': styles.tileOverSlideTop,
      'slide-bottom': styles.tileOverSlideBottom,
      'slide-left': styles.tileOverSlideLeft,
      'slide-right': styles.tileOverSlideRight
    };
    return transitionMap[transition || 'fade'] || styles.tileOverFade;
  }

  private normalizePosition(position?: string): 'top' | 'bottom' | 'left' | 'right' {
    const normalized = (position || '').toLowerCase();
    if (normalized === 'bottom' || normalized === 'left' || normalized === 'right') {
      return normalized;
    }
    return 'top';
  }

  public render(): React.ReactElement<ITileProps> {
    const p = this.props;
    const item = p.item;

    // --- Effective values (tile override wins over global default) ---
    const tileBg     = item.color || p.backgroundColor || '#8A1717';
    const hoverBg    = p.hoverSameAsBackground !== false
      ? tileBg
      : (p.hoverColor || '#6a1010');
    const textColor  = p.textColor || '#FAFAFA';
    const fontFamily = p.fontFamily && p.fontFamily !== 'inherit' ? p.fontFamily : 'inherit';
    const fontStyle  = p.fontStyle || 'normal';
    const fontWeight = p.fontBold ? 700 : 400;
    const fontSize   = p.fontSize  ? `${p.fontSize}px`  : '14px';

    const shapeMap: { [key: string]: string } = {
      round:   '9999px',
      rounded: '10px',
      squared: '0px'
    };
    const borderRadius = shapeMap[p.tileShape || 'rounded'];

    const tileWidth  = p.tileWidth  ? `${p.tileWidth}px`  : '140px';
    const tileHeight = p.tileHeight ? `${p.tileHeight}px` : '140px';

    // CSS custom properties drive all dynamic styles inside the SCSS
    const cssVars = {
      '--tile-bg':       tileBg,
      '--hover-bg':      hoverBg,
      '--text-color':    textColor,
      '--border-radius': borderRadius,
      '--font-family':   fontFamily,
      '--font-style':    fontStyle,
      '--font-weight':   String(fontWeight),
      '--font-size':     fontSize
    } as React.CSSProperties;

    const tileStyle: React.CSSProperties = {
      ...cssVars,
      flexBasis: tileWidth,
      height:    tileHeight
    };

    const linkClassName = item.imageUrl
      ? `${styles.tileLink} ${styles.tileLinkWithImage}`
      : styles.tileLink;

    const transitionClass = this.getTransitionClass(p.hoverTransition);
    const hasHoverImage = !!item.hoverImageUrl;
    const hoverPosition = this.normalizePosition(item.hoverImagePosition || item.imagePosition);
    const hoverImageOnly = !!item.hoverImageOnly && hasHoverImage;

    const hoverOverlayClasses: string[] = [styles.tileOver, transitionClass];
    if (hasHoverImage) {
      hoverOverlayClasses.push(styles.tileOverWithImage);
    }

    const hoverInnerClasses: string[] = [styles.tileInner];
    if (hoverPosition === 'top' || hoverPosition === 'bottom') {
      hoverInnerClasses.push(styles.tileInnerVertical);
    } else {
      hoverInnerClasses.push(styles.tileInnerHorizontal);
    }

    const hoverContentClasses: string[] = [styles.tileContent, styles.tileContentWithImage];
    if (hoverPosition === 'top') {
      hoverContentClasses.push(styles.tileContentTop);
    } else if (hoverPosition === 'bottom') {
      hoverContentClasses.push(styles.tileContentBottom);
    } else if (hoverPosition === 'left') {
      hoverContentClasses.push(styles.tileContentLeft);
    } else {
      hoverContentClasses.push(styles.tileContentRight);
    }

    let hoverContent: React.ReactNode = item.description;
    if (hoverImageOnly && item.hoverImageUrl) {
      hoverContent = (
        <div className={styles.imageOnlyContainer}>
          <img src={item.hoverImageUrl} alt={item.title} />
        </div>
      );
    } else if (hasHoverImage && item.hoverImageUrl) {
      hoverContent = (
        <div className={hoverInnerClasses.join(' ')}>
          <div className={`${styles.tileImage} ${styles[hoverPosition]}`}>
            <img src={item.hoverImageUrl} alt={item.title} />
          </div>
          <div className={hoverContentClasses.join(' ')}>
            <div className={styles.tileOverText}>{item.description}</div>
          </div>
        </div>
      );
    }

    // ── imageOnly: image fills full tile, no text ──────────────────
    if (item.imageOnly && item.imageUrl) {
      return (
        <div className={styles.tile} style={tileStyle}>
          <a href={item.url}
            className={linkClassName}
            target={item.target}
            data-interception="off"
            title={item.title}>
            <div className={styles.imageOnlyContainer}>
              <img src={item.imageUrl} alt={item.title} />
            </div>
            <div className={hoverOverlayClasses.join(' ')}>{hoverContent}</div>
          </a>
        </div>
      );
    }

    // ── Normal tile (optional inline image at a given position) ────
    const position = this.normalizePosition(item.imagePosition);
    const innerClasses: string[] = [styles.tileInner];
    if (item.imageUrl) {
      if (position === 'top' || position === 'bottom') {
        innerClasses.push(styles.tileInnerVertical);
      } else {
        innerClasses.push(styles.tileInnerHorizontal);
      }
    }

    const contentClasses: string[] = [styles.tileContent];
    if (item.imageUrl) {
      contentClasses.push(styles.tileContentWithImage);
      if (position === 'top') {
        contentClasses.push(styles.tileContentTop);
      } else if (position === 'bottom') {
        contentClasses.push(styles.tileContentBottom);
      } else if (position === 'left') {
        contentClasses.push(styles.tileContentLeft);
      } else {
        contentClasses.push(styles.tileContentRight);
      }
    }

    return (
      <div className={styles.tile} style={tileStyle}>
        <a href={item.url}
          className={linkClassName}
          target={item.target}
          data-interception="off"
          title={item.title}>
          <div className={innerClasses.join(' ')}>
            {item.imageUrl && (
              <div className={`${styles.tileImage} ${styles[position]}`}>
                <img src={item.imageUrl} alt={item.title} />
              </div>
            )}
            <div className={contentClasses.join(' ')}>
              <div className={styles.tileTitle}>{item.title}</div>
            </div>
          </div>
          <div className={hoverOverlayClasses.join(' ')}>{hoverContent}</div>
        </a>
      </div>
    );
  }
}
