import * as React from 'react';
import { assign } from '@microsoft/sp-lodash-subset';
import styles from './Carousel.module.scss';
import { ICarouselProps } from './ICarouselProps';

import { Guid } from 'guid-typescript';

//import { Carousel } from 'react-responsive-carousel';

import Slider from 'react-slick';
import 'slick-carousel/slick/slick.css';
import 'slick-carousel/slick/slick-theme.css';


//import "react-responsive-carousel/lib/styles/carousel.min.css";

import { IIskmCarouselControllerState } from './IIskmCarouselControllerState';
import SlideItemModel from './SlideItem';

const LOG_SOURCE: string = '[Carousel] ';
const DOMPurify: any = require('dompurify');

const CarouselArrow = (props: any): React.ReactElement<any> => {
  const isPrevious = props.direction === 'previous';
  return React.createElement('button', {
    type: 'button',
    className: styles.carouselArrow + ' ' + (isPrevious ? styles.previousArrow : styles.nextArrow),
    onClick: props.onClick,
    'aria-label': isPrevious ? 'Previous slide' : 'Next slide'
  }, isPrevious ? '\u2039' : '\u203a');
};


export default class IskmCarouselController extends React.Component<ICarouselProps, IIskmCarouselControllerState> {
    private _slides: Array<SlideItemModel> = new Array<SlideItemModel>();
  private _imageVerticalInsets: Array<number> = new Array<number>();
   
    constructor(props: ICarouselProps) {
        super(props);
        this.state = { imageData: [],
                        dataLoaded: false };

        this.getListData(this.props.carouselSlideLibrary, this.props)
          .catch((e) => { console.error(LOG_SOURCE + 'Failed to load carousel data: ' + e); });
        window.addEventListener('hashchange', () => {
          this.getListData(this.props.carouselSlideLibrary, this.props)
            .catch((e) => { console.error(LOG_SOURCE + 'Failed to load carousel data: ' + e); });
        });
    }

    private logDiagnostic(message: string): void {
      if (this.props.enableDiagnostics === false) {
        return;
      }
      console.log(LOG_SOURCE + message);
    }

    private sanitizeDescription(value: string): string {
      return DOMPurify && typeof DOMPurify.sanitize === 'function'
        ? DOMPurify.sanitize(value || '', { FORBID_TAGS: ['style'], FORBID_ATTR: ['style', 'class'] })
        : '';
    }

  private getListData(slideLib: string, renderProps: ICarouselProps): Promise<void> {

    this.logDiagnostic('SlideLibrary in GetSlides: ' + slideLib);

    return renderProps.recSvc.GetSlides(slideLib)
      .then(iData => {
        this._slides = iData.slice().sort((a, b) => a.slideNumber - b.slideNumber);
        this.logDiagnostic('GetSlides returned ' + iData.length + ' slide(s)');

        const urls = this._slides.map(d => d.slideImgUrl);

        return this.DownloadImages(urls, renderProps)
          .then(data => {
            this.logDiagnostic('DownloadImages completed, rendering ' + data.length + ' slide(s)');
            this.setState({ imageData: data,
                            dataLoaded: true
                    });
          })
          .catch((e) => {
            console.error(LOG_SOURCE + 'DownloadImages failed, using direct URLs. ' + e);
            this.setState({ imageData: urls,
                            dataLoaded: true
                    });
          });
      })
      .catch(e => {
        console.error(LOG_SOURCE + ':GetListDataFailed ' + e);
      });
  }  


  public componentWillReceiveProps(newComponentProps: ICarouselProps): void {
    
      const slideRenderingChanged = newComponentProps.carouselSlideLibrary !== this.props.carouselSlideLibrary ||
        newComponentProps.slideTitleField !== this.props.slideTitleField ||
        newComponentProps.slideDescriptionField !== this.props.slideDescriptionField ||
        newComponentProps.slideLinkField !== this.props.slideLinkField ||
        newComponentProps.carouselWidth !== this.props.carouselWidth ||
        newComponentProps.carouselHeight !== this.props.carouselHeight;

      if (newComponentProps.carouselSlideLibrary !== undefined && slideRenderingChanged) {

      this.getListData(newComponentProps.carouselSlideLibrary, newComponentProps)
        .catch((e) => { console.error(LOG_SOURCE + 'Failed to load carousel data: ' + e); });

      /* tslint:disable-next-line:no-unused-expression */
      window.addEventListener('hashchange', () => {
        this.getListData(newComponentProps.carouselSlideLibrary, newComponentProps)
          .catch((e) => { console.error(LOG_SOURCE + 'Failed to load carousel data: ' + e); });
      });

    }
  }  
    
    public openPane(): void {
        this.props.spfxContext.propertyPane.open();
      }

     public render(): React.ReactElement<ICarouselProps> {

  if (!this.props.carouselSlideLibrary) {
    return (
      <div className={styles.iskmCarousel}>
        <div className={styles.container}>
          <div className={styles.column}>
            <div className={styles.infoRow}><h4>Carousel Web Part</h4></div>
            <div className={styles.infoRow}>Requirements</div>
            <div className={styles.infoRow}> - Document/Picture library on site with image files (.jpg, .jpeg, .png, .gif, .webp, .bmp, .svg)</div>
            <div className={styles.infoRow}>Optional control columns (used when present, otherwise sensible defaults apply):</div>
            <div className={styles.infoRow}>1) SlideOrder (Number) - sort order; if omitted, slides display in library order</div>
            <div className={styles.infoRow}>2) LinkTarget or ClickLink (Single/Multiple Lines of Text) - URL to open when slide is clicked</div>
            <div className={styles.infoRow}>3) Display (Dropdown Yes/No) - set to Yes to show slide; if omitted, all slides display</div>
            <div className={styles.infoRow}>4) StartDate (Date/Time) - slide visible from this date; if omitted, no start filtering</div>
            <div className={styles.infoRow}>5) Expiration (Date/Time) - slide hidden after this date; if omitted, no expiration filtering</div>
            <div className={styles.infoRow}>
              <button className={styles.button} onClick={this.openPane.bind(this)}>
                CLICK HERE TO CONFIGURE
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  if (this.props.carouselSlideLibrary) {
    // Preserve your existing names/logic and map them to react-slick settings.
    const normalizedWidth = Number(this.props.carouselWidth) > 0 ? Number(this.props.carouselWidth) : 500;
    const normalizedHeight = Number(this.props.carouselHeight);
    const dynamicHeight = !(isFinite(normalizedHeight) && normalizedHeight > 0); // same logic as before, but tolerant of invalid input
    const intervalMs = parseInt(this.props.carouselSlideInterval, 10);     // use global parseInt for SPFx 1.5.1
    const transitionMs = parseInt(this.props.carouselTransitionInterval, 10);

    const settings: any = {
      // autoPlay -> autoplay
      autoplay: true,
      // infiniteLoop -> infinite
      infinite: true,
      // stopOnHover -> pauseOnHover
      pauseOnHover: true,
      // showStatus={false} -> no status in slick; use dots: false
      dots: false,
      // showThumbs={false} -> slick has no built-in thumbs; keep behavior as no thumbnails
      arrows: true, // optional; keeps keyboard/mouse navigation visible
      // dynamicHeight -> adaptiveHeight
      adaptiveHeight: dynamicHeight,
      // interval -> autoplaySpeed (ms)
      autoplaySpeed: isNaN(intervalMs) ? 3000 : intervalMs,
      // transitionTime -> speed (ms)
      speed: isNaN(transitionMs) ? 500 : transitionMs,
      slidesToShow: 1,
      slidesToScroll: 1,
      prevArrow: React.createElement(CarouselArrow, { direction: 'previous' }),
      nextArrow: React.createElement(CarouselArrow, { direction: 'next' })
      // width is handled by wrapping container style below
    };

    const containerStyle: React.CSSProperties = {
      width: normalizedWidth.toString() + 'px',
      maxWidth: '100%',
      margin: '0 auto',
      backgroundColor: this.props.webPartBackgroundColor || '#ffffff'
      // If you need a fixed height when carouselHeight > 0, you can also add:
      // height: this.props.carouselHeight > 0 ? this.props.carouselHeight : undefined
    };
    const webPartStyle: React.CSSProperties = {
      backgroundColor: this.props.webPartBackgroundColor || 'transparent',
      borderColor: this.props.webPartBorderColor || '#d2d0ce',
      borderWidth: Math.max(0, Number(this.props.webPartBorderWidth) || 0) + 'px',
      borderStyle: this.props.webPartBorderStyle || 'solid',
      borderRadius: Math.max(0, Number(this.props.webPartCornerRadius) || 0) + 'px',
      padding: Math.max(0, Number(this.props.webPartPadding) || 0) + 'px'
    };
    const titleStyle: any = {
      color: this.props.titleTextColor || '#323130',
      fontFamily: this.props.titleFontFamily || 'inherit',
      fontSize: Math.max(10, Number(this.props.titleFontSize) || 20) + 'px',
      fontStyle: this.props.titleFontStyle || 'normal',
      fontWeight: this.props.titleFontBold !== false ? 'bold' : 'normal',
      textAlign: this.props.titleAlignment || 'center',
      marginBottom: Math.max(0, Number(this.props.titleBottomSpacing) || 0) + 'px'
    };
    const circleDiameter = isFinite(normalizedHeight) && normalizedHeight > 0
      ? Math.min(normalizedWidth, normalizedHeight)
      : normalizedWidth;
    const imageStyle: React.CSSProperties = this.props.imageIsCircle ? {
      width: circleDiameter + 'px',
      height: circleDiameter + 'px',
      maxWidth: '100%',
      display: 'block',
      margin: '0 auto',
      objectFit: 'contain'
    } : {
      width: '100%',
      display: 'block'
    };
    const slideStyle: React.CSSProperties = {
      backgroundColor: this.props.webPartBackgroundColor || '#ffffff'
    };
    const captionStyle: any = {
      backgroundColor: this.props.slideCaptionBackgroundColor || 'rgba(0,0,0,0.65)',
      padding: Math.max(0, Number(this.props.slideCaptionPadding) || 0) + 'px'
    };
    const slideTitleStyle: any = {
      color: this.props.slideTitleColor || '#ffffff',
      fontFamily: this.props.slideTitleFontFamily || 'inherit',
      fontSize: Math.max(10, Number(this.props.slideTitleFontSize) || 22) + 'px',
      fontStyle: this.props.slideTitleFontStyle || 'normal',
      fontWeight: this.props.slideTitleFontBold !== false ? 'bold' : 'normal',
      textAlign: this.props.slideTitleAlignment || 'left'
    };
    const slideDescriptionStyle: any = {
      color: this.props.slideDescriptionColor || '#ffffff',
      fontFamily: this.props.slideDescriptionFontFamily || 'inherit',
      fontSize: Math.max(10, Number(this.props.slideDescriptionFontSize) || 14) + 'px',
      fontStyle: this.props.slideDescriptionFontStyle || 'normal',
      fontWeight: this.props.slideDescriptionFontBold === true ? 'bold' : 'normal',
      textAlign: this.props.slideDescriptionAlignment || 'left'
    };
    const slideLinkStyle: any = {
      color: this.props.slideLinkColor || '#ffffff',
      fontFamily: this.props.slideLinkFontFamily || 'inherit',
      fontSize: Math.max(10, Number(this.props.slideLinkFontSize) || 14) + 'px',
      fontStyle: this.props.slideLinkFontStyle || 'normal',
      fontWeight: this.props.slideLinkFontBold !== false ? 'bold' : 'normal',
      textAlign: this.props.slideLinkAlignment || 'left'
    };

    return (
      <div className={styles.iskmCarousel + ' ' + styles.carouselSurface} style={webPartStyle}>
        {this.props.showTitle === true &&
          <div className={styles.webPartTitle} style={titleStyle}>{this.props.title}</div>
        }
      <div className={styles.carouselDiv} style={slideStyle}>
        {this.state.imageData.length > 0 &&
          <div className={styles.carouselViewport} style={containerStyle}>
            <Slider {...settings}>
              {this.state.imageData.map((d, idx) => {
                const guid = Guid.create().toString();
                const slide = this._slides[idx];
                const imageVerticalInset = this._imageVerticalInsets[idx] || 0;
                const currentImageStyle: React.CSSProperties = assign({}, imageStyle, {
                  marginTop: -imageVerticalInset + 'px',
                  marginBottom: -imageVerticalInset + 'px'
                });
                const showSlideTitle = slide && this.props.showSlideTitle && slide.slideTitle;
                const showSlideFooter = slide && (
                  (this.props.showSlideDescription && slide.slideText) ||
                  (this.props.showSlideLink && slide.slideNavigationUrl));
                return (
                  <div key={guid} className={styles.slideFrame} style={slideStyle}>
                    {showSlideTitle &&
                      <div className={styles.slideTitlePanel} style={captionStyle}>
                        <div className={styles.slideTitle} style={slideTitleStyle}>{slide.slideTitle}</div>
                      </div>
                    }
                    <img
                      src={d}
                      id={guid}
                      className={this.props.imageIsCircle ? styles.imageCircle : styles.imageSquare}
                      style={currentImageStyle}
                      // onClickItem -> add onClick to slide content
                      onClick={() => this._slideClicked(idx)}
                      alt=""
                      role="presentation"
                    />
                    {showSlideFooter &&
                      <div className={styles.slideFooter} style={captionStyle}>
                        {this.props.showSlideDescription && slide.slideText &&
                          <div className={styles.slideDescription}
                            style={slideDescriptionStyle}
                            dangerouslySetInnerHTML={{ __html: this.sanitizeDescription(slide.slideText) } as any} />
                        }
                        {this.props.showSlideLink && slide.slideNavigationUrl &&
                          <div className={styles.slideLinkRow} style={{ textAlign: this.props.slideLinkAlignment || 'left' }}>
                            <a href={slide.slideNavigationUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                              style={slideLinkStyle}
                              onClick={(event: React.MouseEvent<HTMLAnchorElement>): void => event.stopPropagation()}>
                              {slide.slideLinkText || 'Learn more'}
                            </a>
                          </div>
                        }
                      </div>
                    }
                  </div>
                );
              })}
            </Slider>
          </div>
        }
      </div>
      </div>
    );
  }

  return (<div></div>);
}


    private async DownloadImages(imgList: Array<string>, renderProps: ICarouselProps): Promise<Array<string>> {
      this._imageVerticalInsets = new Array<number>();
        return Promise.all(imgList.map(async imgSrc => {
            return await new Promise<HTMLCanvasElement>((resolve, reject) => {
                const thisCanvas: HTMLCanvasElement = document.createElement('canvas');
                const thisCanvasContext: CanvasRenderingContext2D =  thisCanvas.getContext('2d') as any as CanvasRenderingContext2D;
                const thisImage: HTMLImageElement = document.createElement('img');
                thisImage.onload = () => {
                    this.logDiagnostic('GetImages IMAGE ONLOAD:' + imgSrc + ',' + thisImage.width + ',' + thisImage.height);
                    thisCanvas.width = thisImage.width;
                    thisCanvas.height = thisImage.height;
                   
                    thisCanvas.setAttribute('originalImgSrc', imgSrc);
                    thisCanvasContext.drawImage(thisImage, 0, 0);
                    //You can do all sorts of transforms here (filter/fill/line/etc)
                    resolve(thisCanvas);
                };

                thisImage.onerror = () => {
                    thisCanvas.width = renderProps.carouselWidth;
                    thisCanvas.height = 100;
                    thisCanvas.setAttribute('originalImgSrc', imgSrc);
                    thisCanvas.setAttribute('FailedLoad', 'true');
                    thisCanvasContext.font = '20px Arial';
                    thisCanvasContext.fillText('Bad URL:', 10, 20);
                    thisCanvasContext.fillText(thisImage.src, 10, 40);
                    resolve(thisCanvas);
                };
                thisImage.src = imgSrc;
            })
            .then(canvas => {
                return canvas;
            });
        }))
        .then((canvases: Array<HTMLCanvasElement>) => {
            // All images have loaded; start processing the data provided
            const maxWidth = Number(renderProps.carouselWidth) > 0 ? Number(renderProps.carouselWidth) : 500;
            const normalizedHeight = Number(renderProps.carouselHeight);
            const maxHeight = (isFinite(normalizedHeight) && normalizedHeight > 0) ? normalizedHeight : canvases.reduce((t: number, n: HTMLCanvasElement) => {
                const thisCanvasHeight = (maxWidth / n.width) * n.height;
                return thisCanvasHeight > t ? thisCanvasHeight : t;
            }, 0);            

            // Do something here to normalize the sizes of all loaded images and return resized image data
            return canvases.map((d, index) => {
                const tempCanvas: HTMLCanvasElement = document.createElement('canvas');

                const newDims = (d.width / (d.height / maxHeight)) > maxWidth ?
                    { width: maxWidth, height: d.height / (d.width / maxWidth) } :
                    { width: d.width / (d.height / maxHeight), height: maxHeight };

              tempCanvas.width = maxWidth;
              tempCanvas.height = renderProps.imageIsCircle ? maxHeight : newDims.height;
              this._imageVerticalInsets[index] = renderProps.imageIsCircle
                ? Math.max(0, (tempCanvas.height - newDims.height) / 2)
                : 0;

              // Put the image from the original canvas into the temp canvas
              const tempCanvasCtx: CanvasRenderingContext2D = tempCanvas.getContext('2d') as any as CanvasRenderingContext2D;

              tempCanvasCtx.drawImage(d, (tempCanvas.width - newDims.width) / 2,
                renderProps.imageIsCircle ? (tempCanvas.height - newDims.height) / 2 : 0,
                newDims.width, newDims.height);
               

                // Preserve transparent letterboxing so the live carousel background remains visible.
                return tempCanvas.toDataURL('image/png');
            });
        })
        .then(data => {
            return data;
        });
    }

    // Callback for click-action on any slide
    private _slideClicked(index: number): void {
         this.logDiagnostic('Slide clicked, index: ' + index);
        const currentSlide: SlideItemModel = this._slides[index];
        if(currentSlide && currentSlide.slideNavigationUrl) { window.open(currentSlide.slideNavigationUrl, '_blank'); }
    }
}

