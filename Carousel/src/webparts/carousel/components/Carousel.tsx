import * as React from 'react';
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


export default class IskmCarouselController extends React.Component<ICarouselProps, IIskmCarouselControllerState> {
    private _slides: Array<SlideItemModel> = new Array<SlideItemModel>();
   
    constructor(props: ICarouselProps) {
        super(props);
        this.state = { imageData: [],
                        dataLoaded: false };

        this.getListData(this.props.carouselSlideLibrary)
          .catch((e) => { console.error(LOG_SOURCE + 'Failed to load carousel data: ' + e); });
        window.addEventListener('hashchange', () => {
          this.getListData(this.props.carouselSlideLibrary)
            .catch((e) => { console.error(LOG_SOURCE + 'Failed to load carousel data: ' + e); });
        });
    }

    private logDiagnostic(message: string): void {
      if (this.props.enableDiagnostics === false) {
        return;
      }
      console.log(LOG_SOURCE + message);
    }

   private getListData(slideLib: string): Promise<void> {

    this.logDiagnostic('SlideLibrary in GetSlides: ' + slideLib);

    return this.props.recSvc.GetSlides(slideLib)
      .then(iData => {
        this._slides = iData;
        this.logDiagnostic('GetSlides returned ' + iData.length + ' slide(s)');

        const urls = (this._slides).slice().sort((a, b) => a.slideNumber - b.slideNumber).map(d => d.slideImgUrl);

        return this.DownloadImages(urls)
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
    
      if (newComponentProps.carouselSlideLibrary!==undefined && newComponentProps.carouselSlideLibrary !== this.props.carouselSlideLibrary) {

      this.getListData(newComponentProps.carouselSlideLibrary)
        .catch((e) => { console.error(LOG_SOURCE + 'Failed to load carousel data: ' + e); });

      /* tslint:disable-next-line:no-unused-expression */
      window.addEventListener('hashchange', () => {
        this.getListData(newComponentProps.carouselSlideLibrary)
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
      slidesToScroll: 1
      // width is handled by wrapping container style below
    };

    const containerStyle: React.CSSProperties = {
      width: normalizedWidth.toString() + 'px'
      // If you need a fixed height when carouselHeight > 0, you can also add:
      // height: this.props.carouselHeight > 0 ? this.props.carouselHeight : undefined
    };

    return (
      <div className={styles.carouselDiv}>
        {this.state.imageData.length > 0 &&
          <div style={containerStyle}>
            <Slider {...settings}>
              {this.state.imageData.map((d, idx) => {
                const guid = Guid.create().toString();
                return (
                  <div key={guid}>
                    <img
                      src={d}
                      id={guid}
                      className={this.props.imageIsCircle ? styles.imageCircle : styles.imageSquare}
                      style={{ width: '100%', display: 'block' }}
                      // onClickItem -> add onClick to slide content
                      onClick={() => this._slideClicked(idx)}
                      alt=""
                      role="presentation"
                    />
                  </div>
                );
              })}
            </Slider>
          </div>
        }
      </div>
    );
  }

  return (<div></div>);
}


    private async DownloadImages(imgList: Array<string>): Promise<Array<string>> {
        return Promise.all(imgList.map(async imgSrc => {
            return await new Promise<HTMLCanvasElement>((resolve, reject) => {
                const thisCanvas: HTMLCanvasElement = document.createElement('canvas');
                const thisCanvasContext: CanvasRenderingContext2D =  thisCanvas.getContext('2d') as any as CanvasRenderingContext2D;
                const thisImage: HTMLImageElement = document.createElement('img');
                thisImage.onload = () => {
                    this.logDiagnostic('GetImages IMAGE ONLOAD:' + imgSrc + ',' + thisImage.width + ',' + thisImage.height);
                      thisCanvasContext.beginPath();
                      thisCanvasContext.fillRect(0,0,thisImage.width,thisImage.height);
                      thisCanvasContext.fillStyle=this.props.carouselBackgroundColor;

                   thisCanvas.width = thisImage.width;
                    thisCanvas.height = thisImage.height;
                   
                    thisCanvas.setAttribute('originalImgSrc', imgSrc);
                    thisCanvasContext.drawImage(thisImage, 0, 0);
                    //You can do all sorts of transforms here (filter/fill/line/etc)
                    resolve(thisCanvas);
                };

                thisImage.onerror = () => {
                    thisCanvas.width = this.props.carouselWidth;
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
            const maxWidth = Number(this.props.carouselWidth) > 0 ? Number(this.props.carouselWidth) : 500;
            const normalizedHeight = Number(this.props.carouselHeight);
            const maxHeight = (isFinite(normalizedHeight) && normalizedHeight > 0) ? normalizedHeight : canvases.reduce((t: number, n: HTMLCanvasElement) => {
                const thisCanvasHeight = (maxWidth / n.width) * n.height;
                return thisCanvasHeight > t ? thisCanvasHeight : t;
            }, 0);            

            // Do something here to normalize the sizes of all loaded images and return resized image data
            return canvases.map(d => {
                const tempCanvas: HTMLCanvasElement = document.createElement('canvas');

                tempCanvas.width = maxWidth;
                tempCanvas.height = maxHeight;
               
                // Put the image from the original canvas into the temp canvas
                const tempCanvasCtx: CanvasRenderingContext2D = tempCanvas.getContext('2d') as any as CanvasRenderingContext2D;
                tempCanvasCtx.fillStyle = this.props.carouselBackgroundColor;
                tempCanvasCtx.fillRect(0, 0, maxWidth, maxHeight);

                const newDims = (d.width / (d.height / maxHeight)) > maxWidth ?
                    { width: maxWidth, height: d.height / (d.width / maxWidth) } :
                    { width: d.width / (d.height / maxHeight), height: maxHeight };


                tempCanvasCtx.drawImage(d, (tempCanvas.width - newDims.width) / 2, (tempCanvas.height - newDims.height) / 2, newDims.width, newDims.height);
               

                // Return the B64 encoded image data
                return tempCanvas.toDataURL('image/jpeg', 0.75);
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

