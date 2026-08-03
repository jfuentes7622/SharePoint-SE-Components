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

// import {
//   Logger,
//   ConsoleListener,
//   LogLevel
// } from "@pnp/logging";

import { Logger, LogLevel } from 'sp-pnp-js';

declare global {
  interface Window { 
       iskmActiveLog: boolean;
  }
}

const LOG_SOURCE: string = 'Carousel- ';


export default class IskmCarouselController extends React.Component<ICarouselProps, IIskmCarouselControllerState> {
    private _slides: Array<SlideItemModel> = new Array<SlideItemModel>();
   
    constructor(props: ICarouselProps) {
        super(props);
        this.state = { imageData: [],
                        dataLoaded: false };

        this.getListData(this.props.carouselSlideLibrary);
       window.addEventListener('hashchange', (e) => { this.getListData(this.props.carouselSlideLibrary); });
    }

    
   private getListData(slideLib:any): void {
    
    Logger.write(LOG_SOURCE + 'SlideLibrary in GetSlides:' + slideLib, LogLevel.Info);
    Promise.all([
        this.props.recSvc.GetSlides(slideLib)
        .then(iData => { this._slides = iData; 
                        return { list: 'GetSlides', data: iData }; })
        .catch(e => { 
          Logger.write(LOG_SOURCE + e, LogLevel.Error);
        })
      ])
      .then(result => { 
         this.DownloadImages((this._slides).sort((a,b) => a.slideNumber - b.slideNumber).map(d => d.slideImgUrl))
        .then(data => {
            this.setState({ imageData: data, 
                            dataLoaded:true 
                    });
        });        
      })
      .catch(e => {
        Logger.write(LOG_SOURCE+':GetListDataFailed' + e, LogLevel.Error);
      });
    //this.context;
  }  


  public componentWillReceiveProps(newComponentProps: ICarouselProps): void {
    
      if (newComponentProps.carouselSlideLibrary!==undefined) {
     
      this.getListData(newComponentProps.carouselSlideLibrary);

      /* tslint:disable-next-line:no-unused-expression */
      window.addEventListener('hashchange', (e) => { this.getListData(newComponentProps.carouselSlideLibrary); });

    }
  }  
    
    public openPane(): void {
        this.props.spfxContext.propertyPane.open();
      }

    /* public render(): React.ReactElement<ICarouselProps> {
 
      if (!window.iskmActiveLog || window.iskmActiveLog==undefined) {
        Logger.activeLogLevel = LogLevel.Error;
      }
      else {Logger.activeLogLevel=LogLevel.Info;}

 
        if (!this.props.carouselSlideLibrary) {
            return (
              <div className={styles.iskmCarousel}>
              <div className={styles.container}>
                <div className={styles.column}>
                  <div className={styles.infoRow}><h4>Carousel Web Part</h4></div>
                  <div className={styles.infoRow}>Requirements</div>
                  <div className={styles.infoRow}> - Picture library on site with columns:</div>                 
                  <div className={styles.infoRow}>1) Title (Single Line of Text)</div>
                  <div className={styles.infoRow}>2) SlideOrder (Number)</div>
                  <div className={styles.infoRow}>3) LinkTarget (Multiple Lines of Text)</div>
                  <div className={styles.infoRow}>4) Expiration (Date/Time) - Required</div>
                  <div className={styles.infoRow}>5) StartDate (Date/Time) - Required</div>
                  <div className={styles.infoRow}>6) Display (Dropdown Yes/No) - Required</div>
                  <div className={styles.infoRow}><button className={styles.button} onClick={this.openPane.bind(this)}>CLICK HERE TO CONFIGURE</button></div>
                </div>
              </div>
              </div>
            );
          }
       
        if (this.props.carouselSlideLibrary) {
             return(
            <div className={styles.carouselDiv}>
            {
                this.state.imageData.length > 0 && 
                <Carousel autoPlay={true} infiniteLoop stopOnHover
                    showStatus={false} 
                    showThumbs={false}
                    onClickItem={this._slideClicked.bind(this)}
                    width={this.props.carouselWidth.toString() + 'px'} 
                    dynamicHeight={(this.props.carouselHeight > 0?false:true)}
                    interval={parseInt(this.props.carouselSlideInterval, 10)} 
                    transitionTime={parseInt(this.props.carouselTransitionInterval, 10)} >
                    {
                        this.state.imageData.map(d => {
                            const guid = Guid.create().toString();
                            return ( <div key={guid}><img src={d} id={guid} /></div> )
                        })
                    }
                </Carousel>
            }
            </div>
        ); 
     }
     return (
      <div></div>
     )
    } */

     public render(): React.ReactElement<ICarouselProps> {

  if (!window.iskmActiveLog || window.iskmActiveLog === undefined) {
    Logger.activeLogLevel = LogLevel.Error;
  } else {
    Logger.activeLogLevel = LogLevel.Info;
  }

  if (!this.props.carouselSlideLibrary) {
    return (
      <div className={styles.iskmCarousel}>
        <div className={styles.container}>
          <div className={styles.column}>
            <div className={styles.infoRow}><h4>Carousel Web Part</h4></div>
            <div className={styles.infoRow}>Requirements</div>
            <div className={styles.infoRow}> - Picture library on site with columns:</div>
            <div className={styles.infoRow}>1) Title (Single Line of Text)</div>
            <div className={styles.infoRow}>2) SlideOrder (Number)</div>
            <div className={styles.infoRow}>3) LinkTarget (Multiple Lines of Text)</div>
            <div className={styles.infoRow}>4) Expiration (Date/Time) - Required</div>
            <div className={styles.infoRow}>5) StartDate (Date/Time) - Required</div>
            <div className={styles.infoRow}>6) Display (Dropdown Yes/No) - Required</div>
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
    const dynamicHeight = (this.props.carouselHeight > 0 ? false : true); // same logic as before
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
      width: this.props.carouselWidth.toString() + 'px'
      // If you need a fixed height when carouselHeight > 0, you can also add:
      // height: this.props.carouselHeight > 0 ? this.props.carouselHeight : undefined
    };

    return (
      <div className={styles.carouselDiv}>
        {this.state.imageData.length > 0 &&
          <div style={containerStyle}>
            <Slider {...settings}>
              {this.state.imageData.map(d => {
                const guid = Guid.create().toString();
                return (
                  <div key={guid}>
                    <img
                      src={d}
                      id={guid}
                      style={{ width: '100%', display: 'block' }}
                      // onClickItem -> add onClick to slide content
                      onClick={this._slideClicked.bind(this)}
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
                thisImage.onload = (e) => {
                    Logger.write(LOG_SOURCE + 'GetImages IMAGE ONLOAD:' + imgSrc + ',' + thisImage.width + ',' + thisImage.height, LogLevel.Info);
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

                thisImage.onerror = (e) => {
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
            const maxWidth = this.props.carouselWidth;
            const maxHeight = (this.props.carouselHeight > 0)?this.props.carouselHeight: canvases.reduce((t: number, n: HTMLCanvasElement) => {
                const thisCanvasHeight = (this.props.carouselWidth / n.width) * n.height;
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
    private _slideClicked(index: number, item: React.ReactNode): void {
         Logger.write(LOG_SOURCE +  `Slides in SlideClick :` + this._slides, LogLevel.Info);
        const currentSlide: SlideItemModel = this._slides[index];
        if(currentSlide.slideNavigationUrl) { window.open(currentSlide.slideNavigationUrl, '_blank'); }
    }
}

