import * as React from 'react';
import styles from '../Contacts.module.scss';

export interface IimageDisplayProps {
    imageDisplayWidth: number;
    imageDisplayLink: string;
    imageDisplayClickLink?: string;
    imageShape: string;
    imageIsDynamicHeight: boolean;
    imageDisplayHeight?: number;
}

export default class ImageDisplay extends React.Component<IimageDisplayProps, {}> {

    private imgWidth: number;
    private imgHeight: number;
    private inMemImage: HTMLImageElement = document.createElement('img');

    constructor(props: IimageDisplayProps) {
        super(props);
    }

    public render(): React.ReactElement<IimageDisplayProps> {
        const shapeClass = (this.props.imageShape === 'circle') ? styles.imageCircle
            : (this.props.imageShape === 'rounded') ? styles.imageRounded
            : styles.imageSquare;
        return (
            <div><canvas className={shapeClass} ref='canvas'/></div>
        );
    }

    public componentDidMount(): void {
        this.LoadImage();
    }

    public componentDidUpdate(): void {
        this.LoadImage();
    }

    private async getHeight(defaultImageWidth: number): Promise<number> {
        //This is to get the height in order to get the ratio for scaling the image
        if (this.props.imageDisplayLink) {
            if (this.props.imageDisplayLink !== "") {
                return await new Promise<HTMLCanvasElement>((resolve, reject) => {
                    const thisCanvas: HTMLCanvasElement = document.createElement('canvas');
                    const thisCanvasContext: CanvasRenderingContext2D = (thisCanvas.getContext('2d') as any);

                    const thisImage: HTMLImageElement = document.createElement('img');
                    thisImage.onload = (e) => {
                        // console.log('GetImages IMAGE ONLOAD: ', imgSrc, thisImage.width, thisImage.height);
                        thisCanvas.width = thisImage.width;
                        thisCanvas.height = thisImage.height;
                        thisCanvas.setAttribute('originalImgSrc', this.props.imageDisplayLink);
                        thisCanvasContext.drawImage(thisImage, 0, 0);

                        resolve(thisCanvas);
                    };

                    thisImage.onerror = (e) => {
                        thisCanvas.width = defaultImageWidth;
                        thisCanvas.height = 100;
                        thisCanvas.setAttribute('originalImgSrc', this.props.imageDisplayLink);
                        thisCanvas.setAttribute('FailedLoad', 'true');
                        thisCanvasContext.font = '20px Arial';
                        thisCanvasContext.fillText('Bad URL:', 10, 20);
                        thisCanvasContext.fillText(thisImage.src, 10, 40);

                        resolve(thisCanvas);
                    };

                    thisImage.src = this.props.imageDisplayLink;
                }).then(canvas => {
                    if (this.props.imageIsDynamicHeight) {
                        //get ratio
                        const ratio = defaultImageWidth / canvas.width;
                        return Math.ceil(canvas.height * ratio);
                    } else {
                        return this.props.imageDisplayHeight;
                    }
                }) as any;

            }
        }
        return 0;
    }

    private LoadImage(): void {
        const defaultImageWidth = (this.props.imageDisplayWidth && this.props.imageDisplayWidth > 0) ? this.props.imageDisplayWidth : 100;

        const cvs: HTMLCanvasElement = (this.refs.canvas as HTMLCanvasElement);

        // Resolves true when both the User's image is null and the default image url is null
        if (!this.props.imageDisplayLink) {
            cvs.width = defaultImageWidth;
            cvs.height=cvs.width;
            //this.getHeight(defaultImageWidth).then(val => cvs.height = val);
            return;
        }

        // Handles when a error occurs during image loading
        this.inMemImage.onerror = e => {
            cvs.width = defaultImageWidth;
            cvs.height=cvs.width;
            //this.getHeight(defaultImageWidth).then(val => cvs.height = val);
        };

        // Handles when the image has loaded sucessfully
        this.inMemImage.onload = (e) => {
            this.imgWidth = defaultImageWidth;
            this.getHeight(defaultImageWidth).then(val => {

                this.imgHeight = val;
                //set it so that image height = image width, don't do calculation
                this.imgHeight=this.imgWidth;

                cvs.width = this.imgWidth;
                cvs.height = this.imgHeight;

                const ctx: CanvasRenderingContext2D = cvs.getContext('2d') as any;
                ctx.drawImage(this.inMemImage, 0, 0, this.imgWidth, this.imgHeight);

            });
        };

        this.inMemImage.src = this.props.imageDisplayLink;
        if (this.props.imageDisplayClickLink && this.props.imageDisplayClickLink.length > 0) {
            this.inMemImage.onclick = this.newTab;
        }
    }

    private newTab(e: MouseEvent): void {
        window.open(this.props.imageDisplayClickLink, '_blank');
    }
}
