import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { DialogContent } from 'office-ui-fabric-react';

interface IIFrameDialogContentProps {
    close: () => void;
    url: string;
    iframeOnLoad?: (iframe: any) => void;
    width: number;
    height: number;
    titleStr: string;
}

class IFrameDialogContent extends React.Component<IIFrameDialogContentProps, {} > {
    private _iframe: any;
    constructor(props: IIFrameDialogContentProps) {
        super(props);
    }
    public render(): JSX.Element {
        return(
            <DialogContent title={this.props.titleStr} onDismiss={this.props.close} showCloseButton={true}>
                <iframe title='LoadFrame' ref={(iframe) => { this._iframe = iframe; }} onLoad={this._iframeOnLoad.bind(this)}
                style={{ width: this.props.width + 'px', height: this.props.height + 'px' }} src={this.props.url} />
            </DialogContent>
        );
    }

    private _iframeOnLoad(): void {
        try {
            this._iframe.contentWindow.frameElement.cancelPopUp = this.props.close;
        } catch (err) {
            if (err.name !== 'SecurityError') {
                throw err;
            }
        }
        if (this.props.iframeOnLoad) {
            this.props.iframeOnLoad(this._iframe);
        }
    }
}

export default class IFrameDialog extends BaseDialog {
    private url: string = '';
    private title: string = '';
    private _width: number = 600;
    private _height: number = 800;

    constructor(title: string, url: string, width?: number, height?: number) {
        super();
        this.url = url;
        this.title = title;
        if (width) { this._width = width; }
        if (height) { this._height = height; }
    }
    public render(): void {
        window.addEventListener('CloseDialog', () => { this.close(); });
        ReactDOM.render(<IFrameDialogContent titleStr={this.title} width={this._width} height={this._height} close={this.close} url={this.url} />, this.domElement);
    }
    public getConfig(): IDialogConfiguration {
        return {
            isBlocking: false
        };
    }
    protected onAfterClose(): void {
        super.onAfterClose();
        ReactDOM.unmountComponentAtNode(this.domElement);
    }
}
