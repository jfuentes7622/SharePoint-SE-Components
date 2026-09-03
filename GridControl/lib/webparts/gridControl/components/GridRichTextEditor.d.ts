/// <reference types="react" />
import * as React from 'react';
export interface IGridRichTextEditorProps {
    value: string;
    title: string;
    placeholder: string;
    onChange: (html: string) => void;
}
export declare class GridRichTextEditor extends React.Component<IGridRichTextEditorProps, {}> {
    private _editor;
    componentDidMount(): void;
    componentDidUpdate(): void;
    private updateEditorHtml();
    private applyCommand(command);
    private handleInput();
    private renderToolbarButton(command, iconName, label);
    render(): JSX.Element;
}
