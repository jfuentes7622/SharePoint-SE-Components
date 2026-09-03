import * as React from 'react';
import * as strings from 'GridControlWebPartStrings';

export interface IGridRichTextEditorProps {
  value: string;
  title: string;
  placeholder: string;
  onChange: (html: string) => void;
}

export class GridRichTextEditor extends React.Component<IGridRichTextEditorProps, {}> {
  private _editor: HTMLDivElement;

  public componentDidMount(): void {
    this.updateEditorHtml();
  }

  public componentDidUpdate(): void {
    this.updateEditorHtml();
  }

  private updateEditorHtml(): void {
    if (this._editor && this._editor.innerHTML !== String(this.props.value || '')) {
      this._editor.innerHTML = String(this.props.value || '');
    }
  }

  private applyCommand(command: string): void {
    if (!this._editor) {
      return;
    }
    this._editor.focus();
    document.execCommand(command, false, undefined);
    this.props.onChange(this._editor.innerHTML);
  }

  private handleInput(): void {
    if (this._editor) {
      this.props.onChange(this._editor.innerHTML);
    }
  }

  private renderToolbarButton(command: string, iconName: string, label: string): JSX.Element {
    return (
      <button type="button" title={label} aria-label={label} onMouseDown={(ev) => {
        ev.preventDefault();
        this.applyCommand(command);
      }}>
        <i className={'ms-Icon ms-Icon--' + iconName} aria-hidden="true"></i>
      </button>
    );
  }

  public render(): JSX.Element {
    return (
      <div className="gc-rich-text-editor" title={this.props.title}>
        <div className="gc-rich-text-toolbar" role="toolbar" aria-label={strings.RuntimeRichTextToolbar}>
          {this.renderToolbarButton('bold', 'Bold', strings.RuntimeRichTextBold)}
          {this.renderToolbarButton('italic', 'Italic', strings.RuntimeRichTextItalic)}
          {this.renderToolbarButton('underline', 'Underline', strings.RuntimeRichTextUnderline)}
          {this.renderToolbarButton('insertUnorderedList', 'BulletedList', strings.RuntimeRichTextBulletedList)}
          {this.renderToolbarButton('insertOrderedList', 'NumberedList', strings.RuntimeRichTextNumberedList)}
          {this.renderToolbarButton('justifyLeft', 'AlignLeft', strings.RuntimeRichTextAlignLeft)}
          {this.renderToolbarButton('justifyCenter', 'AlignCenter', strings.RuntimeRichTextAlignCenter)}
          {this.renderToolbarButton('justifyRight', 'AlignRight', strings.RuntimeRichTextAlignRight)}
          {this.renderToolbarButton('removeFormat', 'ClearFormatting', strings.RuntimeRichTextClearFormatting)}
        </div>
        <div
          className="gc-rich-text-content"
          contentEditable={true}
          data-placeholder={this.props.placeholder}
          ref={(element) => { this._editor = element; }}
          onInput={() => this.handleInput()}
          onBlur={() => this.handleInput()}
        ></div>
      </div>
    );
  }
}