import * as React from 'react';
import styles from './SharePointDynamicForm.module.scss';

export interface RichTextEditorProps {
  value: string;
  onChange: (html: string) => void;
  placeholder?: string;
}

export class RichTextEditor extends React.Component<RichTextEditorProps> {
  private editorRef: HTMLDivElement;

  public componentDidMount(): void {
    this.updateEditorHtml();
  }

  public componentDidUpdate(): void {
    if (typeof document !== 'undefined' && document.activeElement === this.editorRef) {
      return;
    }
    this.updateEditorHtml();
  }

  private updateEditorHtml(): void {
    if (this.editorRef && this.editorRef.innerHTML !== String(this.props.value || '')) {
      this.editorRef.innerHTML = String(this.props.value || '');
    }
  }

  private applyCommand(command: string, value?: string): void {
    if (!this.editorRef) {
      return;
    }
    this.editorRef.focus();
    document.execCommand(command, false, value || undefined);
    this.props.onChange(this.editorRef.innerHTML);
  }

  private handleInput = () => {
    if (this.editorRef) {
      this.props.onChange(this.editorRef.innerHTML);
    }
  }

  private insertImage = () => {
    const url = prompt('Enter image URL:');
    if (url) {
      this.applyCommand('insertImage', url);
    }
  }

  private clearFormatting = () => {
    this.applyCommand('removeFormat');
  }

  public render() {
    return (
      <div className={styles.richTextEditor} dir="ltr" style={{
        border: '1px solid #d0d0d0',
        borderRadius: '4px',
        overflow: 'hidden',
        direction: 'ltr',
        backgroundColor: '#ffffff'
      }}>
        {/* Toolbar */}
        <div style={{
          display: 'flex',
          flexDirection: 'row',
          gap: '2px',
          padding: '8px',
          backgroundColor: '#f5f5f5',
          borderBottom: '1px solid #d0d0d0',
          flexWrap: 'wrap'
        }}>
          {/* Text Format */}
          <div style={{ display: 'flex', flexDirection: 'row', gap: '2px', marginRight: '4px' }}>
            <button
              type="button"
              onMouseDown={(ev) => { ev.preventDefault(); this.applyCommand('bold'); }}
              title="Bold"
              style={this.getButtonStyle()}
            >
              <strong>B</strong>
            </button>
            <button
              type="button"
              onMouseDown={(ev) => { ev.preventDefault(); this.applyCommand('italic'); }}
              title="Italic"
              style={this.getButtonStyle()}
            >
              <em>I</em>
            </button>
            <button
              type="button"
              onMouseDown={(ev) => { ev.preventDefault(); this.applyCommand('underline'); }}
              title="Underline"
              style={this.getButtonStyle()}
            >
              <u>U</u>
            </button>
            <button
              type="button"
              onMouseDown={(ev) => { ev.preventDefault(); this.applyCommand('strikeThrough'); }}
              title="Strikethrough"
              style={this.getButtonStyle()}
            >
              <s>S</s>
            </button>
          </div>

          {/* Divider */}
          <div style={{ width: '1px', backgroundColor: '#ccc', margin: '0 2px' }} />

          {/* Alignment */}
          <div style={{ display: 'flex', flexDirection: 'row', gap: '2px', marginRight: '4px' }}>
            <button
              type="button"
              onMouseDown={(ev) => { ev.preventDefault(); this.applyCommand('justifyLeft'); }}
              title="Align Left"
              style={this.getButtonStyle()}
            >
              ≡
            </button>
            <button
              type="button"
              onMouseDown={(ev) => { ev.preventDefault(); this.applyCommand('justifyCenter'); }}
              title="Align Center"
              style={this.getButtonStyle()}
            >
              ⋮
            </button>
            <button
              type="button"
              onMouseDown={(ev) => { ev.preventDefault(); this.applyCommand('justifyRight'); }}
              title="Align Right"
              style={this.getButtonStyle()}
            >
              ≡
            </button>
          </div>

          {/* Divider */}
          <div style={{ width: '1px', backgroundColor: '#ccc', margin: '0 2px' }} />

          {/* Lists */}
          <div style={{ display: 'flex', flexDirection: 'row', gap: '2px', marginRight: '4px' }}>
            <button
              type="button"
              onMouseDown={(ev) => { ev.preventDefault(); this.applyCommand('insertUnorderedList'); }}
              title="Bullet List"
              style={this.getButtonStyle()}
            >
              •
            </button>
            <button
              type="button"
              onMouseDown={(ev) => { ev.preventDefault(); this.applyCommand('insertOrderedList'); }}
              title="Numbered List"
              style={this.getButtonStyle()}
            >
              1.
            </button>
          </div>

          {/* Divider */}
          <div style={{ width: '1px', backgroundColor: '#ccc', margin: '0 2px' }} />

          {/* Other */}
          <div style={{ display: 'flex', flexDirection: 'row', gap: '2px' }}>
            <button
              type="button"
              onClick={this.insertImage}
              title="Insert Image"
              style={this.getButtonStyle()}
            >
              🖼
            </button>
            <button
              type="button"
              onClick={this.clearFormatting}
              title="Clear Formatting"
              style={this.getButtonStyle()}
            >
              ✕
            </button>
          </div>
        </div>

        <div
          ref={(el) => { this.editorRef = el as HTMLDivElement; }}
          contentEditable={true}
          onInput={this.handleInput}
          onBlur={this.handleInput}
          data-placeholder={this.props.placeholder || ''}
          style={{
            padding: '12px',
            minHeight: '300px',
            maxHeight: '600px',
            overflowY: 'auto',
            fontFamily: "'Segoe UI', Arial, sans-serif",
            fontSize: '13px',
            lineHeight: '1.6',
            outline: 'none',
            border: 'none',
            width: '100%',
            boxSizing: 'border-box',
            direction: 'ltr',
            backgroundColor: '#ffffff',
            color: '#000000'
          }}
        ></div>
      </div>
    );
  }

  private getButtonStyle(): React.CSSProperties {
    return {
      padding: '4px 8px',
      minWidth: '24px',
      height: '24px',
      border: '1px solid #ccc',
      backgroundColor: '#ffffff',
      borderRadius: '2px',
      cursor: 'pointer',
      fontSize: '12px',
      fontWeight: 'normal',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      transition: 'all 0.2s ease'
    };
  }
}
