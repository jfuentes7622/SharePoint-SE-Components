import * as React from 'react';
import styles from './SharePointDynamicForm.module.scss';

export interface RichTextEditorProps {
  value: string;
  onChange: (html: string) => void;
  placeholder?: string;
}

export class RichTextEditor extends React.Component<RichTextEditorProps> {
  private editorRef: HTMLTextAreaElement;

  private handleChange = (ev: React.ChangeEvent<HTMLTextAreaElement>) => {
    this.props.onChange(ev.currentTarget.value);
  }

  private wrapSelection = (beforeTag: string, afterTag: string = '') => {
    if (!this.editorRef) return;

    const textarea = this.editorRef;
    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const text = textarea.value;
    const selectedText = text.substring(start, end) || 'text';
    const before = text.substring(0, start);
    const after = text.substring(end);

    const newText = before + beforeTag + selectedText + (afterTag || beforeTag.replace('<', '</')) + after;
    this.props.onChange(newText);

    // Reset focus and cursor position
    setTimeout(() => {
      if (this.editorRef) {
        this.editorRef.focus();
        this.editorRef.selectionStart = start + beforeTag.length;
        this.editorRef.selectionEnd = start + beforeTag.length + selectedText.length;
      }
    }, 0);
  }

  private insertTag = (tag: string) => {
    this.wrapSelection(`<${tag}>`, `</${tag}>`);
  }

  private insertDiv = (alignment: 'left' | 'center' | 'right') => {
    const alignStyle = alignment === 'left' ? '' : ` style="text-align:${alignment};"`;
    this.wrapSelection(`<div${alignStyle}>`, '</div>');
  }

  private insertList = (ordered: boolean) => {
    if (!this.editorRef) return;
    const tag = ordered ? 'ol' : 'ul';
    this.wrapSelection(`<${tag}><li>`, `</li></${tag}>`);
  }

  private insertImage = () => {
    const url = prompt('Enter image URL:');
    if (url) {
      this.wrapSelection(`<img src="${url}" alt="image" />`);
    }
  }

  private clearFormatting = () => {
    if (!this.editorRef) return;
    // Simple removal of common HTML tags
    const text = this.editorRef.value;
    const cleaned = text.replace(/<[^>]*>/g, '');
    this.props.onChange(cleaned);
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
              onClick={() => this.insertTag('b')}
              title="Bold"
              style={this.getButtonStyle()}
            >
              <strong>B</strong>
            </button>
            <button
              onClick={() => this.insertTag('i')}
              title="Italic"
              style={this.getButtonStyle()}
            >
              <em>I</em>
            </button>
            <button
              onClick={() => this.insertTag('u')}
              title="Underline"
              style={this.getButtonStyle()}
            >
              <u>U</u>
            </button>
            <button
              onClick={() => this.insertTag('s')}
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
              onClick={() => this.insertDiv('left')}
              title="Align Left"
              style={this.getButtonStyle()}
            >
              ≡
            </button>
            <button
              onClick={() => this.insertDiv('center')}
              title="Align Center"
              style={this.getButtonStyle()}
            >
              ⋮
            </button>
            <button
              onClick={() => this.insertDiv('right')}
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
              onClick={() => this.insertList(false)}
              title="Bullet List"
              style={this.getButtonStyle()}
            >
              •
            </button>
            <button
              onClick={() => this.insertList(true)}
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
              onClick={this.insertImage}
              title="Insert Image"
              style={this.getButtonStyle()}
            >
              🖼
            </button>
            <button
              onClick={this.clearFormatting}
              title="Clear Formatting"
              style={this.getButtonStyle()}
            >
              ✕
            </button>
          </div>
        </div>

        {/* Editor Area - Textarea to avoid RTL issues */}
        <textarea
          ref={(el) => { this.editorRef = el as HTMLTextAreaElement; }}
          onChange={this.handleChange}
          style={{
            padding: '12px',
            minHeight: '300px',
            maxHeight: '600px',
            overflowY: 'auto',
            fontFamily: 'Courier New, monospace',
            fontSize: '13px',
            lineHeight: '1.6',
            outline: 'none',
            resize: 'vertical',
            border: 'none',
            width: '100%',
            boxSizing: 'border-box',
            direction: 'ltr',
            backgroundColor: '#ffffff',
            color: '#000000'
          }}
          value={this.props.value}
          placeholder="Enter your HTML content here. Use toolbar buttons to format selected text."
        />
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
