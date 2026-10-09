import * as React from 'react';
import styles from './SharePointDynamicForm.module.scss';

export interface IReadOnlyRichTextProps {
  html: string;
  ariaLabel: string;
  style?: React.CSSProperties;
}

export interface IReadOnlyMultilineTextProps {
  text: string;
  ariaLabel: string;
  style?: React.CSSProperties;
}

export class ReadOnlyRichText extends React.Component<IReadOnlyRichTextProps> {
  public render(): JSX.Element {
    return (
      <div
        className={styles.richTextReadOnly}
        dir="ltr"
        role="region"
        aria-label={this.props.ariaLabel}
        aria-readonly={true}
        tabIndex={0}
        style={this.props.style}
        dangerouslySetInnerHTML={{ __html: this.props.html } as any}
      />
    );
  }
}

export class ReadOnlyMultilineText extends React.Component<IReadOnlyMultilineTextProps> {
  public render(): JSX.Element {
    return (
      <div
        className={styles.richTextReadOnly}
        dir="ltr"
        role="region"
        aria-label={this.props.ariaLabel}
        aria-readonly={true}
        tabIndex={0}
        style={Object.assign({
          whiteSpace: 'pre-wrap',
          wordBreak: 'break-word',
        }, this.props.style)}
      >
        {this.props.text}
      </div>
    );
  }
}
