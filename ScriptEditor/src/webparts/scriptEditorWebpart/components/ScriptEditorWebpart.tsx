import * as React from 'react';
import { IScriptEditorProps } from './IScriptEditorProps';

// Define a state interface compatible with TS 2.4.x
export interface IScriptEditorState {
  script?: string;
  loaded?: boolean;
}

export default class ScriptEditor extends React.Component<IScriptEditorProps, IScriptEditorState> {
  constructor(props: IScriptEditorProps) {
    super(props);

    this._showDialog = this._showDialog.bind(this);

    // Initialize minimal state in the constructor for React 15
    this.state = {
      script: '',
      loaded: false
    };
  }

  public componentDidMount(): void {
    // Populate from props once mounted
    this.setState({
      script: this.props.script,
      loaded: !!this.props.script
    });
  }

  private _showDialog(): void {
    // Open the web part property pane (provided by parent/web part)
    if (this.props && this.props.propPaneHandle && typeof this.props.propPaneHandle.open === 'function') {
      this.props.propPaneHandle.open();
    }
  }

  public render(): React.ReactElement<IScriptEditorProps> {
    const hasScript = !!this.state.script;

    // Render configured markup (HTML/JS) - NOTE: be mindful of CSP/custom script settings on your tenant/site
    const viewMode = hasScript ? <span dangerouslySetInnerHTML={{ __html: this.state.script || '' }} /> : null;

    // Simple “configure” area that replaces the PnP Placeholder control
    const configureBlock = (
      <div className='ms-Fabric' style={{ marginBottom: 12 }}>
        <div style={{ fontSize: 16, fontWeight: 600 }}>
          {this.props.title || 'Script Editor'}
        </div>
        {!hasScript && (
          <div style={{ marginTop: 6, color: '#666' }}>
            Please configure the web part.
          </div>
        )}
        <button
          type='button'
          className='ms-Button ms-Button--primary'
          style={{ marginTop: 8 }}
          onClick={this._showDialog}
          aria-label='Edit markup'
          title='Edit markup'
        >
          <span className='ms-Button-label'>Edit markup</span>
        </button>
      </div>
    );

    return (
      <div className='ms-Fabric'>
        {configureBlock}
        {viewMode}
      </div>
    );
  }
}
