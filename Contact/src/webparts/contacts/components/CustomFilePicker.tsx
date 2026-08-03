import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';

export interface ICustomFilePickerProps {
  label: string;
  value: string;
  onChange: (newValue: string) => void;
}

export interface ICustomFilePickerState {
  fileUrl: string;
}

export class CustomFilePicker extends React.Component<ICustomFilePickerProps, ICustomFilePickerState> {
  constructor(props: ICustomFilePickerProps) {
    super(props);
    this.state = {
      fileUrl: props.value
    };
  }

  private handleBrowse = (): void => {
    const url = prompt("Enter file URL:");
    if (url) {
      this.setState({ fileUrl: url });
      this.props.onChange(url);
    }
  }

  public render(): React.ReactElement<ICustomFilePickerProps> {
    return (
      <div>
        <TextField label={this.props.label} value={this.state.fileUrl} readOnly />
        <DefaultButton text="Browse" onClick={this.handleBrowse} />
      </div>
    );
  }
}