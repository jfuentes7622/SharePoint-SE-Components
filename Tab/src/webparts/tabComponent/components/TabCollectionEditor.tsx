import * as React from 'react';
import { ITabCollectionItem } from '../TabComponentWebPart';

export interface IImageOption {
  key: string;
  text: string;
}

export interface ITabCollectionEditorProps {
  value: ITabCollectionItem[];
  imageOptions: IImageOption[];
  onChange: (value: ITabCollectionItem[]) => void;
}

export default class TabCollectionEditor extends React.Component<ITabCollectionEditorProps, {}> {
  private getRows(): ITabCollectionItem[] {
    return (this.props.value || []).map((item: ITabCollectionItem) => {
      return {
        Title: item.Title || '',
        textPosition: item.textPosition || 'left',
        imageUrl: item.imageUrl || '',
        imagePosition: item.imagePosition || 'left',
        onlyImage: item.onlyImage === true
      };
    });
  }

  private updateRow(index: number, propertyName: string, value: any): void {
    const rows = this.getRows();
    (rows[index] as any)[propertyName] = value;
    this.props.onChange(rows);
  }

  private addRow(): void {
    const rows = this.getRows();
    rows.push({ Title: '', textPosition: 'left', imageUrl: '', imagePosition: 'left', onlyImage: false });
    this.props.onChange(rows);
  }

  private removeRow(index: number): void {
    const rows = this.getRows();
    rows.splice(index, 1);
    this.props.onChange(rows);
  }

  private moveRow(index: number, direction: number): void {
    const rows = this.getRows();
    const targetIndex = index + direction;
    if (targetIndex < 0 || targetIndex >= rows.length) {
      return;
    }
    const row = rows[index];
    rows[index] = rows[targetIndex];
    rows[targetIndex] = row;
    this.props.onChange(rows);
  }

  private containsImageOption(imageUrl: string): boolean {
    return this.props.imageOptions.some((option: IImageOption) => option.key === imageUrl);
  }

  public render(): React.ReactElement<ITabCollectionEditorProps> {
    const rows = this.getRows();
    const fieldStyle: React.CSSProperties = { width: '100%', minHeight: '30px', marginBottom: '8px' };
    const labelStyle: React.CSSProperties = { display: 'block', fontSize: '12px', fontWeight: 600, marginBottom: '3px' };
    const rowStyle: React.CSSProperties = { borderBottom: '1px solid #d2d0ce', padding: '12px 0' };
    return (
      <div>
        {rows.map((row: ITabCollectionItem, index: number) => (
          <div key={'tab-editor-' + index} style={rowStyle}>
            <label style={labelStyle}>Title</label>
            <input style={fieldStyle} value={row.Title || ''}
              onChange={(event) => { this.updateRow(index, 'Title', event.currentTarget.value); }} />
            <label style={labelStyle}>Text alignment</label>
            <select style={fieldStyle} value={row.textPosition || 'left'}
              onChange={(event) => { this.updateRow(index, 'textPosition', event.currentTarget.value); }}>
              <option value='left' aria-selected={(row.textPosition || 'left') === 'left'}>Left</option>
              <option value='center' aria-selected={row.textPosition === 'center'}>Center</option>
              <option value='right' aria-selected={row.textPosition === 'right'}>Right</option>
            </select>
            <label style={labelStyle}>Image</label>
            <select style={fieldStyle} value={row.imageUrl || ''}
              onChange={(event) => { this.updateRow(index, 'imageUrl', event.currentTarget.value); }}>
              <option value='' aria-selected={!row.imageUrl}>No image</option>
              {row.imageUrl && !this.containsImageOption(row.imageUrl) ? (
                <option value={row.imageUrl} aria-selected={true}>Current image / {row.imageUrl}</option>
              ) : undefined}
              {this.props.imageOptions.map((option: IImageOption) => (
                <option key={option.key} value={option.key} aria-selected={row.imageUrl === option.key}>{option.text}</option>
              ))}
            </select>
            <label style={labelStyle}>Image position</label>
            <select style={fieldStyle} value={row.imagePosition || 'left'}
              onChange={(event) => { this.updateRow(index, 'imagePosition', event.currentTarget.value); }}>
              <option value='left' aria-selected={(row.imagePosition || 'left') === 'left'}>Left</option>
              <option value='right' aria-selected={row.imagePosition === 'right'}>Right</option>
            </select>
            <label style={{ display: 'block', marginBottom: '10px' }}>
              <input type='checkbox' checked={row.onlyImage === true} aria-checked={row.onlyImage === true}
                onChange={(event) => { this.updateRow(index, 'onlyImage', event.currentTarget.checked); }} />
              <span style={{ marginLeft: '6px' }}>Only image</span>
            </label>
            <button type='button' disabled={index === 0} onClick={() => { this.moveRow(index, -1); }}>Move up</button>
            <button type='button' disabled={index === rows.length - 1} style={{ marginLeft: '6px' }}
              onClick={() => { this.moveRow(index, 1); }}>Move down</button>
            <button type='button' style={{ marginLeft: '6px' }} onClick={() => { this.removeRow(index); }}>Remove</button>
          </div>
        ))}
        <button type='button' style={{ marginTop: '12px' }} onClick={() => { this.addRow(); }}>Add tab</button>
      </div>
    );
  }
}