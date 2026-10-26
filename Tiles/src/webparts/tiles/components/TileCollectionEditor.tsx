import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Guid } from '@microsoft/sp-core-library';
import { assign } from '@microsoft/sp-lodash-subset';
import { IPropertyPaneCustomFieldProps, IPropertyPaneField, PropertyPaneFieldType } from '@microsoft/sp-webpart-base';
import { CustomCollectionFieldType, ICustomCollectionField, IPropertyFieldCollectionDataProps } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import * as strings from 'TilesWebPartStrings';

export interface ITileEditorItem {
  [fieldId: string]: any;
  uniqueId: string;
}

export interface ITileCollectionEditorProps extends IPropertyFieldCollectionDataProps {
  onChanged: (items: ITileEditorItem[]) => void;
}

export interface ITileCollectionEditorState {
  open: boolean;
  items: ITileEditorItem[];
  draft: ITileEditorItem;
  customErrors: { [rowId: string]: { [fieldId: string]: string } };
  showValidation: boolean;
  saveError: string;
}

export class TileCollectionEditor extends React.Component<ITileCollectionEditorProps, ITileCollectionEditorState> {
  public constructor(props: ITileCollectionEditorProps) {
    super(props);
    this.state = {
      open: false, items: [], draft: this.createDraft(), customErrors: {}, showValidation: false, saveError: ''
    };
  }

  private createDraft(): ITileEditorItem {
    const draft: ITileEditorItem = { uniqueId: Guid.newGuid().toString() };
    this.props.fields.forEach((field) => {
      draft[field.id] = field.defaultValue !== undefined ? field.defaultValue : '';
    });
    return draft;
  }

  private openPanel = (): void => {
    this.setState({
      open: true,
      items: (this.props.value || []).map((item) => assign({}, item, {
        uniqueId: item.uniqueId || Guid.newGuid().toString()
      })),
      draft: this.createDraft(),
      customErrors: {},
      showValidation: false,
      saveError: ''
    });
  }

  private closePanel = (): void => {
    this.setState({ open: false });
  }

  private hasDraftChanges(): boolean {
    return this.props.fields.some((field) => this.state.draft[field.id] !==
      (field.defaultValue !== undefined ? field.defaultValue : ''));
  }

  private getFieldError(item: ITileEditorItem, field: ICustomCollectionField): string {
    if (field.required && !String(item[field.id] || '').trim()) {
      return strings.tileFieldRequired.replace('{0}', field.title);
    }
    const rowErrors = this.state.customErrors[item.uniqueId];
    return rowErrors && rowErrors[field.id] || '';
  }

  private itemIsValid(item: ITileEditorItem): boolean {
    return this.props.fields.every((field) => !this.getFieldError(item, field));
  }

  private canSave(): boolean {
    return this.state.items.every((item) => this.itemIsValid(item)) &&
      (!this.hasDraftChanges() || this.itemIsValid(this.state.draft));
  }

  private updateField(rowId: string, fieldId: string, value: any): void {
    this.setState((previous) => {
      const changes: any = {};
      changes[fieldId] = value;
      return {
        items: previous.items.map((item) => item.uniqueId === rowId ? assign({}, item, changes) : item),
        draft: previous.draft.uniqueId === rowId ? assign({}, previous.draft, changes) : previous.draft,
        showValidation: false,
        saveError: ''
      };
    });
  }

  private validateCustomField(rowId: string, fieldId: string, error: string): void {
    this.setState((previous) => {
      const rowErrors = assign({}, previous.customErrors[rowId]);
      rowErrors[fieldId] = error;
      const customErrors = assign({}, previous.customErrors);
      customErrors[rowId] = rowErrors;
      return { customErrors: customErrors };
    });
  }

  private addDraft = (): void => {
    if (!this.hasDraftChanges() || !this.itemIsValid(this.state.draft)) {
      this.setState({ showValidation: true });
      return;
    }
    this.setState((previous) => ({
      items: previous.items.concat([previous.draft]),
      draft: this.createDraft(),
      showValidation: false,
      saveError: ''
    }));
  }

  private deleteItem(rowId: string): void {
    this.setState((previous) => ({
      items: previous.items.filter((item) => item.uniqueId !== rowId),
      showValidation: false,
      saveError: ''
    }));
  }

  private moveItem(rowId: string, position: number): void {
    this.setState((previous) => {
      const items = previous.items.slice();
      const current = items.map((item) => item.uniqueId).indexOf(rowId);
      const moved = items.splice(current, 1)[0];
      items.splice(position, 0, moved);
      return { items: items };
    });
  }

  private save = (): void => {
    if (!this.canSave()) {
      this.setState({ showValidation: true });
      return;
    }
    const items = this.hasDraftChanges() ? this.state.items.concat([this.state.draft]) : this.state.items;
    const saved = items.map((item, index) => assign({}, item, { sortIdx: index + 1 }));
    try {
      this.props.onChanged(saved);
      this.closePanel();
    } catch (error) {
      console.error('[Tiles] Unable to save tile configuration.', error);
      this.setState({ saveError: strings.tileSaveError });
    }
  }

  private renderField(item: ITileEditorItem, field: ICustomCollectionField, draft: boolean): React.ReactElement<any> {
    const value = item[field.id];
    const error = !draft || this.hasDraftChanges() || this.state.showValidation ? this.getFieldError(item, field) : '';
    const disabled = !draft && field.disableEdit;
    const onUpdate = (fieldId: string, nextValue: any): void => this.updateField(item.uniqueId, fieldId, nextValue);
    let control: React.ReactElement<any>;
    if (field.type === CustomCollectionFieldType.custom && field.onCustomRender) {
      control = field.onCustomRender(field, value, onUpdate, item, item.uniqueId,
        (fieldId: string, message: string): void => this.validateCustomField(item.uniqueId, fieldId, message));
    } else if (field.type === CustomCollectionFieldType.boolean) {
      control = <Checkbox checked={!!value} disabled={disabled} ariaLabel={field.title}
        onChange={(event, checked): void => onUpdate(field.id, checked)} />;
    } else if (field.type === CustomCollectionFieldType.dropdown) {
      control = <Dropdown ariaLabel={field.title} selectedKey={value} disabled={disabled}
        options={typeof field.options === 'function' ? field.options(field.id, item) : field.options || []}
        onChanged={(option): void => onUpdate(field.id, option.key)} />;
    } else {
      return <TextField ariaLabel={field.title} value={String(value || '')}
        placeholder={field.placeholder || field.title} required={field.required} disabled={disabled}
        errorMessage={error} onChanged={(nextValue): void => onUpdate(field.id, nextValue)} />;
    }
    return <div>{control}{error && <div role="alert" style={{ color: '#a80000' }}>{error}</div>}</div>;
  }

  private renderRow(item: ITileEditorItem, index: number, draft: boolean): React.ReactElement<any> {
    return <tr key={item.uniqueId}>
      {this.props.enableSorting && <td style={{ minWidth: 70, verticalAlign: 'top', padding: 4 }}>
        {!draft && <Dropdown ariaLabel={strings.tileOrderLabel} selectedKey={index} options={this.state.items.map((row, position) => ({
          key: position, text: String(position + 1)
        }))} onChanged={(option): void => this.moveItem(item.uniqueId, Number(option.key))} />}
      </td>}
      {this.props.fields.map((field) => <td key={field.id} style={{ minWidth: 140, verticalAlign: 'top', padding: 4 }}>
        {this.renderField(item, field, draft)}
      </td>)}
      <td style={{ verticalAlign: 'top', padding: 4 }}>
        {draft
          ? <DefaultButton text={strings.tileAddButton} onClick={this.addDraft}
            disabled={!this.hasDraftChanges() || !this.itemIsValid(item) || this.props.disableItemCreation} />
          : <DefaultButton text={strings.tileDeleteButton} onClick={(): void => this.deleteItem(item.uniqueId)}
            disabled={this.props.disableItemDeletion} />}
      </td>
    </tr>;
  }

  public render(): React.ReactElement<ITileCollectionEditorProps> {
    const draftChanged = this.hasDraftChanges();
    const invalid = !this.canSave();
    return <div>
      <div>{this.props.label}</div>
      <DefaultButton text={this.props.manageBtnLabel} onClick={this.openPanel} disabled={this.props.disabled} />
      <Panel isOpen={this.state.open} type={PanelType.large} headerText={this.props.panelHeader}
        onDismiss={this.closePanel} onOuterClick={(): void => undefined} isLightDismiss={false}>
        {this.props.panelDescription && <p>{this.props.panelDescription}</p>}
        {(invalid || this.state.showValidation) && <p role="alert" style={{ color: '#a80000' }}>{strings.tileRequiredFieldsHelp}</p>}
        {this.state.saveError && <p role="alert" style={{ color: '#a80000' }}>{this.state.saveError}</p>}
        <div style={{ overflowX: 'auto' }}>
          <table style={{ borderCollapse: 'collapse', width: '100%' }}>
            <thead><tr>
              {this.props.enableSorting && <th scope="col">{strings.tileOrderLabel}</th>}
              {this.props.fields.map((field) => <th key={field.id} scope="col" style={{ textAlign: 'left', padding: 4 }}>
                {field.title}{field.required && <span aria-label={strings.tileRequiredLabel}> *</span>}
              </th>)}
              <th scope="col">{strings.tileActionsLabel}</th>
            </tr></thead>
            <tbody>
              {this.state.items.map((item, index) => this.renderRow(item, index, false))}
              {!this.props.disableItemCreation && this.renderRow(this.state.draft, this.state.items.length, true)}
            </tbody>
          </table>
        </div>
        <div style={{ marginTop: 16 }}>
          <PrimaryButton text={draftChanged ? strings.tileSaveAndAddButton : strings.tileSaveButton}
            onClick={this.save} disabled={invalid} />
          <DefaultButton text={strings.tileCancelButton} onClick={this.closePanel} style={{ marginLeft: 8 }} />
        </div>
      </Panel>
    </div>;
  }
}

export function PropertyPaneTileCollection(
  targetProperty: string, properties: IPropertyFieldCollectionDataProps
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty: targetProperty,
    properties: {
      key: properties.key,
      onRender: (element: HTMLElement, context: any, changeCallback: (property: string, value: any) => void): void => {
        ReactDom.render(React.createElement(TileCollectionEditor, assign({}, properties, {
          onChanged: (items: ITileEditorItem[]): void => {
            if (!changeCallback) {
              throw new Error('The tile collection property change callback is unavailable.');
            }
            changeCallback(targetProperty, items);
          }
        })), element);
      },
      onDispose: (element: HTMLElement): void => { ReactDom.unmountComponentAtNode(element); }
    }
  };
}
