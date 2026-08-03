import * as React from 'react';
import * as ReactDom from 'react-dom';
//import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
} from '@microsoft/sp-webpart-base';

import {
  PropertyFieldCustomList,
  //IPropertyFieldCustomListProps,
  CustomListFieldType
} from 'sp-client-custom-fields/lib/PropertyFieldCustomList';

import { IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';

//import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/propertyFields/number';
//import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'TilesWebPartStrings';
import { ITilesProps } from './components/ITilesProps';
import { ITileInfo } from './ITileInfo';
import { Tiles } from './components/Tiles';

export interface ITilesWebPartProps {
  collectionData: ITileInfo[];
  tileHeight: number;
  tileEffect: string;
  title: string;
}

export default class TilesWebPart extends BaseClientSideWebPart<ITilesWebPartProps> {

//  private propertyFieldNumber: any;
  // private propertyFieldCollectionData: any;
  // private customCollectionFieldType: any;


  public render(): void {
    const element: React.ReactElement<ITilesProps> = React.createElement(
      Tiles,
      {
        title: this.properties.title,
        tileHeight: this.properties.tileHeight,
        tileEffect: this.properties.tileEffect,
        collectionData: this.properties.collectionData,
        displayMode: this.displayMode,
        fUpdateProperty: (value: string) => {
          this.properties.title = value;
        },
        fPropertyPaneOpen: this.context.propertyPane.open
      }
    );

    ReactDom.render(element, this.domElement);
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  // executes only before property pane is loaded.
  // protected async loadPropertyPaneResources(): Promise<void> {

  //   // import additional controls/components
  //  /* const { PropertyFieldNumber } = await import(
  //     '@pnp/spfx-property-controls/lib/propertyFields/number'
  //   );*/
  //   const { PropertyFieldCollectionData, CustomCollectionFieldType } = await import(
  //     '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData'
  //   );

  //  // this.propertyFieldNumber = PropertyFieldNumber;
  //   this.propertyFieldCollectionData = PropertyFieldCollectionData;
  //   this.customCollectionFieldType = CustomCollectionFieldType;
  // }

  // protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  //   //this.propertyFieldNumber = PropertyFieldNumber;
  //   this.propertyFieldCollectionData = PropertyFieldCollectionData;
  //   this.customCollectionFieldType = CustomCollectionFieldType;
  //   return {
  //     pages: [
  //       {
  //         header: {
  //           description: strings.PropertyPaneDescription + ` v${this.context.manifest.version}`,
  //         },
  //         groups: [
  //           {
  //             groupFields: [
  //               this.propertyFieldCollectionData('collectionData', {
  //                 key: 'collectionData',
  //                 label: strings.tilesDataLabel,
  //                 panelHeader: strings.tilesPanelHeader,
  //                //tslint:disable-next-line:max-line-length,
  //                 // panelDescription: `${strings.iconInformation} https://developer.microsoft.com/en-us/fabric#/styles/icons`,
  //                 manageBtnLabel: strings.tilesManageBtn,
  //                 value: this.properties.collectionData,
  //                 enableSorting: true,
  //                 fields: [
  //                   {
  //                     id: 'title',
  //                     title: strings.titleField,
  //                     type: this.customCollectionFieldType.string,
  //                     required: true
  //                   },
  //                   {
  //                     id: 'description',
  //                     title: strings.descriptionField,
  //                     type: this.customCollectionFieldType.string,
  //                     required: false
  //                   },
  //                   {
  //                     id: 'url',
  //                     title: strings.urlField,
  //                     type: this.customCollectionFieldType.string,
  //                     required: true
  //                   },
  //                 /*   {
  //                     id: 'icon',
  //                     title: strings.iconField,
  //                     type: this.customCollectionFieldType.fabricIcon,
  //                     required: true
  //                   }, */
  //                   {
  //                     id: 'target',
  //                     title: strings.targetField,
  //                     type: this.customCollectionFieldType.dropdown,
  //                     options: [
  //                       {
  //                         key: LinkTarget.parent,
  //                         text: strings.targetCurrent
  //                       },
  //                       {
  //                         key: LinkTarget.blank,
  //                         text: strings.targetNew
  //                       }
  //                     ]
  //                   }
  //                 ]
  //               }),
  //            /*    this.propertyFieldNumber('tileHeight', {
  //                 key: 'tileHeight',
  //                 label: strings.TileHeight,
  //                 value: this.properties.tileHeight
  //               }) */
  //             /*   PropertyPaneDropdown('tileEffect', {
  //                 label: 'Tiles Animation Effect',
  //                 options: [
  //                   { key: '1', text: 'Blend' },
  //                   { key: '2', text: 'Grow' },
  //                   { key: '3', text: 'Scroll Vertical' },
  //                   { key: '4', text: 'Scroll Horizontal' }
  //                 ],
  //                 selectedKey: '3'
  //               }) */
  //             ]
  //           }
  //         ]
  //       }
  //     ]
  //   };
  // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages: [
      {
        header: {
          description: `${strings.PropertyPaneDescription} v${this.context.manifest.version}`
        },
        groups: [
          {
            groupFields: [
              PropertyFieldCustomList('collectionData', {
                key: 'collectionData',
                label: strings.tilesDataLabel,
                headerText: strings.tilesPanelHeader,  // same as in IPropertyFieldCustomListProps
                value: this.properties.collectionData || [],
                context: this.context,
                onPropertyChange: this.onPropertyPaneFieldChanged,
                render: this.render.bind(this),
                properties: this.properties,
                fields: [
                  {
                    id: 'title',
                    title: strings.titleField,
                    type: CustomListFieldType.string,
                    required: true
                  },
                  {
                    id: 'description',
                    title: strings.descriptionField,
                    type: CustomListFieldType.string
                  },
                  {
                    id: 'url',
                    title: strings.urlField,
                    type: CustomListFieldType.string,
                    required: true
                  },
                  {
                    id: 'target',
                    title: strings.targetField,
                    type: CustomListFieldType.string,
                    required: true
                  }
                ],
                disableReactivePropertyChanges: false
              })
            ]
          }
        ]
      }
    ]
  };
}
}
