import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, PropertyPaneCheckbox } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'ReactListViewWebPartStrings';
import ReactListView from './components/ReactListView';
import { IReactListViewProps } from './components/IReactListViewProps';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import axios from 'axios';
import { ColumnParser } from '../../utilities/columnParser';
import { SPCalls } from '../../utilities/spCalls';
export interface IReactListViewWebPartProps {
  description: string;
  siteUrl: string;
  listName: string;
  internalListName: string;
  multiSelect: string[];
  filterQuery: string;
  ColumnNumberForContext: string;
  showEditButton: boolean;
  showNewButton: boolean;
  showExportButton: boolean;
  webpartHeight: number;
  columnFormatter: string;
}

export default class ReactListViewWebPart extends BaseClientSideWebPart<IReactListViewWebPartProps> {
  globalOptions = [];
  listColumns = [];
  public render(): void {
    var pthis = this;
    this.listColumns = [];
    SPCalls.getListColumns(this.properties.siteUrl, this.properties.listName).then(data => {
      pthis.setColumnsForPropertyPane(data, pthis);
      if (pthis.properties.multiSelect) {
        pthis.filterSelectedColumns(pthis, data);
        SPCalls.getListItems(pthis.properties.siteUrl, pthis.properties.listName, pthis.listColumns, pthis.properties.filterQuery).then(listItems => {
          var lcol = pthis.listColumns;
          const element: React.ReactElement<IReactListViewProps> = React.createElement(
            ReactListView,
            {
              description: pthis.properties.description,
              columns: lcol,
              rows: pthis.getListRows(listItems, pthis),
              newURL: pthis.properties.siteUrl + "/Lists/" + pthis.properties.internalListName + "/NewForm.aspx",
              dispURL: pthis.properties.siteUrl + "/Lists/" + pthis.properties.internalListName + "/DispForm.aspx?ID=",
              editURL: pthis.properties.siteUrl + "/Lists/" + pthis.properties.internalListName + "/EditForm.aspx?ID=",
              viewColumnName: pthis.properties.ColumnNumberForContext,
              showNewButton: pthis.properties.showNewButton,
              showEditButton: pthis.properties.showEditButton,
              showExportButton: pthis.properties.showExportButton,
              wrapperHeight: pthis.properties.webpartHeight
            }
          );
          ReactDom.render(element, pthis.domElement);
        })
      }
    });
  }

  private setColumnsForPropertyPane(data, pthis) {
    pthis.globalOptions = [];
    data.forEach(item => {
      if (item.Group != "_Hidden" && item.TypeAsString != "Attachments") {
        pthis.globalOptions.push({
          key: item.InternalName,
          text: item.Title
        });
      }
    });
    pthis.context.propertyPane.refresh();
  }

  private filterSelectedColumns(pthis, data) {
    pthis.properties.multiSelect.forEach(selectedColumn => {
      data.forEach(listColumn => {
        if (listColumn.InternalName == selectedColumn) {
          pthis.listColumns.push(listColumn);
        }
      })
    });
  }

  private getListRows = (listItems, pthis) => {
    var rows = [];
    listItems.forEach(item => {
      var row = {};
      pthis.listColumns.forEach(column => {
        row[column.InternalName] = ColumnParser.getColumnValue(item, column);
        if (pthis.properties.columnFormatter) {
          var formatter = JSON.parse(pthis.properties.columnFormatter.replace(/\s/g, ""));
          if (formatter[column.InternalName]) {
            var expression = formatter[column.InternalName].replace(/COLUMN_VALUE/g, row[column.InternalName]);
            row[column.InternalName] = eval(expression);
          }
        }
      });
      rows.push(row);
    });
    return rows;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('siteUrl', {
                  label: "Site URL:"
                }),
                PropertyPaneTextField('listName', {
                  label: "List Name"
                }),
                PropertyPaneTextField('internalListName', {
                  label: "Internal List Name"
                }),
                PropertyFieldMultiSelect('multiSelect', {
                  key: 'multiSelect',
                  label: "List Columns to show in grid",
                  options: this.globalOptions,
                  selectedKeys: this.properties.multiSelect
                }),
                PropertyPaneTextField('filterQuery', {
                  label: "Filter Query"
                }),
                PropertyPaneTextField('ColumnNumberForContext', {
                  label: "Internal Column Name for View Link"
                }),
                PropertyPaneCheckbox('showNewButton', {
                  text: "Show New Button"
                }),
                PropertyPaneCheckbox('showEditButton', {
                  text: "Show Edit Button"
                }),
                PropertyPaneCheckbox('showExportButton', {
                  text: "Show Export Button"
                }),
                PropertyPaneTextField('webpartHeight', {
                  label: "Webpart Height (Only numeric values allowed. Unit is pixels)"
                }),
                PropertyPaneTextField('columnFormatter', {
                  label: "Column formatting (Beta feature)",
                  multiline: true,
                  rows: 6
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
