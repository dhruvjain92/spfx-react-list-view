import * as React from 'react';
import styles from './ReactListView.module.scss';
import { IReactListViewProps } from './IReactListViewProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn, IDetailsHeaderProps } from 'office-ui-fabric-react/lib/DetailsList';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Stack, DefaultButton, IRenderFunction, TooltipHost, ITooltipHostProps } from 'office-ui-fabric-react';
import 'office-ui-fabric-core/dist/css/fabric.css';
import { Link } from 'office-ui-fabric-react/lib/Link';
import * as XLSX from 'xlsx';
import { ScrollablePane, ScrollbarVisibility } from 'office-ui-fabric-react/lib/ScrollablePane';
import { Sticky, StickyPositionType } from 'office-ui-fabric-react/lib/Sticky';

export default class ReactListView extends React.Component<IReactListViewProps, { stateColumns, stateRows, selectedItemCount }> {
  private _selection: Selection;
  globalColumns = [];
  globalRows = [];
  newURL = "";
  selectedItemID = 0;
  constructor(props: IReactListViewProps) {
    super(props);
    const { columns, rows, newURL, viewColumnName } = props;
    this.newURL = newURL;
    const gridColumns: IColumn[] = [];
    columns.forEach(column => {
      var c;
      if (column.InternalName == viewColumnName) {
        c = {
          key: column.InternalName,
          name: column.Title,
          fieldName: column.InternalName,
          minWidth: 100,
          maxWidth: 200,
          onColumnClick: this._onColumnClick,
          onRender: this.checkRender
        }
      } else {
        c = {
          key: column.InternalName,
          name: column.Title,
          fieldName: column.InternalName,
          minWidth: 100,
          maxWidth: 200,
          onColumnClick: this._onColumnClick
        }
      }
      gridColumns.push(c)
    })
    this.state = {
      stateColumns: gridColumns,
      stateRows: rows,
      selectedItemCount: 0
    }
    this.globalColumns = gridColumns;
    this.globalRows = rows;
    this._selection = new Selection({
      onSelectionChanged: () => {
        var selectedItemCount = this._selection.getSelectedCount();
        if (selectedItemCount == 1) {
          var item = this._selection.getSelection()[0];
          if (item["ID"]) {
            this.selectedItemID = item["ID"];
          } else {
            selectedItemCount = 0;
          }
        }
        this.setState({
          selectedItemCount: selectedItemCount
        })
      }
    });

  }
  public render(): React.ReactElement<IReactListViewProps> {
    const { stateColumns, stateRows, selectedItemCount } = this.state;
    const { showEditButton, showNewButton, showExportButton, wrapperHeight } = this.props;
    var wrapperStyle = { "height": wrapperHeight + "px" } as React.CSSProperties
    return (
      <div className={styles.wrapper} style={wrapperStyle}>
        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
          <div className={styles.reactListView}>
            <div className="ms-Grid" dir="ltr">
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg3">
                  <TextField
                    className={styles.searchBox}
                    onChange={this.searchBox}
                    placeholder="Search List"
                  />
                </div>
              </div>
              <Sticky stickyPosition={StickyPositionType.Header}>
                <div className="ms-Grid-row">
                  {showNewButton == true &&
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg2">
                      <DefaultButton
                        data-automation-id="test"
                        allowDisabledFocus={true}
                        text="New"
                        iconProps={{ iconName: 'Add' }}
                        onClick={this.goToNewItem}
                      />
                    </div>
                  }
                  {showEditButton == true &&
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg2">
                      <DefaultButton
                        data-automation-id="test"
                        allowDisabledFocus={true}
                        text="Edit"
                        iconProps={{ iconName: 'Edit' }}
                        onClick={this.goToEditItem}
                        disabled={!Boolean(selectedItemCount)}
                      />
                    </div>
                  }
                  {showExportButton == true &&
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg2">
                      <DefaultButton
                        data-automation-id="test"
                        allowDisabledFocus={true}
                        text="Export"
                        iconProps={{ iconName: 'ExcelLogo' }}
                        onClick={this.exportExcel}
                      />
                    </div>
                  }
                </div>
              </Sticky>
            </div>
            <DetailsList
              items={stateRows}
              compact={false}
              className={styles.dlist}
              columns={stateColumns}
              selection={this._selection}
              selectionMode={SelectionMode.single}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              isHeaderVisible={true}
              onRenderDetailsHeader={this.onRenderDetailsHeader}
              selectionPreservedOnEmptyClick={true}
              enterModalSelectionOnTouch={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            />
            {!this.state.stateRows.length && (
              <Stack horizontalAlign='center'>
                No items found.
          </Stack>
            )}
          </div>
        </ScrollablePane>
      </div>
    );
  }


  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { stateColumns, stateRows } = this.state;
    const newColumns: IColumn[] = stateColumns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = this.copyAndSort(stateRows, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      stateColumns: newColumns,
      stateRows: newItems
    });
  };

  private onRenderDetailsHeader(props: IDetailsHeaderProps, defaultRender?: IRenderFunction<IDetailsHeaderProps>): JSX.Element {
    return (
      <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced={true}>
        {defaultRender!({
          ...props,
          onRenderColumnHeaderTooltip: (tooltipHostProps: ITooltipHostProps) => <TooltipHost {...tooltipHostProps} />
        })}
      </Sticky>
    );
  }

  private goToNewItem = () => {
    window.open(this.newURL);
  }

  private goToEditItem = () => {
    window.open(this.props.editURL + this.selectedItemID)
  }

  private checkRender = (item) => {
    const { viewColumnName, dispURL } = this.props;
    if (item.ID) {
      return <Link className={styles.customLink} href={dispURL + item.ID} target="_blank">{item[viewColumnName]}</Link>
    } else {
      return item[viewColumnName];
    }
  }

  private searchBox = (ev, value) => {
    value = value.toLowerCase()
    var rows = this.globalRows.filter(row => {
      var ret = false;
      for (var property in row) {
        if (row.hasOwnProperty(property)) {
          if (row[property] && row[property].toString().toLowerCase().includes(value)) {
            ret = true;
          }
        }
      }
      return ret;
    })
    this.setState({
      stateRows: rows
    })
  }

  private copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
  }

  private exportExcel = () => {
    var columns = [];
    var rows = [];
    this.state.stateColumns.forEach(column => {
      columns.push(column.name);
    });
    rows.push(columns);
    this.state.stateRows.forEach(item => {
      var row = [];
      this.state.stateColumns.forEach(column => {
        row.push(item[column.key]);
      });
      rows.push(row);
    });
    var worksheet = XLSX.utils.aoa_to_sheet(rows);
    var new_workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(new_workbook, worksheet, "Exported Data");
    XLSX.writeFile(new_workbook, "Exported_" + ((new Date()).getTime() / 1000) + ".xlsx");
  }

}
