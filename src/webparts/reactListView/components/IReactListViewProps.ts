export interface IReactListViewProps {
  description: string;
  columns: any[];
  rows: any[];
  newURL: string;
  dispURL: string;
  editURL: string;
  viewColumnName: string;
  showNewButton: boolean;
  showEditButton: boolean;
  showExportButton: boolean;
  wrapperHeight: number;
}
