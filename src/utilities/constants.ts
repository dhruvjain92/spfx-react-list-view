export class Constants {
    static readonly BooleanValues = {
        Yes: "Yes",
        No: "No"
    }
    static readonly MultipleValuesSeparator = ", ";
    static readonly NoValue = "";
    static readonly ListFormURLs = {
        NewForm: "NewForm.aspx",
        EditForm: "DispForm.aspx?ID=",
        ViewForm: "EditForm.aspx?ID="
    }
    static readonly HiddenFilter = "_Hidden";
    static readonly AttachmentsType = "Attachments";
    static readonly PropertyPaneLabels = {
        siteUrl: "Site URL",
        listName: "List Name",
        internalListName: "Internal List Name",
        multiSelect: "List Columns to show in grid",
        filterQuery: "Filter Query",
        ColumnNumberForContext: "Internal Column Name for View Link",
        showNewButton: "Show New Button",
        showEditButton: "Show Edit Button",
        showExportButton: "Show Export Button",
        webpartHeight: "Webpart Height (Only numeric values allowed. Unit is pixels)",
        columnFormatter: "Column formatting (Beta feature)"
    }
}