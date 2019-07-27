# React List Grid View
React List Grid View is a SPFx webpart built to display a list view from any site collection or subsite.

## Features
 - Displays list view from any site collection and any list to which the user has experience.
 - Every component is from Office Fabric UI and provides responsive interface.
 - Ability to export data into excel
 - Extensive configurable options
 - Column Formatting

## Webpart properties
Following properties are available in the webpart configuration.

| **Property** | **Description** |
|-----------------|--------------|
|Site URL|URL of the site where the list is located. It should have a trailing slash.|
|List Name|Name of the list.
|List Columns to show in grid| This shows the list columns. This is generated dynamically based on site URL and list name.
|Filter Query|This fields contain the filter for the list. For example, if you have to match field *FIELD_1* with value *VALUE_1*, the query will be *FIELD_1 eq VALUE_1*.
|Internal Column Name for View Link| This is the internal column name for the view link. Leave blank to disable view link. ID column must be selected in list columns for view link to work.|
|Show New Button| Shows the new item button. This opens the new item form in a new tab.|
|Show Edit Button| Shows the edit item button. This opens the edit item form in a new tab.|
|Show Export Button| Shows the export button. This button downloads the excel with data from current table.|
|Webpart Height| This is the vertical height of the webpart after which scrollbar will be visible.|
|Column formatting| This is a beta feature with JavaScript formatting for the columns. This section will be updated in the future|

## Known Issues
 - Multiple Metadata columns in the same list causes issues. 
 - ID column is a must for view and edit links for now. It will be removed in future.
 - List name and list's internal name are assumed to be same. 
