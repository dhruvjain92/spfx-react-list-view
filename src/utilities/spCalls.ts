import axios from 'axios';
import { ColumnTypes } from './columnTypes';
export class SPCalls {
    static getListColumns(siteUrl, listName) {
        var promise = new Promise((resolve, reject) => {
            var axiosConfig = {
                headers: {
                    "Accept": "application/json;odata=verbose"
                }
            }
            axios.get(siteUrl + "/_api/web/lists/getbytitle('" + listName + "')/fields?$filter=Hidden eq false", axiosConfig).then(data => {
                debugger;
                resolve(data.data.d.results);
            })
        })
        return promise;
    }

    static getListItems(siteUrl, listName, listColumns, filterQuery) {
        var columnNames = [];
        var expandColumns = [];
        listColumns.forEach(item => {
            if (item.InternalName[0] == "_") {
                item.InternalName = "OData_" + item.InternalName;
            }
            if (item.TypeAsString == ColumnTypes.User || item.TypeAsString == ColumnTypes.MultipleUser) {
                expandColumns.push(item.InternalName);
                columnNames.push(item.InternalName + "/Title")
            } if (item.TypeAsString == ColumnTypes.Lookup || item.TypeAsString == ColumnTypes.MultipleLookup) {
                expandColumns.push(item.InternalName);
                columnNames.push(item.InternalName + "/" + item.LookupField)
            } else if (item.TypeAsString == ColumnTypes.Taxonomy || item.TypeAsString == ColumnTypes.MultipleTaxonomy) {
                expandColumns.push("TaxCatchAll");
                columnNames.push("TaxCatchAll/ID")
                columnNames.push("TaxCatchAll/Term")
                columnNames.push(item.InternalName)
            }
            else {
                columnNames.push(item.InternalName);
            }
        });
        var expand = (expandColumns.length > 0) ? "&$expand=" + expandColumns.join(",") : ""
        var promise = new Promise((resolve, reject) => {
            var axiosConfig = {
                headers: {
                    "Accept": "application/json;odata=verbose"
                }
            }
            var filter = filterQuery ? "&$filter=" + filterQuery : "";
            axios.get(siteUrl + "/_api/web/lists/getbytitle('" + listName + "')/items?$select=" + columnNames.join(",") + expand + filter, axiosConfig).then(data => {
                resolve(data.data.d.results);
            })
        })
        return promise;
    }
}