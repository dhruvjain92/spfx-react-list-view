import { ColumnTypes } from "./columnTypes";
import { Constants } from "./constants";

export class ColumnParser {
    static getColumnValue(item, column) {
        var columnValue = "";
        switch (column.TypeAsString) {
            case ColumnTypes.User:
                columnValue = this.userParser(item, column);
                break;
            case ColumnTypes.MultipleUser:
                columnValue = this.multipleUserParser(item, column);
                break;
            case ColumnTypes.DateTime:
                columnValue = this.datetimeParser(item, column);
                break;
            case ColumnTypes.Boolean:
                columnValue = this.booleanParser(item, column);
                break;
            case ColumnTypes.Choice:
                columnValue = this.choiceParser(item, column);
                break;
            case ColumnTypes.MultipleChoices:
                columnValue = this.multipleChoicesParser(item, column);
                break;
            case ColumnTypes.Lookup:
                columnValue = this.lookupParser(item, column);
                break;
            case ColumnTypes.MultipleLookup:
                columnValue = this.multipleLookupParser(item, column);
                break;
            case ColumnTypes.URL:
                columnValue = this.urlParser(item, column);
                break;
            case ColumnTypes.Taxonomy:
                columnValue = this.taxonomyParser(item, column);
                break;
            case ColumnTypes.MultipleTaxonomy:
                columnValue = this.taxonomyParser(item, column);
                break;
            default:
                columnValue = this.defaultParser(item, column);
                break;
        }
        return columnValue;
    }

    static userParser(item, column) {
        return item[column.InternalName] ? item[column.InternalName].Title : Constants.NoValue
    }

    static multipleUserParser(item, column) {
        var columnValue = "";
        if (item[column.InternalName].results) {
            var users = [];
            item[column.InternalName].results.forEach(user => {
                users.push(user.Title);
            })
            columnValue = users.join(Constants.MultipleValuesSeparator);
        } else {
            columnValue = Constants.NoValue;
        }
        return columnValue;
    }

    static datetimeParser(item, column) {
        return item[column.InternalName] ? (new Date(item[column.InternalName])).toLocaleDateString() : Constants.NoValue;
    }

    static booleanParser(item, column) {
        return item[column.InternalName] ? Constants.BooleanValues.Yes : Constants.BooleanValues.Yes;
    }

    static choiceParser(item, column) {
        return item[column.InternalName];
    }

    static multipleChoicesParser(item, column) {
        return item[column.InternalName] ? (item[column.InternalName].results ? item[column.InternalName].results.join(Constants.MultipleValuesSeparator) : Constants.NoValue) : Constants.NoValue;
    }

    static lookupParser(item, column) {
        return item[column.InternalName] ? item[column.InternalName][column.LookupField] : "";
    }

    static multipleLookupParser(item, column) {
        var columnValue = "";
        if (item[column.InternalName].results) {
            var lookupValues = [];
            item[column.InternalName].results.forEach(lookupValue => {
                lookupValues.push(lookupValue[column.LookupField]);
            })
            columnValue = lookupValues.join(Constants.MultipleValuesSeparator);
        } else {
            columnValue = Constants.NoValue;
        }
        return columnValue;
    }

    static urlParser(item, column) {
        return item[column.InternalName] ? item[column.InternalName].Url : Constants.NoValue;
    }

    //Currently handles one MMD column in a list. Will need to update this for multiple MMD columns
    // Match item results IDs with tax result IDs.
    static taxonomyParser(item, column) {
        var columnValue = "";
        if (item["TaxCatchAll"].results) {
            var taxValues = [];
            item["TaxCatchAll"].results.forEach(taxValue => {
                taxValues.push(taxValue["Term"]);
            })
            columnValue = taxValues.join(Constants.MultipleValuesSeparator);
        } else {
            columnValue = Constants.NoValue;
        }
        return columnValue;
    }

    static defaultParser(item, column) {
        return item[column.InternalName];
    }

}