export type ExportMode = "download" | "openInWeb" | "powerQuery" | "linkTable";

/** Determines what data is used when the Export dialog is confirmed */
export type ExportScope = "selection" | "fullTable" | "powerQuery";

export interface LinkedQuery {
    name: string;
    formula: string;   // raw M expression, no section wrapper
}

export interface TableInfo {
    name: string;
    isSubset: boolean;               // true when selection < full table
    fullValues: (string | number | boolean)[][];
    fullAddress: string;
    fullRowCount: number;
    fullColumnCount: number;
}

export interface SelectionData {
    values: (string | number | boolean)[][];
    address: string;       // e.g. "Sheet1!A1:D5"
    rowCount: number;
    columnCount: number;
    tableInfo?: TableInfo;       // present when selection overlaps a table
    linkedQuery?: LinkedQuery;   // present when that table has a Power Query
}

export interface AppState {
    mode: ExportMode;
    promoteHeaders: boolean;
    queryMashup: string;
    queryName: string;
    refreshOnOpen: boolean;
    workbookUrl: string | null;  // linkTable: source workbook URL
    linkQueryName: string;       // linkTable: name for the generated query
}
