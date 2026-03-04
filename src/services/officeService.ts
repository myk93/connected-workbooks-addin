import { SelectionData, TableInfo, LinkedQuery } from "../types/addin";
import { extractMFormula } from "./mashupExtractor";

/**
 * All Office.js interactions are isolated here so they can be mocked in tests.
 */

export async function getSelectedRange(): Promise<SelectionData> {
    return Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(["values", "address", "rowCount", "columnCount"]);
        await context.sync();

        const { tableInfo, linkedQuery } = await detectTableContext(context, range);

        return {
            values: range.values as (string | number | boolean)[][],
            address: range.address,
            rowCount: range.rowCount,
            columnCount: range.columnCount,
            tableInfo,
            linkedQuery,
        };
    });
}

/**
 * Detects whether `range` overlaps a table, and if so:
 *   - loads the full table data for the "full table" export option
 *   - checks workbook.queries for a Power Query backing that table (ExcelApi 1.14)
 *   - decodes the M formula from the DataMashup custom XML
 *
 * All steps are wrapped in try/catch so older ExcelApi versions degrade silently.
 */
async function detectTableContext(
    context: Excel.RequestContext,
    range: Excel.Range
): Promise<{ tableInfo?: TableInfo; linkedQuery?: LinkedQuery }> {
    try {
        // Requires ExcelApi 1.9
        const tables = range.getTables(false /* intersecting, not fully contained */);
        tables.load("items/name");
        await context.sync();

        if (tables.items.length === 0) return {};

        const table = tables.items[0];

        // Load full table range (includes header row)
        const tableRange = table.getRange();
        tableRange.load(["address", "values", "rowCount", "columnCount"]);
        await context.sync();

        const isSubset =
            tableRange.rowCount > range.rowCount ||
            tableRange.columnCount > range.columnCount;

        const tableInfo: TableInfo = {
            name: table.name,
            isSubset,
            fullValues: tableRange.values as (string | number | boolean)[][],
            fullAddress: tableRange.address,
            fullRowCount: tableRange.rowCount,
            fullColumnCount: tableRange.columnCount,
        };

        // Detect linked Power Query (requires ExcelApi 1.14)
        let linkedQuery: LinkedQuery | undefined;
        try {
            const query = context.workbook.queries.getItem(table.name);
            query.load("name");
            await context.sync();

            const formula = await extractMFormula(context, query.name);
            if (formula !== undefined) {
                linkedQuery = { name: query.name, formula };
            }
        } catch {
            // No matching Power Query — plain table, ignore
        }

        return { tableInfo, linkedQuery };
    } catch {
        return {};
    }
}
