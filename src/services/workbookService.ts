import { workbookManager } from "@microsoft/connected-workbooks";
import type { Grid } from "@microsoft/connected-workbooks";
import { AppState, ExportScope, SelectionData } from "../types/addin";

const { generateTableWorkbookFromGrid, generateSingleQueryWorkbook, downloadWorkbook, openInExcelWeb } = workbookManager;

function buildGrid(data: (string | number | boolean)[][], promoteHeaders: boolean): Grid {
    return { data, config: { promoteHeaders, adjustColumnNames: true } };
}

/**
 * Runs the appropriate export path for the given scope:
 *
 *  "selection"  — use the currently selected cells with the active mode tab
 *  "fullTable"  — use the full table range (always treats row 1 as headers)
 *  "powerQuery" — re-export the linked PQ query; always downloads
 */
export async function exportSelection(
    selection: SelectionData,
    state: AppState,
    scope: ExportScope
): Promise<void> {
    const { mode, promoteHeaders, queryMashup, queryName, refreshOnOpen } = state;

    // ── Power Query scope ──────────────────────────────────────────────────────
    if (scope === "powerQuery") {
        const lq = selection.linkedQuery;
        if (!lq) throw new Error("No linked Power Query found for this selection.");
        const blob = await generateSingleQueryWorkbook(
            { queryMashup: lq.formula, queryName: lq.name, refreshOnOpen },
            buildGrid(selection.values, promoteHeaders)
        );
        downloadWorkbook(blob, `${lq.name}.xlsx`);
        return;
    }

    // ── Full table scope ───────────────────────────────────────────────────────
    if (scope === "fullTable") {
        if (!selection.tableInfo) throw new Error("No table data available.");
        const blob = await generateTableWorkbookFromGrid(
            buildGrid(selection.tableInfo.fullValues, true /* table always has headers */)
        );
        if (mode === "download") {
            downloadWorkbook(blob, `${selection.tableInfo.name}.xlsx`);
        } else {
            await openInExcelWeb(blob, selection.tableInfo.name, true);
        }
        return;
    }

    // ── Selection scope (default) ──────────────────────────────────────────────
    const grid = buildGrid(selection.values, promoteHeaders);

    switch (mode) {
        case "download": {
            const blob = await generateTableWorkbookFromGrid(grid);
            downloadWorkbook(blob, "export.xlsx");
            break;
        }
        case "openInWeb": {
            const blob = await generateTableWorkbookFromGrid(grid);
            await openInExcelWeb(blob, "export", true);
            break;
        }
        case "powerQuery": {
            const blob = await generateSingleQueryWorkbook(
                { queryMashup, queryName, refreshOnOpen },
                grid
            );
            downloadWorkbook(blob, `${queryName || "export"}.xlsx`);
            break;
        }
        default: {
            const exhaustive: never = mode;
            throw new Error(`Unhandled export mode: ${exhaustive}`);
        }
    }
}
