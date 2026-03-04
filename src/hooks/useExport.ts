import { useState, useCallback } from "react";
import { AppState, ExportScope, SelectionData } from "../types/addin";
import { exportSelection } from "../services/workbookService";

interface UseExportResult {
    exporting: boolean;
    success: string | null;
    error: string | null;
    /** Pass pre-loaded selection + scope when coming from the table-scope dialog */
    run: (state: AppState, selection: SelectionData, scope: ExportScope) => Promise<void>;
    clearStatus: () => void;
}

export function useExport(): UseExportResult {
    const [exporting, setExporting] = useState(false);
    const [success, setSuccess] = useState<string | null>(null);
    const [error, setError] = useState<string | null>(null);

    const clearStatus = useCallback(() => {
        setSuccess(null);
        setError(null);
    }, []);

    const run = useCallback(async (
        state: AppState,
        selection: SelectionData,
        scope: ExportScope
    ) => {
        setExporting(true);
        setSuccess(null);
        setError(null);

        try {
            await exportSelection(selection, state, scope);

            const isDownload =
                scope === "powerQuery" ||
                state.mode === "powerQuery" ||
                (scope !== "fullTable" && state.mode === "download") ||
                (scope === "fullTable" && state.mode === "download");
            setSuccess(isDownload
                ? "Workbook downloaded successfully."
                : "Workbook opened in Excel for the Web."
            );
        } catch (err) {
            setError(err instanceof Error ? err.message : String(err));
        } finally {
            setExporting(false);
        }
    }, []);

    return { exporting, success, error, run, clearStatus };
}
