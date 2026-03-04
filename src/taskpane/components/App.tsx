import React, { useEffect, useState, useCallback } from "react";
import { makeStyles, tokens, Text, Divider } from "@fluentui/react-components";
import { ModeSelector } from "./ModeSelector";
import { HeaderToggle } from "./HeaderToggle";
import { SelectionPreview } from "./SelectionPreview";
import { PowerQueryPanel, validateQueryName } from "./PowerQueryPanel";
import { ExportButton } from "./ExportButton";
import { StatusBar } from "./StatusBar";
import { TableScopeDialog } from "./TableScopeDialog";
import { useSelection } from "../../hooks/useSelection";
import { useExport } from "../../hooks/useExport";
import { AppState, ExportMode, ExportScope, SelectionData } from "../../types/addin";
import { getSelectedRange } from "../../services/officeService";

const useStyles = makeStyles({
    root: {
        padding: "16px",
        display: "flex",
        flexDirection: "column",
        height: "100vh",
        boxSizing: "border-box",
        fontFamily: tokens.fontFamilyBase,
    },
    header: {
        marginBottom: "16px",
    },
    title: {
        fontSize: tokens.fontSizeBase500,
        fontWeight: tokens.fontWeightSemibold,
        display: "block",
        marginBottom: "4px",
    },
    subtitle: {
        color: tokens.colorNeutralForeground3,
        fontSize: tokens.fontSizeBase200,
    },
    divider: {
        marginBottom: "16px",
    },
    footer: {
        marginTop: "auto",
    },
});

export default function App() {
    const styles = useStyles();

    const { selection, loading: selectionLoading, error: selectionError, refresh } = useSelection();
    const { exporting, success, error: exportError, run, clearStatus } = useExport();

    // Form state
    const [mode, setMode] = useState<ExportMode>("download");
    const [promoteHeaders, setPromoteHeaders] = useState(true);
    const [queryMashup, setQueryMashup] = useState("");
    const [queryName, setQueryName] = useState("Query1");
    const [refreshOnOpen, setRefreshOnOpen] = useState(true);

    // Table-scope dialog state
    const [dialogOpen, setDialogOpen] = useState(false);
    const [pendingExport, setPendingExport] = useState<SelectionData | null>(null);
    const [loadingForDialog, setLoadingForDialog] = useState(false);

    // Auto-populate Power Query fields when landing on a new PQ-backed table
    const linkedQueryName = selection?.linkedQuery?.name;
    useEffect(() => {
        if (selection?.linkedQuery) {
            setQueryName(selection.linkedQuery.name);
            setQueryMashup(selection.linkedQuery.formula);
            setMode("powerQuery");
        }
    // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [linkedQueryName]);

    const appState: AppState = { mode, promoteHeaders, queryMashup, queryName, refreshOnOpen };

    const isExportDisabled =
        (!selection || selection.rowCount === 0) ||
        (mode === "powerQuery" && (!queryMashup.trim() || validateQueryName(queryName) !== null));

    // ── Export click ───────────────────────────────────────────────────────────
    const handleExportClick = useCallback(async () => {
        clearStatus();
        setLoadingForDialog(true);
        try {
            // Always do a fresh read so dialog shows accurate row counts
            const fresh = await getSelectedRange();

            // Show dialog if the selection is inside a table AND there is at least
            // one extra option to offer (subset → full table, or linked PQ)
            // Never interrupt PQ tab exports with the scope dialog —
            // the user has already explicitly chosen Power Query mode.
            const hasExtraOption =
                mode !== "powerQuery" &&
                fresh.tableInfo &&
                (fresh.tableInfo.isSubset || fresh.linkedQuery);

            if (hasExtraOption) {
                setPendingExport(fresh);
                setDialogOpen(true);
            } else {
                // No dialog needed — export directly using the current mode tab
                await run(appState, fresh, "selection");
            }
        } catch (err) {
            // Rare: getSelectedRange itself failed
        } finally {
            setLoadingForDialog(false);
        }
    // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [appState, run, clearStatus]);

    const handleDialogConfirm = useCallback((scope: ExportScope) => {
        setDialogOpen(false);
        if (pendingExport) {
            run(appState, pendingExport, scope);
            setPendingExport(null);
        }
    // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [appState, pendingExport, run]);

    const handleDialogCancel = useCallback(() => {
        setDialogOpen(false);
        setPendingExport(null);
    }, []);

    return (
        <div className={styles.root}>
            <div className={styles.header}>
                <Text className={styles.title}>Connected Workbooks</Text>
                <Text className={styles.subtitle}>Export your Excel selection</Text>
            </div>

            <Divider className={styles.divider} />

            <ModeSelector value={mode} onChange={setMode} />

            <SelectionPreview
                selection={selection}
                loading={selectionLoading}
                error={selectionError}
                onRefresh={refresh}
            />

            <HeaderToggle checked={promoteHeaders} onChange={setPromoteHeaders} />

            {mode === "powerQuery" && (
                <PowerQueryPanel
                    queryMashup={queryMashup}
                    onQueryMashupChange={setQueryMashup}
                    queryName={queryName}
                    onQueryNameChange={setQueryName}
                    refreshOnOpen={refreshOnOpen}
                    onRefreshOnOpenChange={setRefreshOnOpen}
                />
            )}

            <div className={styles.footer}>
                <ExportButton
                    onClick={handleExportClick}
                    loading={exporting || loadingForDialog}
                    disabled={isExportDisabled}
                />

                <StatusBar
                    success={success}
                    error={exportError}
                    onDismiss={clearStatus}
                />
            </div>

            {/* Table-scope dialog — rendered outside the footer so it isn't clipped */}
            {pendingExport?.tableInfo && (
                <TableScopeDialog
                    open={dialogOpen}
                    tableInfo={pendingExport.tableInfo}
                    selectionDims={{
                        rowCount: pendingExport.rowCount,
                        columnCount: pendingExport.columnCount,
                    }}
                    linkedQuery={pendingExport.linkedQuery}
                    onConfirm={handleDialogConfirm}
                    onCancel={handleDialogCancel}
                />
            )}
        </div>
    );
}
