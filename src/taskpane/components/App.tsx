import React, { useEffect, useState, useCallback } from "react";
import { makeStyles, tokens, Text, Divider } from "@fluentui/react-components";
import { ModeSelector } from "./ModeSelector";
import { HeaderToggle } from "./HeaderToggle";
import { SelectionPreview } from "./SelectionPreview";
import { PowerQueryPanel, validateQueryName } from "./PowerQueryPanel";
import { LinkTablePanel } from "./LinkTablePanel";
import { ExportButton } from "./ExportButton";
import { StatusBar } from "./StatusBar";
import { TableScopeDialog } from "./TableScopeDialog";
import { useSelection } from "../../hooks/useSelection";
import { useExport } from "../../hooks/useExport";
import { AppState, ExportMode, ExportScope, SelectionData } from "../../types/addin";
import { getSelectedRange, getWorkbookUrl } from "../../services/officeService";

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
    const [linkQueryName, setLinkQueryName] = useState("Query1");

    // Workbook URL (null if not saved to SharePoint/OneDrive)
    const [workbookUrl, setWorkbookUrl] = useState<string | null>(null);
    useEffect(() => {
        setWorkbookUrl(getWorkbookUrl());
    }, []);

    // Table-scope dialog state
    const [dialogOpen, setDialogOpen] = useState(false);
    const [pendingExport, setPendingExport] = useState<SelectionData | null>(null);
    const [loadingForDialog, setLoadingForDialog] = useState(false);

    // Auto-populate PQ + Link Table fields when landing on a PQ-backed table
    const linkedQueryName = selection?.linkedQuery?.name;
    useEffect(() => {
        if (selection?.linkedQuery) {
            setQueryName(selection.linkedQuery.name);
            setQueryMashup(selection.linkedQuery.formula);
            setMode("powerQuery");
        }
    // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [linkedQueryName]);

    // Auto-set link query name to match the detected table name
    const tableName = selection?.tableInfo?.name;
    useEffect(() => {
        if (tableName) setLinkQueryName(tableName);
    }, [tableName]);

    const appState: AppState = {
        mode, promoteHeaders, queryMashup, queryName, refreshOnOpen,
        workbookUrl, linkQueryName,
    };

    const isExportDisabled = (() => {
        if (!selection || selection.rowCount === 0) return true;
        if (mode === "powerQuery")
            return !queryMashup.trim() || validateQueryName(queryName) !== null;
        if (mode === "linkTable")
            return !workbookUrl || !selection.tableInfo || validateQueryName(linkQueryName) !== null;
        return false;
    })();

    // ── Export click ───────────────────────────────────────────────────────────
    const handleExportClick = useCallback(async () => {
        clearStatus();
        setLoadingForDialog(true);
        try {
            const fresh = await getSelectedRange();

            // linkTable and powerQuery modes never show the scope dialog —
            // the user has already made an explicit mode choice.
            const hasExtraOption =
                mode !== "powerQuery" &&
                mode !== "linkTable" &&
                fresh.tableInfo &&
                (fresh.tableInfo.isSubset || fresh.linkedQuery);

            if (hasExtraOption) {
                setPendingExport(fresh);
                setDialogOpen(true);
            } else {
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

            <ModeSelector
                value={mode}
                onChange={setMode}
                isShareable={workbookUrl !== null}
            />

            <SelectionPreview
                selection={selection}
                loading={selectionLoading}
                error={selectionError}
                onRefresh={refresh}
            />

            {mode !== "linkTable" && (
                <HeaderToggle checked={promoteHeaders} onChange={setPromoteHeaders} />
            )}

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

            {mode === "linkTable" && (
                <LinkTablePanel
                    workbookUrl={workbookUrl}
                    tableInfo={selection?.tableInfo}
                    queryName={linkQueryName}
                    onQueryNameChange={setLinkQueryName}
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
