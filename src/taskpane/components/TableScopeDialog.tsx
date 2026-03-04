import React, { useState } from "react";
import {
    Dialog,
    DialogSurface,
    DialogTitle,
    DialogBody,
    DialogContent,
    DialogActions,
    Button,
    RadioGroup,
    Radio,
    Text,
    makeStyles,
    tokens,
} from "@fluentui/react-components";
import { TableInfo, LinkedQuery, ExportScope } from "../../types/addin";

const useStyles = makeStyles({
    description: {
        color: tokens.colorNeutralForeground2,
        marginBottom: "12px",
        display: "block",
    },
    radioGroup: {
        display: "flex",
        flexDirection: "column",
        gap: "4px",
    },
    radioLabel: {
        fontSize: tokens.fontSizeBase300,
    },
    dimLabel: {
        color: tokens.colorNeutralForeground3,
        fontSize: tokens.fontSizeBase200,
        marginLeft: "4px",
    },
});

interface TableScopeDialogProps {
    open: boolean;
    tableInfo: TableInfo;
    selectionDims: { rowCount: number; columnCount: number };
    linkedQuery?: LinkedQuery;
    onConfirm: (scope: ExportScope) => void;
    onCancel: () => void;
}

export const TableScopeDialog: React.FC<TableScopeDialogProps> = ({
    open,
    tableInfo,
    selectionDims,
    linkedQuery,
    onConfirm,
    onCancel,
}) => {
    const styles = useStyles();

    // Default to the "richest" available option
    const defaultScope: ExportScope = linkedQuery
        ? "powerQuery"
        : tableInfo.isSubset
        ? "fullTable"
        : "selection";

    const [scope, setScope] = useState<ExportScope>(defaultScope);

    // Reset to default whenever the dialog re-opens
    React.useEffect(() => {
        if (open) setScope(defaultScope);
    // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [open]);

    return (
        <Dialog
            open={open}
            onOpenChange={(_, data) => { if (!data.open) onCancel(); }}
        >
            <DialogSurface>
                <DialogTitle>Export from "{tableInfo.name}"</DialogTitle>
                <DialogBody>
                    <DialogContent>
                        <Text className={styles.description}>
                            Your selection is part of a table. Choose what to export:
                        </Text>

                        <RadioGroup
                            className={styles.radioGroup}
                            value={scope}
                            onChange={(_, data) => setScope(data.value as ExportScope)}
                        >
                            <Radio
                                value="selection"
                                label={
                                    <span>
                                        <span className={styles.radioLabel}>Selected cells only</span>
                                        <span className={styles.dimLabel}>
                                            {selectionDims.rowCount} rows × {selectionDims.columnCount} cols
                                        </span>
                                    </span>
                                }
                            />

                            {tableInfo.isSubset && (
                                <Radio
                                    value="fullTable"
                                    label={
                                        <span>
                                            <span className={styles.radioLabel}>Entire table</span>
                                            <span className={styles.dimLabel}>
                                                {tableInfo.fullRowCount} rows × {tableInfo.fullColumnCount} cols
                                            </span>
                                        </span>
                                    }
                                />
                            )}

                            {linkedQuery && (
                                <Radio
                                    value="powerQuery"
                                    label={
                                        <span>
                                            <span className={styles.radioLabel}>
                                                Re-export as Power Query
                                            </span>
                                            <span className={styles.dimLabel}>
                                                "{linkedQuery.name}"
                                            </span>
                                        </span>
                                    }
                                />
                            )}
                        </RadioGroup>
                    </DialogContent>

                    <DialogActions>
                        <Button appearance="subtle" onClick={onCancel}>
                            Cancel
                        </Button>
                        <Button appearance="primary" onClick={() => onConfirm(scope)}>
                            Export
                        </Button>
                    </DialogActions>
                </DialogBody>
            </DialogSurface>
        </Dialog>
    );
};
