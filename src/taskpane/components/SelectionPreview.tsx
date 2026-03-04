import React from "react";
import {
    Button,
    Text,
    Spinner,
    Badge,
    makeStyles,
    tokens,
} from "@fluentui/react-components";
import { ArrowClockwise16Regular, DataFunnel20Regular } from "@fluentui/react-icons";
import { SelectionData } from "../../types/addin";

const useStyles = makeStyles({
    root: {
        marginBottom: "12px",
        padding: "8px 12px",
        backgroundColor: tokens.colorNeutralBackground2,
        borderRadius: tokens.borderRadiusMedium,
        display: "flex",
        alignItems: "center",
        gap: "8px",
        minHeight: "36px",
    },
    info: {
        flex: 1,
        display: "flex",
        flexDirection: "column",
        gap: "4px",
    },
    address: {
        fontWeight: tokens.fontWeightSemibold,
        fontFamily: "monospace",
    },
    dims: {
        color: tokens.colorNeutralForeground3,
        marginLeft: "8px",
    },
    pqBadge: {
        display: "flex",
        alignItems: "center",
        gap: "4px",
    },
});

interface SelectionPreviewProps {
    selection: SelectionData | null;
    loading: boolean;
    error: string | null;
    onRefresh: () => void;
}

export const SelectionPreview: React.FC<SelectionPreviewProps> = ({
    selection,
    loading,
    error,
    onRefresh,
}) => {
    const styles = useStyles();

    let content: React.ReactNode;

    if (loading) {
        content = <Spinner size="tiny" label="Reading selection…" />;
    } else if (error) {
        content = <Text style={{ color: "var(--colorStatusDangerForeground1)" }}>Error: {error}</Text>;
    } else if (selection) {
        content = (
            <>
                <Text>
                    <span className={styles.address}>{selection.address}</span>
                    <span className={styles.dims}>
                        {selection.rowCount} rows × {selection.columnCount} cols
                    </span>
                </Text>
                {selection.linkedQuery && (
                    <div className={styles.pqBadge}>
                        <DataFunnel20Regular style={{ color: tokens.colorBrandForeground1 }} />
                        <Badge appearance="tint" color="brand" size="small">
                            Power Query: {selection.linkedQuery.name}
                        </Badge>
                    </div>
                )}
            </>
        );
    } else {
        content = <Text style={{ color: "var(--colorNeutralForeground3)" }}>No selection — click Refresh</Text>;
    }

    return (
        <div className={styles.root}>
            <div className={styles.info}>{content}</div>
            <Button
                icon={<ArrowClockwise16Regular />}
                appearance="subtle"
                size="small"
                onClick={onRefresh}
                aria-label="Refresh selection"
            />
        </div>
    );
};
