import React, { useMemo } from "react";
import {
    Field,
    Textarea,
    Input,
    Switch,
    Text,
    makeStyles,
    tokens,
} from "@fluentui/react-components";

// Mirrors pqUtils.validateQueryName constraints (max 80 chars for UX, no " . or control chars)
const MAX_QUERY_NAME_LENGTH = 80;
const INVALID_QUERY_NAME_RE = /[".\\x00-\\x1F\\x7F-\\x9F]/;

function validateQueryName(name: string): string | null {
    if (!name.trim()) return "Query name is required.";
    if (name.length > MAX_QUERY_NAME_LENGTH)
        return `Query name must be ${MAX_QUERY_NAME_LENGTH} characters or fewer.`;
    if (/[".]/.test(name) || /[\x00-\x1F\x7F-\x9F]/.test(name))
        return 'Query name cannot contain quotes ("), periods (.), or control characters.';
    return null;
}

const useStyles = makeStyles({
    root: {
        display: "flex",
        flexDirection: "column",
        gap: "12px",
        marginBottom: "12px",
    },
    hint: {
        color: tokens.colorNeutralForeground3,
        fontSize: tokens.fontSizeBase200,
        marginTop: "2px",
    },
});

interface PowerQueryPanelProps {
    queryMashup: string;
    onQueryMashupChange: (value: string) => void;
    queryName: string;
    onQueryNameChange: (value: string) => void;
    refreshOnOpen: boolean;
    onRefreshOnOpenChange: (value: boolean) => void;
}

export const PowerQueryPanel: React.FC<PowerQueryPanelProps> = ({
    queryMashup,
    onQueryMashupChange,
    queryName,
    onQueryNameChange,
    refreshOnOpen,
    onRefreshOnOpenChange,
}) => {
    const styles = useStyles();
    const queryNameError = useMemo(() => validateQueryName(queryName), [queryName]);

    return (
        <div className={styles.root}>
            <Field
                label="M Expression"
                hint={
                    <Text className={styles.hint}>
                        Enter a raw M expression (e.g. <code>= Web.Page(...)</code>). The{" "}
                        <code>section Section1;</code> wrapper is added automatically.
                    </Text>
                }
                required
            >
                <Textarea
                    value={queryMashup}
                    onChange={(_, data) => onQueryMashupChange(data.value)}
                    rows={8}
                    resize="vertical"
                    placeholder='let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content]
in
    Source'
                />
            </Field>

            <Field
                label="Query name"
                validationMessage={queryNameError ?? undefined}
                validationState={queryNameError ? "error" : "none"}
                required
            >
                <Input
                    value={queryName}
                    onChange={(_, data) => onQueryNameChange(data.value)}
                    placeholder="Query1"
                    maxLength={MAX_QUERY_NAME_LENGTH}
                />
            </Field>

            <Switch
                label="Refresh on open"
                checked={refreshOnOpen}
                onChange={(_, data) => onRefreshOnOpenChange(data.checked)}
            />
        </div>
    );
};

export { validateQueryName };
