import React from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { ArrowUpload20Regular } from "@fluentui/react-icons";

interface ExportButtonProps {
    onClick: () => void;
    loading: boolean;
    disabled: boolean;
}

export const ExportButton: React.FC<ExportButtonProps> = ({
    onClick,
    loading,
    disabled,
}) => {
    return (
        <Button
            appearance="primary"
            icon={loading ? <Spinner size="tiny" /> : <ArrowUpload20Regular />}
            onClick={onClick}
            disabled={disabled || loading}
            style={{ width: "100%" }}
        >
            {loading ? "Exporting…" : "Export"}
        </Button>
    );
};
