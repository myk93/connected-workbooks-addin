import React from "react";
import {
    MessageBar,
    MessageBarBody,
    MessageBarActions,
    Button,
    makeStyles,
} from "@fluentui/react-components";
import { DismissRegular } from "@fluentui/react-icons";

const useStyles = makeStyles({
    root: {
        marginTop: "12px",
    },
});

interface StatusBarProps {
    success: string | null;
    error: string | null;
    onDismiss: () => void;
}

export const StatusBar: React.FC<StatusBarProps> = ({ success, error, onDismiss }) => {
    const styles = useStyles();

    if (!success && !error) return null;

    return (
        <div className={styles.root}>
            <MessageBar intent={success ? "success" : "error"}>
                <MessageBarBody>{success ?? error}</MessageBarBody>
                <MessageBarActions
                    containerAction={
                        <Button
                            aria-label="Dismiss"
                            appearance="transparent"
                            icon={<DismissRegular />}
                            onClick={onDismiss}
                        />
                    }
                />
            </MessageBar>
        </div>
    );
};
