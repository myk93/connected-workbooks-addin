import React from "react";
import { Checkbox, makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
    root: {
        marginBottom: "12px",
    },
});

interface HeaderToggleProps {
    checked: boolean;
    onChange: (checked: boolean) => void;
}

export const HeaderToggle: React.FC<HeaderToggleProps> = ({ checked, onChange }) => {
    const styles = useStyles();

    return (
        <div className={styles.root}>
            <Checkbox
                label="Treat first row as headers"
                checked={checked}
                onChange={(_, data) => onChange(Boolean(data.checked))}
            />
        </div>
    );
};
