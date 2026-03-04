import React from "react";
import { Tab, TabList, makeStyles } from "@fluentui/react-components";
import { ExportMode } from "../../types/addin";

const useStyles = makeStyles({
    root: {
        marginBottom: "12px",
    },
});

interface ModeSelectorProps {
    value: ExportMode;
    onChange: (mode: ExportMode) => void;
}

export const ModeSelector: React.FC<ModeSelectorProps> = ({ value, onChange }) => {
    const styles = useStyles();

    return (
        <div className={styles.root}>
            <TabList
                selectedValue={value}
                onTabSelect={(_, data) => onChange(data.value as ExportMode)}
            >
                <Tab value="download">Download</Tab>
                <Tab value="openInWeb">Open in Web</Tab>
                <Tab value="powerQuery">Power Query</Tab>
            </TabList>
        </div>
    );
};
