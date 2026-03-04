import React from "react";
import { Tab, TabList, Tooltip, makeStyles } from "@fluentui/react-components";
import { ExportMode } from "../../types/addin";

const useStyles = makeStyles({
    root: {
        marginBottom: "12px",
    },
});

interface ModeSelectorProps {
    value: ExportMode;
    onChange: (mode: ExportMode) => void;
    isShareable: boolean;
}

export const ModeSelector: React.FC<ModeSelectorProps> = ({ value, onChange, isShareable }) => {
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
                <Tooltip
                    content="Save the workbook to SharePoint or OneDrive to enable this mode"
                    relationship="description"
                    hideDelay={0}
                >
                    <Tab value="linkTable" disabled={!isShareable}>
                        Link Table
                    </Tab>
                </Tooltip>
            </TabList>
        </div>
    );
};
