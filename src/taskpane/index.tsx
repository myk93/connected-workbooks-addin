import React from "react";
import ReactDOM from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import App from "./components/App";

/* Office.onReady fires once the Office.js runtime is initialized.
   We wait for it before mounting React to ensure Office APIs are available. */
Office.onReady(() => {
    const container = document.getElementById("root");
    if (!container) {
        throw new Error("Root element not found");
    }

    const root = ReactDOM.createRoot(container);
    root.render(
        <React.StrictMode>
            <FluentProvider theme={webLightTheme}>
                <App />
            </FluentProvider>
        </React.StrictMode>
    );
});
