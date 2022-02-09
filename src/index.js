import DialogComponent from "./components/DialogComponent";

import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import * as React from "react";
import * as ReactDOM from "react-dom";
/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "myapp Add-in";

console.log("index.js called")

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} />
    </AppContainer>,
    document.getElementById("container_dialog")
  );
};

/* Render application after Office initializes */
Office.initialize = () => {
    console.log("Office.initialized in index.js !!! ")
    isOfficeInitialized = true;
    render(DialogComponent);
};

/* Initial render showing a progress bar */
// render(DialogComponent);

// if (module.hot) {
//   module.hot.accept("./components/DialogComponent", () => {
//     const NextApp = require("./components/DialogComponent").default;
//     render(NextApp);
//   });
// }
