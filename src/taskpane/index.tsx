/* global document, Office, module, require, HTMLElement */
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import * as React from "react";
import { createRoot } from "react-dom/client";
import { IntlProvider } from "react-intl";
import App from "./components/App";
import DialogContextProvider from "./context/DialogContext";
import English from "./localization/en.json";
import French from "./localization/fr.json";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  const displayLanguage = Office.context.displayLanguage.toLowerCase();

  let language = "en";

  if (displayLanguage === "fr" || displayLanguage === "fr-fr" || displayLanguage === "fr-ca") {
    language = "fr";
  }

  root?.render(
    <IntlProvider
      locale={language}
      defaultLocale="en"
      messages={language === "fr" ? French : English}
    >
      <FluentProvider theme={webLightTheme}>
        <DialogContextProvider>
          <App />
        </DialogContextProvider>
      </FluentProvider>
    </IntlProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
