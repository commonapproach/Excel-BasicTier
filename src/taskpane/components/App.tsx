/* global window, fetch, document, HTMLInputElement, FileReader */
import { Button, makeStyles, Spinner, tokens } from "@fluentui/react-components";
import {
  Add24Regular,
  ArrowCircleDown24Regular,
  ArrowCircleUp24Regular,
  Book24Regular,
} from "@fluentui/react-icons";
import { ArrowSync24Regular } from "@fluentui/react-icons/fonts";
import * as React from "react";
import { FormattedMessage, useIntl } from "react-intl";
import { useDialogContext } from "../context/DialogContext";
import { populateSeliGLI } from "../helpers/seliGLI";
import { importData } from "../import/import";
import {
  createSFFModuleSheetsAndTables,
  createSheetsAndTables,
  populateCodeLists,
} from "../taskpane";
import ExportDialog from "./ExportDialog";
import Header from "./Header";

interface AppProps {}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    width: "100%",
    position: "relative",
  },
  buttons_group: {
    display: "flex",
    flexDirection: "column",
    justifyContent: "center",
    alignItems: "center",
    margin: "0.7rem", // Reduced margin
    paddingBottom: "1rem", // Reduced padding
    gap: "0.7rem", // Reduced gap
  },
  button: {
    width: "160px",
  },
  overlay: {
    position: "fixed",
    top: 0,
    left: 0,
    width: "100%",
    height: "100%",
    backgroundColor: "rgba(255, 255, 255, 0.8)",
    zIndex: 1000,
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
  },
  warning_note: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightBold,
    color: tokens.colorNeutralForeground1,
    textAlign: "center",
    marginTop: "5px",
    marginBottom: "5px",
  },
  note_message: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightRegular,
    color: tokens.colorNeutralForeground1,
    textAlign: "center",
    marginTop: "5px",
    marginBottom: "5px",
  },
  link: {
    cursor: "pointer",
    color: tokens.colorBrandForeground1,
    textDecoration: "underline",
  },
});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();
  const dialog = useDialogContext();
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [isLoading, setIsLoading] = React.useState(false);
  const [isExportDialogOpen, setIsExportDialogOpen] = React.useState(false);
  const intl = useIntl();

  const handleImportData = () => {
    if (fileInputRef.current) {
      fileInputRef.current.click();
    }
  };

  const handleFileChange = async (
    event: any,
    onSuccess: (data: any) => Promise<void>,
    onError: (error: any) => void
  ): Promise<void> => {
    const file = event.target.files[0];
    if (file && (file.name.endsWith(".jsonld") || file.name.endsWith(".json"))) {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const data = JSON.parse(e.target?.result as any);
          await onSuccess(data);
        } catch (error) {
          onError(new Error(intl.formatMessage({ id: "import.messages.error.notValidJson" })));
        }
      };
      reader.readAsText(file);
    } else {
      onError(new Error(intl.formatMessage({ id: "import.messages.error.notJson" })));
    }
  };

  const downloadSampleData = async (event: React.MouseEvent) => {
    event.preventDefault();
    try {
      const url =
        "https://ontology.commonapproach.org/examples/CIDSBasicZerokitsTestData-SHARED.json";
      const response = await fetch(url);
      const data = await response.blob();

      // Create a blob URL and trigger download
      const blobUrl = window.URL.createObjectURL(data);
      const a = document.createElement("a");
      a.style.display = "none";
      a.href = blobUrl;
      a.download = "CIDSBasicZerokitsTestData-SHARED.json";
      document.body.appendChild(a);
      a.click();

      // Clean up
      window.URL.revokeObjectURL(blobUrl);
      document.body.removeChild(a);
    } catch (error) {
      dialog.showDialog(
        intl.formatMessage({ id: "generics.error" }),
        intl.formatMessage({
          id: "import.messages.error.downloadingSampleData",
          defaultMessage: "Error downloading sample data",
        })
      );
    }
  };

  const downloadSampleDataSFF = async (event: React.MouseEvent) => {
    event.preventDefault();
    try {
      const url = "https://ontology.commonapproach.org/examples/CIDSBasictestandSFFSampleData.json";
      const response = await fetch(url);
      const data = await response.blob();

      // Create a blob URL and trigger download
      const blobUrl = window.URL.createObjectURL(data);
      const a = document.createElement("a");
      a.style.display = "none";
      a.href = blobUrl;
      a.download = "CIDSBasictestandSFFSampleData.json";
      document.body.appendChild(a);
      a.click();

      // Clean up
      window.URL.revokeObjectURL(blobUrl);
      document.body.removeChild(a);
    } catch (error) {
      dialog.showDialog(
        intl.formatMessage({ id: "generics.error" }),
        intl.formatMessage({
          id: "import.messages.error.downloadingSampleData",
          defaultMessage: "Error downloading sample data",
        })
      );
    }
  };

  return (
    <div className={styles.root}>
      {isLoading && (
        <div className={styles.overlay}>
          <Spinner />
        </div>
      )}
      <ExportDialog
        isDialogOpen={isExportDialogOpen}
        setDialogOpen={setIsExportDialogOpen}
        setIslLoading={setIsLoading}
      />
      <Header />
      <div className={styles.buttons_group}>
        <input
          ref={fileInputRef}
          style={{ display: "none" }}
          title="file"
          type="file"
          onChange={async (e) => {
            setIsLoading(true);
            await handleFileChange(
              e,
              async (data) => {
                try {
                  await importData(intl, data, dialog.showDialog, setIsLoading);
                } catch (error: any) {
                  dialog.showDialog(
                    `${intl.formatMessage({ id: "generics.error" })}!`,
                    error.message || intl.formatMessage({ id: "generics.error.message" })
                  );
                } finally {
                  setIsLoading(false);
                }
              },
              (error) => {
                setIsLoading(false);
                dialog.showDialog(
                  `${intl.formatMessage({ id: "generics.error" })}!`,
                  error.message
                );
              }
            );
            // clear the file input
            if (fileInputRef.current) {
              fileInputRef.current.value = "";
            }
          }}
        />
        <Button
          content={intl.formatMessage({ id: "app.button.importData" })}
          onClick={handleImportData}
          appearance="outline"
          icon={<ArrowCircleUp24Regular />}
          iconPosition="before"
          disabled={isLoading}
          className={styles.button}
          style={{
            borderColor: "rgb(60, 174, 163)",
            color: "rgb(60, 174, 163)",
          }}
        >
          <FormattedMessage id="app.button.importData" />
        </Button>
        <Button
          content={intl.formatMessage({ id: "app.button.exportData" })}
          onClick={() => {
            setIsExportDialogOpen(true);
          }}
          appearance="outline"
          icon={<ArrowCircleDown24Regular />}
          iconPosition="before"
          disabled={isLoading}
          className={styles.button}
          style={{
            borderColor: "rgb(80, 183, 224)",
            color: "rgb(80, 183, 224)",
          }}
        >
          <FormattedMessage id="app.button.exportData" />
        </Button>
        <Button
          content={intl.formatMessage({ id: "app.button.createTables" })}
          onClick={async () => {
            setIsLoading(true);
            try {
              await createSheetsAndTables(intl);
              dialog.showDialog(
                intl.formatMessage({
                  id: "generics.success",
                  defaultMessage: "Success",
                }),
                intl.formatMessage({
                  id: "createTables.messages.success",
                  defaultMessage: "Tables created successfully",
                })
              );
            } catch (error: any) {
              dialog.showDialog(
                intl.formatMessage({
                  id: "generics.error",
                  defaultMessage: "Error",
                }),
                error.message ||
                  intl.formatMessage({
                    id: "generics.error.message",
                    defaultMessage: "Something went wrong",
                  })
              );
            } finally {
              setIsLoading(false);
            }
          }}
          appearance="outline"
          color="brand"
          icon={<Add24Regular />}
          iconPosition="before"
          disabled={isLoading}
          className={styles.button}
          style={{
            borderColor: "rgb(45, 98, 215)",
            color: "rgb(45, 98, 215)",
          }}
        >
          <FormattedMessage id="app.button.createTables" />
        </Button>
        <Button
          content={intl.formatMessage({ id: "app.button.createSFFTables" })}
          onClick={async () => {
            setIsLoading(true);
            try {
              await createSFFModuleSheetsAndTables(intl);
              dialog.showDialog(
                intl.formatMessage({
                  id: "generics.success",
                  defaultMessage: "Success",
                }),
                intl.formatMessage({
                  id: "createTables.messages.success",
                  defaultMessage: "Tables created successfully",
                })
              );
            } catch (error: any) {
              dialog.showDialog(
                intl.formatMessage({
                  id: "generics.error",
                  defaultMessage: "Error",
                }),
                error.message ||
                  intl.formatMessage({
                    id: "generics.error.message",
                    defaultMessage: "Something went wrong",
                  })
              );
            } finally {
              setIsLoading(false);
            }
          }}
          appearance="outline"
          color="brand"
          icon={<Add24Regular />}
          iconPosition="before"
          disabled={isLoading}
          className={styles.button}
          style={{
            borderColor: "#A6A6A6",
            color: "#A6A6A6",
          }}
        >
          <FormattedMessage id="app.button.createSFFTables" />
        </Button>
        <Button
          content={intl.formatMessage({ id: "app.button.syncCodeLists" })}
          onClick={async () => {
            setIsLoading(true);
            try {
              await populateCodeLists();
              dialog.showDialog(
                intl.formatMessage({
                  id: "generics.success",
                  defaultMessage: "Success",
                }),
                intl.formatMessage({
                  id: "syncCodeLists.messages.success",
                  defaultMessage: "Code lists synchronized successfully",
                })
              );
            } catch (error: any) {
              dialog.showDialog(
                intl.formatMessage({
                  id: "generics.error",
                  defaultMessage: "Error",
                }),
                intl
                  .formatMessage(
                    {
                      id: "createTables.messages.error.populateCodeList",
                      defaultMessage: "Error populating code list: {tableName}",
                    },
                    {
                      tableName: error.message,
                    }
                  )
                  ?.toString() || intl.formatMessage({ id: "generics.error.message" })
              );
            } finally {
              setIsLoading(false);
            }
          }}
          appearance="outline"
          color="brand"
          icon={<ArrowSync24Regular />}
          iconPosition="before"
          disabled={isLoading}
          className={styles.button}
          style={{
            borderColor: "#1B4B9D",
            color: "#1B4B9D",
          }}
        >
          <FormattedMessage id="app.button.syncCodeLists" />
        </Button>
        <Button
          content={intl.formatMessage({ id: "app.button.importSeliGLI" })}
          onClick={async () => {
            setIsLoading(true);
            try {
              await populateSeliGLI();
              dialog.showDialog(
                intl.formatMessage({
                  id: "generics.success",
                  defaultMessage: "Success",
                }),
                intl.formatMessage({
                  id: "app.button.importSeliGLI.success",
                  defaultMessage:
                    "SELI-GLI Themes, Outcomes, and Indicators imported successfully!",
                })
              );
            } catch (error: any) {
              dialog.showDialog(
                intl.formatMessage({
                  id: "generics.error",
                  defaultMessage: "Error",
                }),
                error.message ||
                  intl.formatMessage({
                    id: "generics.error.message",
                    defaultMessage: "Something went wrong",
                  })
              );
            } finally {
              setIsLoading(false);
            }
          }}
          appearance="outline"
          color="brand"
          icon={<ArrowSync24Regular />}
          iconPosition="before"
          disabled={isLoading}
          className={styles.button}
          style={{
            borderColor: "#1B4B9D",
            color: "#1B4B9D",
          }}
        >
          <FormattedMessage
            id="app.button.importSeliGLI"
            defaultMessage="Import SELI-GLI"
          />
        </Button>
        <Button
          content={intl.formatMessage({ id: "app.button.userGuide" })}
          onClick={() => {
            window.open(
              "https://www.commonapproach.org/wp-content/uploads/2025/05/Guide-for-Excel-Add-In-Basic-Tier-V3.0-and-SFF.pdf"
            );
          }}
          appearance="outline"
          icon={<Book24Regular />}
          iconPosition="before"
          disabled={isLoading}
          className={styles.button}
          style={{
            borderColor: "#FF8B3C",
            color: "#FF8B3C",
          }}
        >
          <FormattedMessage
            id="app.button.userGuide"
            defaultMessage="User Guide"
          />
        </Button>
      </div>

      <p className={styles.warning_note}>
        <FormattedMessage
          id="app.taskpane.warning"
          defaultMessage="This task pane must be open for the add-in to work!"
        />
      </p>

      <p className={styles.note_message}>
        <FormattedMessage
          id="app.getSampleData"
          defaultMessage="New user? Try importing a"
        />{" "}
        <span
          aria-label="sample data file"
          className={styles.link}
          onClick={downloadSampleData}
          role="button"
          tabIndex={0}
        >
          <FormattedMessage
            id="app.link.sampleData"
            defaultMessage="Basic Tier sample data file"
          />
        </span>{" "}
        <FormattedMessage
          id="generics.or"
          defaultMessage="or"
        />{" "}
        <span
          aria-label="sample data file + sff module"
          className={styles.link}
          onClick={downloadSampleDataSFF}
          role="button"
          tabIndex={0}
        >
          <FormattedMessage
            id="app.link.sampleDataSFF"
            defaultMessage="Basic Tier + SFF sample data file"
          />
        </span>
      </p>
    </div>
  );
};

export default App;
