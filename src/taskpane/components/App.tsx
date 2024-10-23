import { Button, makeStyles } from "@fluentui/react-components";
import {
  Add24Regular,
  ArrowCircleDown24Regular,
  ArrowCircleUp24Regular,
} from "@fluentui/react-icons";
import { ArrowSync24Regular } from "@fluentui/react-icons/fonts";
import * as React from "react";
import { FormattedMessage, useIntl } from "react-intl";
import { useDialogContext } from "../context/DialogContext";
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
  },
  buttons_group: {
    display: "flex",
    flexDirection: "column",
    justifyContent: "center",
    alignItems: "center",
    margin: "1rem",
    gap: "1rem",
  },
  button: {
    width: "160px",
  },
});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();
  const dialog = useDialogContext();
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [isImporting, setIsImporting] = React.useState(false);
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

  return (
    <div className={styles.root}>
      <ExportDialog
        isDialogOpen={isExportDialogOpen}
        setDialogOpen={setIsExportDialogOpen}
      />
      <Header />
      <div className={styles.buttons_group}>
        <input
          ref={fileInputRef}
          style={{ display: "none" }}
          title="file"
          type="file"
          onChange={async (e) => {
            await handleFileChange(
              e,
              async (data) => {
                try {
                  await importData(intl, data, dialog.showDialog, setIsImporting);
                } catch (error: any) {
                  setIsImporting(false);
                  dialog.showDialog(
                    `${intl.formatMessage({ id: "generics.error" })}!`,
                    error.message || intl.formatMessage({ id: "generics.error.message" })
                  );
                }
              },
              (error) => {
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
          disabled={isImporting}
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
          disabled={isImporting}
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
            try {
              await createSheetsAndTables();
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
            }
          }}
          appearance="outline"
          color="brand"
          icon={<Add24Regular />}
          iconPosition="before"
          disabled={isImporting}
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
            try {
              await createSFFModuleSheetsAndTables();
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
            }
          }}
          appearance="outline"
          color="brand"
          icon={<Add24Regular />}
          iconPosition="before"
          disabled={isImporting}
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
            }
          }}
          appearance="outline"
          color="brand"
          icon={<ArrowSync24Regular />}
          iconPosition="before"
          disabled={isImporting}
          className={styles.button}
          style={{
            borderColor: "#1B4B9D",
            color: "#1B4B9D",
          }}
        >
          <FormattedMessage id="app.button.syncCodeLists" />
        </Button>
      </div>
    </div>
  );
};

export default App;
