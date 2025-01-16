import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Input,
} from "@fluentui/react-components";
import React, { useState } from "react";
import { FormattedMessage, useIntl } from "react-intl";
import { useDialogContext } from "../context/DialogContext";
import { exportData } from "../export/export";

interface ExportDialogProps {
  isDialogOpen: boolean;
  setDialogOpen: (isOpen: boolean) => void;
  setIslLoading: (isLoading: boolean) => void;
}

const ExportDialog: React.FC<ExportDialogProps> = ({
  isDialogOpen,
  setDialogOpen,
  setIslLoading,
}) => {
  const { showDialog } = useDialogContext();
  const [inputValue, setInputValue] = useState("");
  const intl = useIntl();

  const handleExport = async () => {
    // Set the loading state
    setIslLoading(true);

    // Close the dialog
    setDialogOpen(false);

    // Clean the input value to make it compatible with all file systems
    const cleanedOrgName = inputValue.replace(/[^\w]/gi, "");

    // Check if the input value is empty
    if (!cleanedOrgName) {
      showDialog(
        `${intl.formatMessage({ id: "generics.error" })}!`,
        intl.formatMessage({ id: "export.messages.enterOrganization" })
      );
      return;
    }

    // Set the cleaned org name using the provided hook
    try {
      showDialog(
        intl.formatMessage({ id: "export.messages.exporting" }),
        intl.formatMessage({ id: "export.messages.waitExport" })
      );
      await exportData(intl, cleanedOrgName, showDialog);
    } catch (error: any) {
      showDialog(
        `${intl.formatMessage({ id: "generics.error" })}!`,
        error.message ||
          intl.formatMessage({
            id: "generics.error.message",
            defaultMessage: "Something went wrong",
          })
      );
    }

    // Reset the loading state and the
    setInputValue("");
    setIslLoading(false);
  };

  return (
    <>
      {isDialogOpen && (
        <Dialog
          open={isDialogOpen}
          onOpenChange={(_, data) => {
            if (!data.open) {
              setDialogOpen(false);
            }
          }}
        >
          <DialogSurface>
            <DialogBody>
              <DialogTitle>
                <FormattedMessage
                  id="export.button.export"
                  defaultMessage="Export Data"
                />
              </DialogTitle>
              <DialogContent>
                <p>
                  <FormattedMessage
                    id="export.messages.enterOrganization"
                    defaultMessage="Enter the name of the organization you want to export:"
                  />
                </p>
                <Input
                  value={inputValue}
                  onChange={(e) => setInputValue(e.target.value)}
                  placeholder={intl.formatMessage({
                    id: "export.placeholder.organization",
                    defaultMessage: "Organization name",
                  })}
                />
              </DialogContent>

              <DialogActions
                style={{
                  display: "flex",
                  flexDirection: "row",
                  justifyContent: "flex-end",
                  gap: "1rem",
                }}
              >
                <Button
                  appearance="secondary"
                  onClick={() => {
                    setDialogOpen(false);
                  }}
                >
                  <FormattedMessage
                    id="generics.button.close"
                    defaultMessage="Cancel"
                  />
                </Button>
                <Button
                  appearance="primary"
                  onClick={handleExport}
                >
                  <FormattedMessage
                    id="export.button.export"
                    defaultMessage="Export"
                  />
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      )}
    </>
  );
};

export default ExportDialog;
