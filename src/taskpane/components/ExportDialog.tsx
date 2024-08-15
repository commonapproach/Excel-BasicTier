import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Input,
} from '@fluentui/react-components';
import React, { useState } from 'react';
import { useDialogContext } from '../context/DialogContext';
import { exportData } from '../export/export';

interface ExportDialogProps {
  isDialogOpen: boolean;
  setDialogOpen: (isOpen: boolean) => void;
}

const ExportDialog: React.FC<ExportDialogProps> = ({ isDialogOpen, setDialogOpen }) => {
  const { showDialog } = useDialogContext();
  const [inputValue, setInputValue] = useState('');

  const handleExport = async () => {
    // Clean the input value to make it compatible with all file systems
    const cleanedOrgName = inputValue.replace(/[^\w]/gi, '');

    // Check if the input value is empty
    if (!cleanedOrgName) {
      showDialog('Error!', 'Please enter the name of the organization you want to export');
      return;
    }

    // Set the cleaned org name using the provided hook
    try {
      showDialog('Exporting...', 'Please wait while we export the data');
      await exportData(cleanedOrgName, showDialog);
    } catch (error: any) {
      showDialog('Error!', error.message || 'Something went wrong');
    }

    // Close the dialog
    setDialogOpen(false);
    setInputValue('');
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
              <DialogTitle>Export</DialogTitle>
              <DialogContent>
                <p>Enter the name of the organization you want to export:</p>
                <Input
                  value={inputValue}
                  onChange={(e) => setInputValue(e.target.value)}
                  placeholder={'Enter organization name'}
                />
              </DialogContent>

              <DialogActions
                style={{
                  display: 'flex',
                  flexDirection: 'row',
                  justifyContent: 'flex-end',
                  gap: '1rem',
                }}
              >
                <Button
                  appearance='secondary'
                  onClick={() => {
                    setDialogOpen(false);
                  }}
                >
                  Close
                </Button>
                <Button
                  appearance='primary'
                  onClick={handleExport}
                >
                  Export
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
