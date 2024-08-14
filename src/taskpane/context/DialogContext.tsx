import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
} from '@fluentui/react-components';
import React, { createContext, FC, useContext, useState } from 'react';

// Define the dialog context type
type DialogContextType = {
  showDialog: (header: string, content: string, onClose?: Function) => void;
};

// Create the dialog context
const DialogContext = createContext<DialogContextType | undefined>(undefined);

// Dialog handler function
export var dialogHandler: Function | null = null;

// Custom hook to access the dialog context
export const useDialogContext = () => {
  const context = useContext(DialogContext);
  if (!context) {
    throw new Error('useDialogContext must be used within a DialogContextProvider');
  }
  return context;
};

interface DialogContextProviderProps {
  children: React.ReactNode;
}

// Dialog context provider component
const DialogContextProvider: FC<DialogContextProviderProps> = ({ children }) => {
  const [open, setOpen] = useState(false);
  const [dialogHeader, setDialogHeader] = useState('');
  const [dialogContent, setDialogContent] = useState('');
  const [onNext, setOnNext] = useState<Function | null>(null);

  function showDialog(header: string, content: string, handleNextCallback: Function | null = null) {
    setDialogHeader(header);
    setDialogContent(content);
    setOnNext(handleNextCallback ? () => handleNextCallback : null);
    setOpen(true);
  }

  dialogHandler = showDialog;

  return (
    <DialogContext.Provider
      value={{
        showDialog,
      }}
    >
      <Dialog
        open={open}
        onOpenChange={(_, data) => {
          if (!data.open) {
            setDialogHeader('');
            setDialogContent('');
            setOnNext(null);
          }
        }}
      >
        <DialogSurface>
          <DialogBody>
            <DialogTitle>{dialogHeader}</DialogTitle>
            <DialogContent>
              <p dangerouslySetInnerHTML={{ __html: dialogContent }} />
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
                  setOpen(false);
                }}
              >
                Close
              </Button>
              {onNext && (
                <Button
                  appearance='primary'
                  onClick={() => {
                    if (onNext) onNext();
                  }}
                >
                  Next
                </Button>
              )}
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
      {children}
    </DialogContext.Provider>
  );
};

export default DialogContextProvider;
