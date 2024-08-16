import { Button, makeStyles } from '@fluentui/react-components';
import {
  Add24Regular,
  ArrowCircleDown24Regular,
  ArrowCircleUp24Regular,
} from '@fluentui/react-icons';
import * as React from 'react';
import { useDialogContext } from '../context/DialogContext';
import { importData } from '../import/import';
import { createSheetsAndTables } from '../taskpane';
import ExportDialog from './ExportDialog';
import Header from './Header';

interface AppProps {}

const useStyles = makeStyles({
  root: {
    minHeight: '100vh',
    width: '100%',
  },
  buttons_group: {
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'center',
    alignItems: 'center',
    margin: '1rem',
    gap: '1rem',
  },
  button: {
    width: '160px',
  },
});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();
  const dialog = useDialogContext();
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [isImporting, setIsImporting] = React.useState(false);
  const [isExportDialogOpen, setIsExportDialogOpen] = React.useState(false);

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
    if (file && (file.name.endsWith('.jsonld') || file.name.endsWith('.json'))) {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const data = JSON.parse(e.target?.result as any);
          await onSuccess(data);
        } catch (error) {
          onError(new Error('File is not a valid JSON/JSON-LD file.'));
        }
      };
      reader.readAsText(file);
    } else {
      onError(new Error('File is not a JSON/JSON-LD file.'));
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
        <Button
          content='Import Data'
          onClick={handleImportData}
          appearance='outline'
          icon={<ArrowCircleUp24Regular />}
          iconPosition='before'
          disabled={isImporting}
          className={styles.button}
          style={{
            borderColor: 'rgb(60, 174, 163)',
            color: 'rgb(60, 174, 163)',
          }}
        >
          Import Data
        </Button>
        <Button
          content='Export Data'
          onClick={() => {
            setIsExportDialogOpen(true);
          }}
          appearance='outline'
          icon={<ArrowCircleDown24Regular />}
          iconPosition='before'
          disabled={isImporting}
          className={styles.button}
          style={{
            borderColor: 'rgb(80, 183, 224)',
            color: 'rgb(80, 183, 224)',
          }}
        >
          Export Data
        </Button>
        <Button
          content='Create Sheets and Tables'
          onClick={() => {
            createSheetsAndTables();
          }}
          appearance='outline'
          color='brand'
          icon={<Add24Regular />}
          iconPosition='before'
          disabled={isImporting}
          className={styles.button}
          style={{
            borderColor: 'rgb(45, 98, 215)',
            color: 'rgb(45, 98, 215)',
          }}
        >
          Create Tables
        </Button>
      </div>
      <input
        ref={fileInputRef}
        style={{ display: 'none' }}
        title='file'
        type='file'
        onChange={async (e) => {
          await handleFileChange(
            e,
            async (data) => {
              try {
                await importData(data, dialog.showDialog, setIsImporting);
              } catch (error: any) {
                setIsImporting(false);
                dialog.showDialog('Error!', error.message || 'Something went wrong');
              }
            },
            (error) => {
              dialog.showDialog('Error!', error.message);
            }
          );
          // clear the file input
          if (fileInputRef.current) {
            fileInputRef.current.value = '';
          }
        }}
      />
    </div>
  );
};

export default App;
