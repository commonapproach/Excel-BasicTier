import { Button, makeStyles } from '@fluentui/react-components';
import {
  ArrowExport16Filled,
  ArrowImport16Filled,
  DesignIdeas24Regular,
} from '@fluentui/react-icons';
import * as React from 'react';
import { useDialogContext } from '../context/DialogContext';
import { importData } from '../import/import';
import { createSheetsAndTables } from '../taskpane';
import Header from './Header';

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: '100vh',
    width: '100%',
  },
});

const App: React.FC<AppProps> = ({ title }) => {
  const styles = useStyles();
  const dialog = useDialogContext();
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [isImporting, setIsImporting] = React.useState(false);

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
      <Header
        logo='assets/logo.png'
        title={title}
        message='Common Impact Data Standard'
      />
      <div
        style={{
          display: 'flex',
          flexDirection: 'column',
          justifyContent: 'center',
          alignItems: 'center',
          margin: '1rem',
          gap: '1rem',
        }}
      >
        <Button
          content='Create Sheets and Tables'
          onClick={() => {
            createSheetsAndTables();
          }}
          appearance='primary'
          icon={<DesignIdeas24Regular />}
          iconPosition='before'
        >
          Create Sheets and Tables
        </Button>
        <Button
          content='Import Data'
          onClick={handleImportData}
          appearance='primary'
          icon={<ArrowImport16Filled />}
          iconPosition='before'
          disabled={isImporting}
        >
          Import Data
        </Button>
        <Button
          content='Export Data'
          onClick={() => {
            dialog.showDialog('Exporting Data', 'Exporting data...', () => {});
          }}
          appearance='primary'
          icon={<ArrowExport16Filled />}
          iconPosition='before'
        >
          Export Data
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
