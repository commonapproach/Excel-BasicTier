import { Button, makeStyles } from '@fluentui/react-components';
import { DesignIdeas24Regular } from '@fluentui/react-icons';
import * as React from 'react';
import { createSheetsAndTables } from '../taskpane';
import Header from './Header';

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: '100vh',
  },
});

const App: React.FC<AppProps> = ({ title }) => {
  const styles = useStyles();
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
          justifyContent: 'center',
          margin: '1rem',
          width: '100%',
        }}
      >
        <Button
          content='Create Sheets and Tables'
          onClick={createSheetsAndTables}
          appearance='primary'
          icon={<DesignIdeas24Regular />}
          iconPosition='before'
        >
          Create Sheets and Tables
        </Button>
      </div>
    </div>
  );
};

export default App;
