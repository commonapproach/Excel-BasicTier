import { Image, makeStyles, tokens } from '@fluentui/react-components';
import * as React from 'react';

export interface HeaderProps {}

const useStyles = makeStyles({
  welcome__header: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    paddingTop: '30px',
    paddingBottom: '20px',
    backgroundColor: tokens.colorNeutralBackground1,
  },
  message: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightRegular,
    fontColor: tokens.colorNeutralBackgroundStatic,
    textAlign: 'center',
    marginBottom: '10px',
  },
  tier_level: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightBold,
    fontColor: tokens.colorNeutralBackgroundStatic,
    textAlign: 'center',
    marginTop: '0px',
    marginBottom: '10px',
  },
  note_message: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightBold,
    fontColor: tokens.colorNeutralBackgroundStatic,
    textAlign: 'center',
    marginTop: '0px',
    marginBottom: '10px',
  },
  warning_note: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightRegular,
    fontColor: tokens.colorNeutralBackgroundStatic,
    textAlign: 'center',
    marginTop: '0px',
    marginBottom: '10px',
  },
});

const Header: React.FC<HeaderProps> = () => {
  const styles = useStyles();

  return (
    <section className={styles.welcome__header}>
      <Image
        width='200'
        src='assets/logo.png'
        alt='Common Impact Data Standard Add-in'
      />
      <p className={styles.message}>Compliant with Common Impact Data Standard Version 2.1</p>
      <p className={styles.tier_level}>Basic Tier</p>
      <p className={styles.note_message}>
        New user? Try importing this &nbsp;
        <a
          href='https://ontology.commonapproach.org/examples/CIDSBasicZerokitsTestData-SHARED.json'
          rel='noreferrer'
          download='CIDSBasicZerokitsTestData-SHARED.json'
        >
          sample data file
        </a>
      </p>
      <p className={styles.warning_note}>
        Please note that this task pane must be open for the add-in to work!
      </p>
    </section>
  );
};

export default Header;
