import { Image, makeStyles, tokens } from '@fluentui/react-components';
import * as React from 'react';

export interface HeaderProps {
  title: string;
  logo: string;
  message: string;
}

const useStyles = makeStyles({
  welcome__header: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    paddingBottom: '30px',
    paddingTop: '50px',
    backgroundColor: tokens.colorNeutralBackground3,
  },
  message: {
    fontSize: tokens.fontSizeHero700,
    fontWeight: tokens.fontWeightRegular,
    fontColor: tokens.colorNeutralBackgroundStatic,
    textAlign: 'center',
    lineHeight: '1.5',
  },
});

const Header: React.FC<HeaderProps> = (props: HeaderProps) => {
  const { title, logo, message } = props;
  const styles = useStyles();

  return (
    <section className={styles.welcome__header}>
      <Image
        width='250'
        src={logo}
        alt={title}
      />
      <h1 className={styles.message}>{message}</h1>
    </section>
  );
};

export default Header;
