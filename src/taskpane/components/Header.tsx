import { Image, makeStyles, tokens } from "@fluentui/react-components";
import * as React from "react";
import { FormattedMessage } from "react-intl";

export interface HeaderProps {}

const useStyles = makeStyles({
  welcome__header: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    paddingTop: "15px", // Reduced padding from top
    paddingBottom: "10px", // Reduced padding from bottom
    backgroundColor: tokens.colorNeutralBackground1,
  },
  message: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightRegular,
    fontColor: tokens.colorNeutralBackgroundStatic,
    textAlign: "center",
    marginBottom: "5px", // Reduced margin
  },
  warning_note: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightBold, // Changed to bold
    fontColor: tokens.colorNeutralBackgroundStatic,
    textAlign: "center",
    marginTop: "5px",
    marginBottom: "5px", // Reduced margin
  },
  note_message: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightRegular,
    fontColor: tokens.colorNeutralBackgroundStatic,
    textAlign: "center",
    marginTop: "5px",
    marginBottom: "5px", // Reduced margin
  },
  link: {
    cursor: "pointer",
    color: tokens.colorBrandForeground1,
    textDecoration: "underline",
  },
});

const Header: React.FC<HeaderProps> = () => {
  const styles = useStyles();

  return (
    <section className={styles.welcome__header}>
      <Image
        width="100"
        src="assets/logo.png"
        alt="Common Impact Data Standard Add-in"
      />
      <p className={styles.message}>
        <FormattedMessage
          id="app.description"
          defaultMessage="Common Impact Data Standard Version 3.0"
        />
        {" - "}
        <FormattedMessage
          id="app.standardTier"
          defaultMessage="Basic Tier"
        />
      </p>
    </section>
  );
};

export default Header;
