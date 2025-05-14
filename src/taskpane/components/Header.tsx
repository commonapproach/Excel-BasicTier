/* global console, fetch, window, document */
import { Image, makeStyles, tokens } from "@fluentui/react-components";
import * as React from "react";
import { FormattedMessage } from "react-intl";

export interface HeaderProps {}

const useStyles = makeStyles({
  welcome__header: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    paddingTop: "30px",
    paddingBottom: "20px",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  message: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightRegular,
    fontColor: tokens.colorNeutralBackgroundStatic,
    textAlign: "center",
    marginBottom: "10px",
  },
  tier_level: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightBold,
    fontColor: tokens.colorNeutralBackgroundStatic,
    textAlign: "center",
    marginTop: "0px",
    marginBottom: "10px",
  },
  note_message: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightBold,
    fontColor: tokens.colorNeutralBackgroundStatic,
    textAlign: "center",
    marginTop: "0px",
    marginBottom: "10px",
  },
  warning_note: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightRegular,
    fontColor: tokens.colorNeutralBackgroundStatic,
    textAlign: "center",
    marginTop: "0px",
    marginBottom: "10px",
  },
  link: {
    cursor: "pointer",
    color: tokens.colorBrandForeground1,
    textDecoration: "underline",
  },
});

const Header: React.FC<HeaderProps> = () => {
  const styles = useStyles();

  const downloadSampleData = async (event: React.MouseEvent) => {
    event.preventDefault();
    try {
      const url =
        "https://ontology.commonapproach.org/examples/CIDSBasicZerokitsTestData-SHARED.json";
      const response = await fetch(url);
      const data = await response.blob();

      // Create a blob URL and trigger download
      const blobUrl = window.URL.createObjectURL(data);
      const a = document.createElement("a");
      a.style.display = "none";
      a.href = blobUrl;
      a.download = "CIDSBasicZerokitsTestData-SHARED.json";
      document.body.appendChild(a);
      a.click();

      // Clean up
      window.URL.revokeObjectURL(blobUrl);
      document.body.removeChild(a);
    } catch (error) {
      console.error("Error downloading sample data:", error);
    }
  };

  const downloadSampleDataSFF = async (event: React.MouseEvent) => {
    event.preventDefault();
    try {
      const url = "https://ontology.commonapproach.org/examples/CIDSBasictestandSFFSampleData.json";
      const response = await fetch(url);
      const data = await response.blob();

      // Create a blob URL and trigger download
      const blobUrl = window.URL.createObjectURL(data);
      const a = document.createElement("a");
      a.style.display = "none";
      a.href = blobUrl;
      a.download = "CIDSBasictestandSFFSampleData.json";
      document.body.appendChild(a);
      a.click();

      // Clean up
      window.URL.revokeObjectURL(blobUrl);
      document.body.removeChild(a);
    } catch (error) {
      console.error("Error downloading sample data:", error);
    }
  };

  return (
    <section className={styles.welcome__header}>
      <Image
        width="200"
        src="assets/logo.png"
        alt="Common Impact Data Standard Add-in"
      />
      <p className={styles.message}>
        <FormattedMessage
          id="app.description"
          defaultMessage="Compliant with Common Impact Data Standard Version 2.1"
        />
      </p>
      <p className={styles.tier_level}>
        <FormattedMessage
          id="app.standardTier"
          defaultMessage="Basic Tier"
        />
      </p>
      <p className={styles.note_message}>
        <FormattedMessage
          id="app.getSampleData"
          defaultMessage="New user? Try importing a"
        />{" "}
        &nbsp;
        <span
          aria-label="sample data file"
          className={styles.link}
          onClick={downloadSampleData}
          role="button"
          tabIndex={0}
        >
          <FormattedMessage
            id="app.link.sampleData"
            defaultMessage="sample data file"
          />
        </span>
        &nbsp;
        <FormattedMessage
          id="generics.or"
          defaultMessage="or"
        />
        &nbsp;
        <span
          aria-label="sample data file + sff module"
          className={styles.link}
          onClick={downloadSampleDataSFF}
          role="button"
          tabIndex={0}
        >
          <FormattedMessage
            id="app.link.sampleDataSFF"
            defaultMessage="sample data file + SFF module"
          />
        </span>
      </p>
      <p className={styles.warning_note}>
        <FormattedMessage
          id="app.taskpane.warning"
          defaultMessage="Please note that this task pane must be open for the add-in to work!"
        />
      </p>
    </section>
  );
};

export default Header;
