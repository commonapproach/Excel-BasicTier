# Common Approach Excel Add-in

This is a sample Excel Add-in that allows users to import, export, and manage their data according to the Common Approach to Impact Measurement Data Standard. It's intended to streamline data management for social purpose organizations, enhancing their ability to measure and report impact.

## Features

- **Import Functionality**: Allows users to import data in JSON-LD format, conforming to the Common Impact Data Standard.
- **Export Functionality**: Enables users to export their Excel data as JSON-LD files, adhering to the specified data standard.

## Prerequisites

The add-in is built using the [Office JavaScript API](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins) and [React](https://reactjs.org/).

- Node.js (version 20)
- Office 365 account
- Excel installed on your machine

## Local Installation and Testing

### Clone the Repository:

```bash
git clone https://github.com/commonapproach/Excel-BasicTier.git
```

### Navigate to Project Directory:

    ```bash
    cd Excel-BasicTier
    ```

### Install Dependencies:

```bash
npm install
```

### Set Up Excel Add-in:

1. Go to the [Office Add-ins](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins) documentation and follow the instructions to sideload your add-in in Excel.
2. Update the `manifest.xml` file with the correct URLs pointing to your local or production server.

### Run the Add-in Locally:

1. Start the local server:

```bash
npm start
```

### Testing the Add-in in Excel:

1. Once your server is running, Excel will open automatically.
2. In the new workbook you will see in the Home tab a new group withe the Common Approach logo.
3. Click on the logo to open the task pane and start testing.

### Contributing

Contributions to this project are welcome. Please ensure that your code adheres to the project's standards and guidelines.
