# React Excel to Google Sheets Client

This project is a React application that serves as a front-end interface for uploading Excel files and processing them to Google Sheets. It interacts with a Flask server that handles the backend logic.

## Features

- Upload Excel files (.xlsx) for processing.
- Check network status and internet speed.
- Summarize attendance data from Excel files.
- Upload processed data to Google Sheets.

## Getting Started

### Prerequisites

- Node.js and npm installed on your machine.
- Access to a Google account with Google Sheets.

### Installation

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/react-excel-to-google-sheets.git
   ```

2. Navigate to the client directory:
   ```
   cd react-excel-to-google-sheets/client
   ```

3. Install the dependencies:
   ```
   npm install
   ```

### Running the Application

To start the React application, run:
```
npm start
```
This will start the development server and open the application in your default web browser.

### Building for Production

To create a production build of the application, run:
```
npm run build
```
This will generate a `build` folder containing the optimized application.

## API Endpoints

The client interacts with the following API endpoints provided by the Flask server:

- `POST /check-network`: Checks the network status and internet speed.
- `POST /process`: Processes uploaded Excel files and uploads data to Google Sheets.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

## License

This project is licensed under the MIT License. See the LICENSE file for details.