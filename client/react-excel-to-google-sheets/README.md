# React Excel to Google Sheets

This project is a web application that allows users to upload Excel files and automatically process them to summarize attendance data, which is then uploaded to Google Sheets. The application is built using Flask for the backend and React for the frontend.

## Project Structure

```
react-excel-to-google-sheets
├── server
│   ├── app.py                # Main entry point for the Flask server
│   ├── handleExcel.py        # Functions for processing Excel files
│   ├── uploadGgSheet.py      # Functions for Google Sheets integration
│   ├── network_checker.py     # Functions to check network status
│   ├── requirements.txt       # Python dependencies for the server
│   └── .env                   # Environment variables (e.g., Google credentials)
├── client
│   ├── public
│   │   └── index.html         # Main HTML file for the React application
│   ├── src
│   │   ├── App.js             # Main React component
│   │   ├── index.js           # Entry point for the React application
│   │   └── components
│   │       └── YourComponent.js # Additional React component
│   ├── package.json           # Configuration file for the React application
│   └── README.md              # Documentation for the client-side application
├── README.md                  # Overview and documentation for the entire project
└── Dockerfile                 # Instructions for building a Docker image
```

## Features

- Upload Excel files (.xlsx) for processing.
- Validate and summarize attendance data from the uploaded files.
- Upload processed data to Google Sheets.
- Check network connectivity and internet speed before processing files.

## Getting Started

### Prerequisites

- Python 3.x
- Node.js and npm
- Google Cloud account with Sheets API enabled

### Installation

1. Clone the repository:
   ```
   git clone <repository-url>
   cd react-excel-to-google-sheets
   ```

2. Set up the server:
   - Navigate to the `server` directory.
   - Install the required Python packages:
     ```
     pip install -r requirements.txt
     ```

3. Set up the client:
   - Navigate to the `client` directory.
   - Install the required Node.js packages:
     ```
     npm install
     ```

4. Configure environment variables:
   - Create a `.env` file in the `server` directory and set the `GOOGLE_CREDENTIALS_PATH` variable to the path of your Google credentials JSON file.

### Running the Application

1. Start the Flask server:
   ```
   cd server
   python app.py
   ```

2. Start the React application:
   ```
   cd client
   npm start
   ```

3. Open your web browser and navigate to `http://localhost:3000` to access the application.

## Deployment

To deploy the application using Docker, build the Docker image with the following command:
```
docker build -t react-excel-to-google-sheets .
```

Run the Docker container:
```
docker run -p 5000:5000 react-excel-to-google-sheets
```

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

## License

This project is licensed under the MIT License. See the LICENSE file for details.