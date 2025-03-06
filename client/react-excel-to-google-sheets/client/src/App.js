import React, { useState } from 'react';
import './App.css';
import YourComponent from './components/YourComponent';

function App() {
    const [files, setFiles] = useState([]);
    const [googleSheetUrl, setGoogleSheetUrl] = useState('');
    const [sheetName, setSheetName] = useState('');
    const [facultyName, setFacultyName] = useState('');
    const [statusMessage, setStatusMessage] = useState('');

    const handleFileChange = (event) => {
        setFiles(event.target.files);
    };

    const handleSubmit = async (event) => {
        event.preventDefault();
        const formData = new FormData();
        for (let i = 0; i < files.length; i++) {
            formData.append('files', files[i]);
        }
        formData.append('googleSheetUrl', googleSheetUrl);
        formData.append('sheetName', sheetName);
        formData.append('faculty', facultyName);

        try {
            const response = await fetch('http://localhost:5000/process', {
                method: 'POST',
                body: formData,
            });
            const data = await response.json();
            setStatusMessage(data.status);
        } catch (error) {
            setStatusMessage('Error processing files: ' + error.message);
        }
    };

    return (
        <div className="App">
            <h1>Excel to Google Sheets Uploader</h1>
            <form onSubmit={handleSubmit}>
                <input type="file" multiple onChange={handleFileChange} />
                <input
                    type="text"
                    placeholder="Google Sheet URL"
                    value={googleSheetUrl}
                    onChange={(e) => setGoogleSheetUrl(e.target.value)}
                    required
                />
                <input
                    type="text"
                    placeholder="Sheet Name"
                    value={sheetName}
                    onChange={(e) => setSheetName(e.target.value)}
                    required
                />
                <input
                    type="text"
                    placeholder="Faculty Name"
                    value={facultyName}
                    onChange={(e) => setFacultyName(e.target.value)}
                    required
                />
                <button type="submit">Upload</button>
            </form>
            {statusMessage && <p>{statusMessage}</p>}
            <YourComponent />
        </div>
    );
}

export default App;