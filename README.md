# Medical Appointment Excel Matcher

A web-based Python application that matches medical appointment data from one Excel file with remarks and agent names from another Excel file, creating a combined output with all relevant information.

## Features

- 🏥 **Excel File Matching**: Matches Patient IDs between two Excel files
- 📊 **Excel Generation**: Creates formatted Excel files with matched appointment data
- 🔄 **Multiple Match Support**: Handles cases where Patient ID appears multiple times (creates separate rows)
- 📋 **Multi-Sheet Support**: Automatically finds the correct sheet in multi-sheet Excel files
- 🏷️ **Agent Name Matching**: Includes agent names from remarks Excel
- 🌐 **Web Interface**: User-friendly web interface with drag-and-drop support
- 📱 **Responsive Design**: Works on desktop and mobile devices
- ⚡ **Fast Processing**: Quick processing with real-time feedback
- 🎨 **Modern UI**: Clean, professional interface with smooth animations
- 🔍 **Smart Column Detection**: Automatically finds columns even if headers are in different rows

## Installation

1. **Clone or download this project**

   ```bash
   cd excel-remark
   ```

2. **Install Python dependencies**

   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**

   ```bash
   python app.py
   ```

4. **Open your browser**
   Navigate to `http://localhost:8080`

## Usage

1. **Upload Appointments Excel**: Click the upload area or drag and drop an Excel file containing appointment data
2. **Upload Remarks Excel**: Click the upload area or drag and drop an Excel file containing Patient ID, Remark, and Agent Name columns
3. **Process**: Click "Process Excel Files" button
4. **Review**: View the matched results in the web interface
5. **Download**: Download the combined Excel file with all matched data

## Input Format

### Appointments Excel File

Must contain:

- **Pat ID** or **Patient ID** column (required)
- Any other appointment data columns (Office, Appt Date, Time, Pat Name, etc.)

### Remarks Excel File

Must contain:

- **Patient ID** column (matches with Pat ID from appointments)
- **Remark** column (status/notes for each patient)
- **Agent Name** column (optional but recommended)

## Output Format

The generated Excel file contains:

- All original columns from the appointments Excel
- **Pat ID**: Patient identification number
- **Remark**: Matched remarks from the remarks Excel
- **Agent Name**: Matched agent names from the remarks Excel

## Multiple Matches Handling

If a Patient ID appears multiple times in the remarks Excel:

- The app creates **separate rows** for each match
- Each row contains the original appointment data plus the specific remark/agent name from that match

## Requirements

- Python 3.7+
- Flask 2.3.3
- openpyxl 3.1.2
- Werkzeug 2.3.7
- Jinja2 3.1.2

## File Structure

```
excel-remark/
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── templates/
│   ├── index.html        # Main upload interface
│   └── results.html      # Results display page
├── railway.json          # Railway deployment configuration
├── Procfile              # Process file for Railway
└── README.md            # This file
```

## Technical Details

- **Excel Processing**: Uses openpyxl for reliable Excel file reading and writing
- **Multi-Sheet Support**: Automatically detects and uses the correct sheet from multi-sheet workbooks
- **Flexible Column Detection**: Handles headers in rows 1-5, not just row 1
- **Column Name Matching**: Flexible matching for column names (handles spaces, underscores, hyphens, dots)
- **Excel Generation**: Uses openpyxl for creating formatted Excel files with auto-adjusted column widths
- **Web Framework**: Flask for lightweight web interface
- **File Handling**: Secure file upload with validation
- **Error Handling**: Comprehensive error handling and user feedback

## Security Notes

- Only Excel files (.xlsx, .xls) are accepted for upload
- Files are processed in memory (no temporary file storage)
- File names are sanitized to prevent security issues

## Deployment

This application is deployable to Railway. See `railway.json` for configuration.

### Railway Deployment Steps:

1. Connect your GitHub repository to Railway
2. Set environment variables:
   - `SECRET_KEY`: A secure secret key for Flask sessions
   - `FLASK_DEBUG`: Set to `False` for production
3. Deploy: Railway will automatically detect and deploy the application

## Troubleshooting

**Common Issues:**

1. **"Pat ID column not found"**: Ensure your appointments Excel has a column named "Pat ID" or "Patient ID"
2. **"Patient ID column not found"**: Ensure your remarks Excel has a "Patient ID" column
3. **"Remark column not found"**: Ensure your remarks Excel has a "Remark" column
4. **"Invalid file type"**: Make sure you're uploading Excel files (.xlsx or .xls)
5. **Port already in use**: Change the port using `PORT` environment variable or modify `app.py`

## License

This project is open source and available under the MIT License.
