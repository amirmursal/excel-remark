# Medical Appointment PDF to Excel Converter

A web-based Python application that converts medical appointment PDF files to Excel format, extracting structured appointment data and organizing it into specific medical columns.

## Features

- 🏥 **Medical Data Extraction**: Extracts structured medical appointment data from PDFs
- 📊 **Excel Generation**: Creates formatted Excel files with medical appointment columns
- 🌐 **Web Interface**: User-friendly web interface with drag-and-drop support
- 📱 **Responsive Design**: Works on desktop and mobile devices
- ⚡ **Fast Processing**: Quick conversion with progress indicators
- 🎨 **Modern UI**: Clean, professional interface with smooth animations
- 🔍 **Smart Parsing**: Automatically identifies and extracts medical appointment fields
- 👥 **Multi-Patient Support**: Captures multiple patients per office/location
- 📈 **Processing Status**: Shows appointment count and processing feedback

## Installation

1. **Clone or download this project**

   ```bash
   cd pdf-excel-comparison
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

1. **Upload PDF**: Click the upload area or drag and drop a PDF file
2. **Convert**: Click "Convert to Excel" button
3. **Download**: The converted Excel file will be automatically downloaded

## Output Format

The generated Excel file contains the following medical appointment columns:

- **Office**: Medical office or clinic name
- **Appt Date**: Appointment date
- **Time**: Appointment time
- **Pat ID**: Patient identification number
- **Pat Name**: Patient name
- **Group No.**: Group or policy number
- **Insurance Note**: Insurance information

## Requirements

- Python 3.7+
- Flask 2.3.3
- PyPDF2 3.0.1
- openpyxl 3.1.2

## File Structure

```
pdf-excel-comparison/
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── templates/
│   └── index.html        # Web interface template
├── uploads/              # Temporary upload directory (auto-created)
└── README.md            # This file
```

## Technical Details

- **PDF Processing**: Uses PyPDF2 library for reliable text extraction
- **Excel Generation**: Uses openpyxl for creating formatted Excel files
- **Web Framework**: Flask for lightweight web interface
- **File Handling**: Secure file upload with validation
- **Error Handling**: Comprehensive error handling and user feedback

## Security Notes

- Only PDF files are accepted for upload
- Uploaded files are automatically deleted after processing
- File names are sanitized to prevent security issues
- Temporary files are stored in a dedicated uploads directory

## Troubleshooting

**Common Issues:**

1. **"No text content found"**: The PDF might be image-based or password-protected
2. **"Invalid file type"**: Make sure you're uploading a PDF file
3. **Port already in use**: Change the port in `app.py` if port 5000 is occupied

**For image-based PDFs:**
This tool extracts text only. For PDFs with scanned images, you would need OCR (Optical Character Recognition) capabilities.

## License

This project is open source and available under the MIT License.
