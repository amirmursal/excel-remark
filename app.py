from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import io
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')

# Configuration
ALLOWED_EXCEL_EXTENSIONS = {'xlsx', 'xls'}

# Global variables for storing processed data
processed_appointments = []
pdf_filename = ""
excel_data = {}

def allowed_excel_file(filename):
    """Check if the uploaded file has an allowed Excel extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXCEL_EXTENSIONS

def process_excel_file(file_stream):
    """Process uploaded Excel file and extract Patient ID and Remark data."""
    try:
        from openpyxl import load_workbook
        
        wb = load_workbook(file_stream)
        ws = wb.active
        
        # Find the Patient ID and Remark columns
        patient_id_col = None
        remark_col = None
        
        # Look for headers in the first row
        for col in range(1, ws.max_column + 1):
            cell_value = str(ws.cell(row=1, column=col).value or '').strip().lower()
            if 'patient' in cell_value and 'id' in cell_value:
                patient_id_col = col
            elif 'remark' in cell_value:
                remark_col = col
        
        if not patient_id_col:
            raise Exception("Patient ID column not found in Excel file")
        
        if not remark_col:
            raise Exception("Remark column not found in Excel file")
        
        # Extract data
        excel_data = {}
        for row in range(2, ws.max_row + 1):  # Skip header row
            patient_id = str(ws.cell(row=row, column=patient_id_col).value or '').strip()
            remark = str(ws.cell(row=row, column=remark_col).value or '').strip()
            
            # Clean up Patient ID - remove .0 if it's a float
            if patient_id.endswith('.0'):
                patient_id = patient_id[:-2]
            
            if patient_id:  # Only add non-empty patient IDs
                excel_data[patient_id] = remark
        
        return excel_data
        
    except Exception as e:
        raise Exception(f"Error processing Excel file: {str(e)}")

def process_appointments_excel(file_stream):
    """Read an appointments Excel and return list of appointment dicts with all columns.
    
    Only requires 'Pat ID' column. All other columns are preserved as-is.
    """
    from openpyxl import load_workbook

    wb = load_workbook(file_stream)
    ws = wb.active

    # Get all headers from first row
    headers = []
    for col in range(1, ws.max_column + 1):
        raw = ws.cell(row=1, column=col).value
        name = (str(raw or '')).strip()
        headers.append(name)

    # Find Pat ID column (case-insensitive)
    pat_id_col = None
    for i, header in enumerate(headers):
        if 'pat' in header.lower() and 'id' in header.lower():
            pat_id_col = i + 1  # 1-based column index
            break
    
    if pat_id_col is None:
        raise Exception("Pat ID column not found in appointments Excel. Please ensure there's a column containing 'Pat ID' or 'Patient ID'")

    # Read all rows
    appointments = []
    for row in range(2, ws.max_row + 1):
        record = {}
        
        # Read all columns
        for col, header in enumerate(headers, 1):
            value = ws.cell(row=row, column=col).value
            record[header] = '' if value is None else str(value)
        
        # Normalize Patient ID to string without trailing .0
        pat_id_value = record.get(headers[pat_id_col - 1], '')
        pid = str(pat_id_value).strip()
        if pid.endswith('.0'):
            pid = pid[:-2]
        record['Pat ID'] = pid  # Standardize the key name
        
        # Ensure Remark exists
        if 'Remark' not in record:
            record['Remark'] = ''
        
        # Skip empty rows (no Pat ID)
        if pid:
            appointments.append(record)

    return appointments

def update_appointments_with_remarks(appointments, excel_data):
    """Update appointments with remarks from Excel data based on Patient ID matching."""
    updated_count = 0
    
    for appointment in appointments:
        patient_id = str(appointment.get('Pat ID', '')).strip()
        
        # Try exact match first
        if patient_id and patient_id in excel_data:
            appointment['Remark'] = excel_data[patient_id]
            updated_count += 1
        # Try with .0 suffix (in case Excel has float format)
        elif patient_id and f"{patient_id}.0" in excel_data:
            appointment['Remark'] = excel_data[f"{patient_id}.0"]
            updated_count += 1
    
    return updated_count

def create_excel_from_pdf_data(appointments, filename):
    """Create Excel file from processed appointment data with all columns."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Appointment Data"
    
    if not appointments:
        return wb
    
    # Get all unique headers from all appointments
    all_headers = set()
    for appointment in appointments:
        all_headers.update(appointment.keys())
    
    # Convert to list and ensure Pat ID and Remark are at the end for visibility
    headers = list(all_headers)
    if 'Pat ID' in headers:
        headers.remove('Pat ID')
    if 'Remark' in headers:
        headers.remove('Remark')
    headers.extend(['Pat ID', 'Remark'])  # Put these at the end
    
    # Set headers
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    # Add appointment data
    for row, appointment in enumerate(appointments, 2):
        for col, header in enumerate(headers, 1):
            value = appointment.get(header, '')
            ws.cell(row=row, column=col, value=value)
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        
        adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Save to memory
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    
    return excel_buffer

@app.route('/')
def index():
    """Main page with file upload form."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file uploads - two Excel inputs: appointments and remarks."""
    global processed_appointments, pdf_filename, pdf_data, excel_data

    appointments_file = request.files.get('appointments_file')
    remarks_file = request.files.get('remarks_file')

    # Require both Excel files
    if not appointments_file or appointments_file.filename == '':
        flash('Please upload the Appointments Excel file.')
        return redirect(url_for('index'))

    if not remarks_file or remarks_file.filename == '':
        flash('Please upload the Remarks Excel file.')
        return redirect(url_for('index'))

    if not allowed_excel_file(appointments_file.filename):
        flash('Invalid Appointments Excel file type. Please upload .xlsx or .xls')
        return redirect(url_for('index'))

    if not allowed_excel_file(remarks_file.filename):
        flash('Invalid Remarks Excel file type. Please upload .xlsx or .xls')
        return redirect(url_for('index'))

    try:
        # Process appointments Excel directly from memory
        appointments_filename = secure_filename(appointments_file.filename)
        processed_appointments = process_appointments_excel(appointments_file)
        pdf_filename = os.path.splitext(appointments_filename)[0]

        flash(f'Successfully processed appointments with {len(processed_appointments)} rows.')
    except Exception as e:
        flash(f'Error processing Appointments Excel: {str(e)}')
        return redirect(url_for('index'))

    # Process remarks Excel
    try:
        excel_data = process_excel_file(remarks_file)
        updated_count = update_appointments_with_remarks(processed_appointments, excel_data)

        flash(f'Successfully updated {updated_count} appointments with remarks.')
    except Exception as e:
        flash(f'Error processing Remarks Excel: {str(e)}')
        return redirect(url_for('index'))

    return redirect(url_for('results'))


@app.route('/results')
def results():
    """Show extracted appointment data."""
    global processed_appointments, pdf_filename
    
    if not processed_appointments:
        flash('No data found. Please upload a PDF file first.')
        return redirect(url_for('index'))
    
    return render_template('results.html', 
                         appointments=processed_appointments, 
                         filename=pdf_filename,
                         total_appointments=len(processed_appointments))

@app.route('/download')
def download_excel():
    """Download the Excel file."""
    global processed_appointments, pdf_filename
    
    if not processed_appointments:
        flash('No data to download. Please upload a PDF first.')
        return redirect(url_for('index'))
    
    try:
        # Create Excel file
        excel_buffer = create_excel_from_pdf_data(processed_appointments, pdf_filename)
        
        return send_file(excel_buffer, 
                        as_attachment=True, 
                        download_name=f'{pdf_filename}_appointments.xlsx',
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
    except Exception as e:
        flash(f'Error creating Excel file: {str(e)}')
        return redirect(url_for('index'))

if __name__ == '__main__':
    debug_mode = os.environ.get('FLASK_DEBUG', 'False').lower() == 'true'
    port = int(os.environ.get('PORT', 8080))
    app.run(debug=debug_mode, host='0.0.0.0', port=port)
