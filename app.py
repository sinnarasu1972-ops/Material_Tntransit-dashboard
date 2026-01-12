from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from io import BytesIO
import uvicorn
import json
import numpy as np
import os

app = FastAPI(title="Unnati Motors Material In Transit Dashboard")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# For local development: D:\Material Intransit\Material_Intransit_All_Divisions.xlsx
# For Render: upload to /data folder or use environment variable
EXCEL_FILE = os.getenv('EXCEL_FILE', './Material_Intransit_All_Divisions.xlsx')

def clean_dataframe(df):
    for col in df.columns:
        if df[col].dtype in ['float64', 'float32']:
            df[col] = df[col].apply(lambda x: None if (pd.isna(x) or np.isnan(x) if isinstance(x, (int, float)) else False) else x)
        df[col] = df[col].fillna('')
    return df

@app.on_event("startup")
def load_data():
    global df
    try:
        if not os.path.exists(EXCEL_FILE):
            print(f"Excel file not found at {EXCEL_FILE}")
            print(f"Please ensure Material_Intransit_All_Divisions.xlsx is in the project root or set EXCEL_FILE environment variable")
            df = pd.DataFrame()
        else:
            df = pd.read_excel(EXCEL_FILE)
            df = clean_dataframe(df)
            print(f"Data loaded: {len(df)} rows")
    except Exception as e:
        print(f"Error loading data: {e}")
        df = pd.DataFrame()

@app.get("/api/filters")
def get_filters():
    try:
        if df.empty:
            return {
                "divisions": [],
                "age_buckets": [],
                "transporters": []
            }
        
        divisions = sorted([str(x) for x in df['Division'].dropna().unique().tolist() if str(x).strip()])
        age_buckets_raw = [str(x) for x in df['Age Bucket'].dropna().unique().tolist() if str(x).strip()]
        transporters = sorted([str(x) for x in df['Transporter Name'].dropna().unique().tolist() if str(x).strip() and str(x) != 'nan'])
        
        age_bucket_order = ['<5 Days', '5-10 Days', '10-20 Days', '20-30 Days', '30-60 Days', '>60 Days']
        age_buckets = [ab for ab in age_bucket_order if ab in age_buckets_raw]
        
        # LR Details - check if LR No. column has values
        lr_details = ['LR Generated', 'LR Not Generated']
        
        return {
            "divisions": divisions,
            "age_buckets": age_buckets,
            "transporters": [""] + transporters if transporters else [""],
            "lr_details": lr_details
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/data")
def get_data(division: str = None, age_bucket: str = None, transporter: str = None, po_no: str = None, lr_details: str = None):
    try:
        if df.empty:
            return {"total_records": 0, "data": []}
        
        filtered_df = df.copy()
        
        if division and division != "All" and division.strip():
            filtered_df = filtered_df[filtered_df['Division'].astype(str) == division]
        
        if age_bucket and age_bucket != "All" and age_bucket.strip():
            filtered_df = filtered_df[filtered_df['Age Bucket'].astype(str) == age_bucket]
        
        if transporter and transporter != "All" and transporter.strip():
            filtered_df = filtered_df[filtered_df['Transporter Name'].astype(str) == transporter]
        
        if po_no and po_no.strip():
            filtered_df = filtered_df[filtered_df['Po No'].astype(str).str.contains(po_no.strip(), case=False, na=False)]
        
        if lr_details and lr_details != "All" and lr_details.strip():
            if lr_details == "LR Generated":
                # LR Generated means LR No. column has a value (not blank)
                filtered_df = filtered_df[filtered_df['LR No.'].astype(str).str.strip() != '']
            elif lr_details == "LR Not Generated":
                # LR Not Generated means LR No. column is blank
                filtered_df = filtered_df[filtered_df['LR No.'].astype(str).str.strip() == '']
        
        data_dict = filtered_df.to_dict('records')
        
        for record in data_dict:
            for key, value in record.items():
                if pd.isna(value) or (isinstance(value, float) and np.isnan(value)):
                    record[key] = ''
                elif isinstance(value, (np.integer, np.floating)):
                    record[key] = float(value) if isinstance(value, np.floating) else int(value)
        
        return {
            "total_records": len(filtered_df),
            "data": data_dict
        }
    except Exception as e:
        print(f"Error: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/export")
def export_data(division: str = None, age_bucket: str = None, transporter: str = None, po_no: str = None, lr_details: str = None):
    try:
        if df.empty:
            raise HTTPException(status_code=400, detail="No data available to export")
        
        filtered_df = df.copy()
        
        if division and division != "All" and division.strip():
            filtered_df = filtered_df[filtered_df['Division'].astype(str) == division]
        
        if age_bucket and age_bucket != "All" and age_bucket.strip():
            filtered_df = filtered_df[filtered_df['Age Bucket'].astype(str) == age_bucket]
        
        if transporter and transporter != "All" and transporter.strip():
            filtered_df = filtered_df[filtered_df['Transporter Name'].astype(str) == transporter]
        
        if po_no and po_no.strip():
            filtered_df = filtered_df[filtered_df['Po No'].astype(str).str.contains(po_no.strip(), case=False, na=False)]
        
        if lr_details and lr_details != "All" and lr_details.strip():
            if lr_details == "LR Generated":
                # LR Generated means LR No. column has a value (not blank)
                filtered_df = filtered_df[filtered_df['LR No.'].astype(str).str.strip() != '']
            elif lr_details == "LR Not Generated":
                # LR Not Generated means LR No. column is blank
                filtered_df = filtered_df[filtered_df['LR No.'].astype(str).str.strip() == '']
        
        export_df = filtered_df.copy()
        
        for col in export_df.columns:
            export_df[col] = export_df[col].fillna('')
            if export_df[col].dtype in ['float64', 'float32']:
                def clean_float(x):
                    if isinstance(x, str):
                        return x
                    if pd.isna(x):
                        return ''
                    try:
                        if np.isnan(x):
                            return ''
                    except (TypeError, ValueError):
                        pass
                    return x
                export_df[col] = export_df[col].apply(clean_float)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            export_df.to_excel(writer, index=False, sheet_name='Material In Transit')
        
        output.seek(0)
        
        return StreamingResponse(
            iter([output.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=Material_Intransit_Export.xlsx"}
        )
    except Exception as e:
        print(f"Export error: {str(e)}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Export failed: {str(e)}")

@app.get("/api/status")
def status():
    return {
        "status": "running",
        "data_loaded": not df.empty,
        "total_records": len(df) if not df.empty else 0
    }

@app.get("/", response_class=HTMLResponse)
def read_root():
    html_content = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Unnati Motors Material In Transit Dashboard</title>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/axios/1.6.2/axios.min.js"></script>
        <style>
            * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }

            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                min-height: 100vh;
                padding: 20px;
            }

            .container {
                max-width: 1400px;
                margin: 0 auto;
            }

            .header {
                background: white;
                padding: 30px;
                border-radius: 10px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                margin-bottom: 30px;
            }

            .header h1 {
                color: #667eea;
                font-size: 32px;
                margin-bottom: 10px;
            }

            .header p {
                color: #666;
                font-size: 14px;
            }

            .filters-section {
                background: white;
                padding: 30px;
                border-radius: 10px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                margin-bottom: 30px;
            }

            .filters-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
                gap: 20px;
                margin-bottom: 20px;
            }

            .filter-group {
                display: flex;
                flex-direction: column;
            }

            .filter-group label {
                color: #333;
                font-weight: 600;
                margin-bottom: 8px;
                font-size: 14px;
            }

            .filter-group select,
            .filter-group input {
                padding: 12px 15px;
                border: 2px solid #e0e0e0;
                border-radius: 6px;
                font-size: 14px;
                background-color: white;
                cursor: pointer;
                transition: all 0.3s;
            }

            .filter-group select:hover,
            .filter-group input:hover {
                border-color: #667eea;
            }

            .filter-group select:focus,
            .filter-group input:focus {
                outline: none;
                border-color: #667eea;
                box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
            }

            .button-group {
                display: flex;
                gap: 10px;
                justify-content: flex-end;
                flex-wrap: wrap;
            }

            .btn {
                padding: 12px 30px;
                border: none;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 600;
                cursor: pointer;
                transition: all 0.3s;
            }

            .btn-primary {
                background: #667eea;
                color: white;
            }

            .btn-primary:hover {
                background: #5568d3;
                transform: translateY(-2px);
                box-shadow: 0 6px 12px rgba(102, 126, 234, 0.4);
            }

            .btn-export {
                background: #4caf50;
                color: white;
            }

            .btn-export:hover {
                background: #45a049;
                transform: translateY(-2px);
                box-shadow: 0 6px 12px rgba(76, 175, 80, 0.4);
            }

            .btn-clear {
                background: #ff9800;
                color: white;
            }

            .btn-clear:hover {
                background: #e68900;
                transform: translateY(-2px);
                box-shadow: 0 6px 12px rgba(255, 152, 0, 0.4);
            }

            .stats-section {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                gap: 15px;
                margin-bottom: 30px;
            }

            .stat-card {
                background: white;
                padding: 20px;
                border-radius: 10px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                border-left: 5px solid #667eea;
            }

            .stat-label {
                color: #666;
                font-size: 12px;
                font-weight: 600;
                text-transform: uppercase;
                margin-bottom: 10px;
            }

            .stat-value {
                color: #667eea;
                font-size: 28px;
                font-weight: 700;
            }

            .table-section {
                background: white;
                padding: 30px;
                border-radius: 10px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                overflow-x: auto;
            }

            .table-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 20px;
                flex-wrap: wrap;
                gap: 15px;
            }

            .table-header h2 {
                color: #333;
                font-size: 18px;
            }

            table {
                width: 100%;
                border-collapse: collapse;
                font-size: 13px;
            }

            thead {
                background: #f5f5f5;
            }

            thead th {
                padding: 15px;
                text-align: left;
                font-weight: 600;
                color: #333;
                border-bottom: 2px solid #e0e0e0;
                white-space: nowrap;
            }

            tbody td {
                padding: 12px 15px;
                border-bottom: 1px solid #f0f0f0;
            }

            tbody tr:hover {
                background: #f9f9f9;
            }

            .loading {
                text-align: center;
                padding: 40px;
                color: #667eea;
            }

            .spinner {
                border: 4px solid #f3f3f3;
                border-top: 4px solid #667eea;
                border-radius: 50%;
                width: 40px;
                height: 40px;
                animation: spin 1s linear infinite;
                margin: 0 auto 20px;
            }

            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }

            .no-data {
                text-align: center;
                padding: 40px;
                color: #999;
            }

            .error {
                text-align: center;
                padding: 40px;
                color: #d32f2f;
                background: #ffebee;
                border-radius: 6px;
            }

            .lr-status {
                font-weight: 600;
                padding: 4px 8px;
                border-radius: 4px;
                display: inline-block;
            }

            .lr-generated {
                background: #d4edda;
                color: #155724;
            }

            .lr-not-generated {
                background: #f8d7da;
                color: #721c24;
            }

            .blank {
                color: #999;
                font-style: italic;
            }

            @media (max-width: 768px) {
                .filters-grid {
                    grid-template-columns: 1fr;
                }

                .button-group {
                    flex-direction: column;
                }

                .btn {
                    width: 100%;
                }

                table {
                    font-size: 12px;
                }

                thead th, tbody td {
                    padding: 8px;
                }

                .table-header {
                    flex-direction: column;
                    align-items: flex-start;
                }
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>Unnati Motors Material In Transit Dashboard</h1>
                <p>Real-time tracking of material in transit across all divisions</p>
            </div>

            <div class="filters-section">
                <h3 style="margin-bottom: 20px; color: #333;">Filters</h3>
                <div class="filters-grid">
                    <div class="filter-group">
                        <label for="division">Division</label>
                        <select id="division" onchange="applyFilters()">
                            <option value="All">All Divisions</option>
                        </select>
                    </div>

                    <div class="filter-group">
                        <label for="ageBucket">Age Bucket</label>
                        <select id="ageBucket" onchange="applyFilters()">
                            <option value="All">All Age Buckets</option>
                        </select>
                    </div>

                    <div class="filter-group">
                        <label for="transporter">Transporter Name</label>
                        <select id="transporter" onchange="applyFilters()">
                            <option value="All">All Transporters</option>
                        </select>
                    </div>

                    <div class="filter-group">
                        <label for="poNo">Search SPO No</label>
                        <input type="text" id="poNo" placeholder="Enter SPO No..." />
                    </div>

                    <div class="filter-group">
                        <label for="lrDetails">LR Details</label>
                        <select id="lrDetails" onchange="applyFilters()">
                            <option value="All">All LR Details</option>
                        </select>
                    </div>
                </div>

                <div class="button-group">
                    <button class="btn btn-primary" onclick="applyFilters()">Apply Filters</button>
                    <button class="btn btn-export" onclick="exportData()">Export to Excel</button>
                    <button class="btn btn-clear" onclick="clearFilters()">Clear All</button>
                </div>
            </div>

            <div class="stats-section">
                <div class="stat-card">
                    <div class="stat-label">Total Records</div>
                    <div class="stat-value" id="totalRecords">-</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">LR Generated</div>
                    <div class="stat-value" id="lrGenerated">-</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">LR Not Generated</div>
                    <div class="stat-value" id="lrNotGenerated">-</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Total NDP Value</div>
                    <div class="stat-value" id="ndpValue">-</div>
                </div>
            </div>

            <div class="table-section">
                <div class="table-header">
                    <h2>Material Records</h2>
                    <div style="display: flex; gap: 10px; align-items: center;">
                        <label for="recordsPerPage" style="margin: 0; font-weight: 600; color: #333;">Records per page:</label>
                        <select id="recordsPerPage" onchange="applyPagination()" style="padding: 8px 12px; border: 2px solid #e0e0e0; border-radius: 6px; font-size: 14px;">
                            <option value="10">10 Records</option>
                            <option value="20">20 Records</option>
                            <option value="50">50 Records</option>
                            <option value="100">100 Records</option>
                            <option value="all">>100 Records (All)</option>
                        </select>
                    </div>
                </div>
                <div id="tableContainer">
                    <div class="loading">
                        <div class="spinner"></div>
                        Loading data...
                    </div>
                </div>
                <div id="paginationInfo" style="margin-top: 15px; text-align: center; color: #666; font-size: 14px;"></div>
            </div>
        </div>

        <script>
            let allData = [];
            let currentPage = 1;

            function applyPagination() {
                currentPage = 1;
                renderTable();
            }

            async function loadFilters() {
                try {
                    const response = await axios.get('/api/filters');
                    populateSelect('division', response.data.divisions);
                    populateSelect('ageBucket', response.data.age_buckets);
                    populateSelect('transporter', response.data.transporters);
                    populateSelect('lrDetails', response.data.lr_details);
                } catch (error) {
                    console.error('Error loading filters:', error);
                    document.getElementById('tableContainer').innerHTML = '<div class="error">Error loading filters</div>';
                }
            }

            function populateSelect(elementId, options) {
                const select = document.getElementById(elementId);
                const defaultOption = select.options[0];
                select.innerHTML = '';
                select.appendChild(defaultOption);

                options.forEach(option => {
                    if (option && option.trim()) {
                        const opt = document.createElement('option');
                        opt.value = option;
                        opt.textContent = option;
                        select.appendChild(opt);
                    }
                });
            }

            async function applyFilters() {
                const division = document.getElementById('division').value;
                const ageBucket = document.getElementById('ageBucket').value;
                const transporter = document.getElementById('transporter').value;
                const poNo = document.getElementById('poNo').value;
                const lrDetails = document.getElementById('lrDetails').value;

                document.getElementById('tableContainer').innerHTML = '<div class="loading"><div class="spinner"></div>Loading data...</div>';

                try {
                    const response = await axios.get('/api/data', {
                        params: {
                            division: division === 'All' ? null : division,
                            age_bucket: ageBucket === 'All' ? null : ageBucket,
                            transporter: transporter === 'All' ? null : transporter,
                            po_no: poNo || null,
                            lr_details: lrDetails === 'All' ? null : lrDetails
                        }
                    });

                    allData = response.data.data;
                    currentPage = 1;
                    updateStats();
                    renderTable();
                } catch (error) {
                    console.error('Error loading data:', error);
                    document.getElementById('tableContainer').innerHTML = '<div class="error">Error loading data. Please check console for details.</div>';
                }
            }

            function updateStats() {
                document.getElementById('totalRecords').textContent = allData.length;
                
                const lrGen = allData.filter(row => {
                    const status = String(row.TAT_Invoice_To_LR || '');
                    return status && status !== 'LR not generated';
                }).length;
                
                const lrNotGen = allData.filter(row => String(row.TAT_Invoice_To_LR || '') === 'LR not generated').length;

                const totalNDP = allData.reduce((sum, row) => {
                    const ndpValue = parseFloat(row.NDP) || 0;
                    return sum + ndpValue;
                }, 0);

                document.getElementById('lrGenerated').textContent = lrGen;
                document.getElementById('lrNotGenerated').textContent = lrNotGen;
                document.getElementById('ndpValue').textContent = '₹ ' + totalNDP.toLocaleString('en-IN', {maximumFractionDigits: 2});
            }

            function renderTable() {
                if (allData.length === 0) {
                    document.getElementById('tableContainer').innerHTML = '<div class="no-data">No records found</div>';
                    document.getElementById('paginationInfo').innerHTML = '';
                    return;
                }

                const recordsPerPageValue = document.getElementById('recordsPerPage').value;
                let recordsPerPage = recordsPerPageValue === 'all' ? allData.length : parseInt(recordsPerPageValue);
                
                const startIndex = (currentPage - 1) * recordsPerPage;
                const endIndex = startIndex + recordsPerPage;
                const paginatedData = allData.slice(startIndex, endIndex);
                
                const totalPages = Math.ceil(allData.length / recordsPerPage);

                const columns = [
                    'Division', 'Po No', 'Po Date', 'Sales Order', 'Invoice No', 'Invoice Date',
                    'Part Description', 'Quantity', 'Invoice Amount', 'Transporter Name',
                    'LR No.', 'LR Date', 'TAT_Po_To_Invoice', 'TAT_Invoice_To_LR', 'Age Bucket'
                ];

                let html = '<table><thead><tr>';
                columns.forEach(col => {
                    html += `<th>${col}</th>`;
                });
                html += '</tr></thead><tbody>';

                paginatedData.forEach(row => {
                    html += '<tr>';
                    columns.forEach(col => {
                        let cellValue = row[col];
                        let cellClass = '';

                        if (col === 'TAT_Invoice_To_LR') {
                            const status = String(cellValue || '');
                            if (status === 'LR not generated') {
                                cellClass = 'lr-status lr-not-generated';
                            } else if (status && status.trim()) {
                                cellClass = 'lr-status lr-generated';
                            }
                        }

                        if (!cellValue || cellValue === '') {
                            cellValue = '<span class="blank">-</span>';
                        }

                        html += `<td ${cellClass ? `class="${cellClass}"` : ''}>${cellValue}</td>`;
                    });
                    html += '</tr>';
                });

                html += '</tbody></table>';
                document.getElementById('tableContainer').innerHTML = html;
                
                const startRecord = startIndex + 1;
                const endRecord = Math.min(endIndex, allData.length);
                document.getElementById('paginationInfo').innerHTML = 
                    `Showing ${startRecord} to ${endRecord} of ${allData.length} records (Page ${currentPage} of ${totalPages})`;
            }

            async function exportData() {
                const division = document.getElementById('division').value;
                const ageBucket = document.getElementById('ageBucket').value;
                const transporter = document.getElementById('transporter').value;
                const poNo = document.getElementById('poNo').value;
                const lrDetails = document.getElementById('lrDetails').value;

                const params = new URLSearchParams();
                if (division !== 'All') params.append('division', division);
                if (ageBucket !== 'All') params.append('age_bucket', ageBucket);
                if (transporter !== 'All') params.append('transporter', transporter);
                if (poNo) params.append('po_no', poNo);
                if (lrDetails !== 'All') params.append('lr_details', lrDetails);

                try {
                    const url = `/api/export?${params.toString()}`;
                    
                    const response = await axios({
                        method: 'get',
                        url: url,
                        responseType: 'blob',
                        timeout: 30000
                    });
                    
                    const blob = new Blob([response.data], { 
                        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
                    });
                    const link = document.createElement('a');
                    const objectUrl = URL.createObjectURL(blob);
                    link.href = objectUrl;
                    link.download = 'Material_Intransit_Export.xlsx';
                    document.body.appendChild(link);
                    link.click();
                    
                    setTimeout(() => {
                        document.body.removeChild(link);
                        URL.revokeObjectURL(objectUrl);
                    }, 100);
                    
                } catch (error) {
                    console.error('Error exporting data:', error);
                    alert('Error exporting data. Please check the console for details.');
                }
            }

            function clearFilters() {
                document.getElementById('division').value = 'All';
                document.getElementById('ageBucket').value = 'All';
                document.getElementById('transporter').value = 'All';
                document.getElementById('poNo').value = '';
                document.getElementById('lrDetails').value = 'All';
                applyFilters();
            }

            window.addEventListener('load', () => {
                loadFilters();
                applyFilters();
            });
        </script>
    </body>
    </html>
    """
    return html_content

if __name__ == "__main__":
    print("Starting Unnati Motors Material In Transit Dashboard...")
    print("Dashboard: http://localhost:8000")
    uvicorn.run(app, host="0.0.0.0", port=8000)
