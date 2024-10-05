let data = []; // This holds the initial Excel data
let filteredData = []; // This holds the filtered data after user operations

// Function to load and display the Excel sheet initially
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames[0]; // Load the first sheet
        const sheet = workbook.Sheets[sheetName];

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];

        // Initially display the full sheet
        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Function to display the Excel sheet as an HTML table
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = ''; // Clear existing content

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');

    // Create table headers
    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(key => {
        const th = document.createElement('th');
        th.textContent = key;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Create table rows
    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell !== null ? cell : 'NULL'; // Display 'NULL' for null values
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Apply operation button click event
document.getElementById('apply-operation').addEventListener('click', () => {
    const primaryColumn = document.getElementById('primary-column').value.trim();
    const operationColumns = document.getElementById('operation-columns').value.trim().split(',');
    const operationType = document.getElementById('operation-type').value;
    const operationValue = document.getElementById('operation').value;

    if (!primaryColumn || operationColumns.length === 0) {
        alert('Please enter the required columns.');
        return;
    }

    // Filter the data based on the operations selected
    filteredData = data.filter(row => {
        const primaryValue = row[primaryColumn];
        const checkValues = operationColumns.map(col => row[col]);
        const isNullCheck = operationValue === 'null';
        
        const checks = checkValues.map(value => (isNullCheck ? value === null : value !== null));
        
        if (operationType === 'and') {
            return checks.length && primaryValue === null && checks.every(Boolean);
        } else { // operationType === 'or'
            return checks.length && primaryValue === null && checks.some(Boolean);
        }
    });

    displaySheet(filteredData);
});

// Download button click event
document.getElementById('download-button').addEventListener('click', () => {
    const modal = document.getElementById('download-modal');
    modal.style.display = 'flex'; // Show modal
});

// Close modal button click event
document.getElementById('close-modal').addEventListener('click', () => {
    const modal = document.getElementById('download-modal');
    modal.style.display = 'none'; // Hide modal
});

// Confirm download button click event
document.getElementById('confirm-download').addEventListener('click', () => {
    const filename = document.getElementById('filename').value || 'download';
    const fileFormat = document.getElementById('file-format').value;

    if (fileFormat === 'xlsx') {
        const worksheet = XLSX.utils.json_to_sheet(filteredData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        XLSX.writeFile(workbook, `${filename}.xlsx`);
    } else if (fileFormat === 'csv') {
        const csvContent = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(filteredData));
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.setAttribute('download', `${filename}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    } else if (fileFormat === 'jpg' || fileFormat === 'jpeg') {
        // Implement image download (e.g., take a screenshot of the table)
        html2canvas(document.querySelector('#sheet-content')).then(canvas => {
            canvas.toBlob(blob => {
                const link = document.createElement('a');
                link.href = URL.createObjectURL(blob);
                link.setAttribute('download', `${filename}.${fileFormat}`);
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            });
        });
    } else if (fileFormat === 'pdf') {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();
        doc.autoTable({ html: '#sheet-content table' });
        doc.save(`${filename}.pdf`);
    }

    const modal = document.getElementById('download-modal');
    modal.style.display = 'none'; // Hide modal after download
});

// Load the Excel sheet initially
loadExcelSheet('path_to_your_excel_file.xlsx'); // Change this to your Excel file path
