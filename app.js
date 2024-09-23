document.getElementById('excelForm').addEventListener('submit', function(event) {
    event.preventDefault();
    const file1 = document.getElementById('file1').files[0];
    const file2 = document.getElementById('file2').files[0];

    if (file1 && file2) {
        document.getElementById('loadingSpinner').style.display = 'block';
        readExcelFiles(file1, file2);
    }
});

function readExcelFiles(file1, file2) {
    const reader1 = new FileReader();
    const reader2 = new FileReader();

    reader1.onload = function(e) {
        const data1 = new Uint8Array(e.target.result);
        const workbook1 = XLSX.read(data1, { type: 'array' });
        const sheet1 = XLSX.utils.sheet_to_json(workbook1.Sheets[workbook1.SheetNames[0]]);

        reader2.onload = function(e) {
            const data2 = new Uint8Array(e.target.result);
            const workbook2 = XLSX.read(data2, { type: 'array' });
            const sheet2 = XLSX.utils.sheet_to_json(workbook2.Sheets[workbook2.SheetNames[0]]);

            document.getElementById('loadingSpinner').style.display = 'none';
            showPreview(sheet1, 'previewTable1');
            showPreview(sheet2, 'previewTable2');
            setupColumnSelect(sheet1);

            document.getElementById('columnSelect').style.display = 'block';
            document.getElementById('excelForm').addEventListener('submit', function(event) {
                event.preventDefault();
                const selectedColumns = Array.from(document.getElementById('columns').selectedOptions).map(opt => opt.value);
                compareSheets(sheet1, sheet2, selectedColumns);
            });
        };

        reader2.readAsArrayBuffer(file2);
    };

    reader1.readAsArrayBuffer(file1);
}

function showPreview(sheet, tableId) {
    const tableContainer = document.getElementById(tableId);
    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');
    
    const keys = Object.keys(sheet[0]);
    
    // Head
    const trHead = document.createElement('tr');
    keys.forEach(key => {
        const th = document.createElement('th');
        th.innerText = key;
        trHead.appendChild(th);
    });
    thead.appendChild(trHead);
    
    // Body
    sheet.slice(0, 10).forEach(row => {
        const tr = document.createElement('tr');
        keys.forEach(key => {
            const td = document.createElement('td');
            td.innerText = row[key] || '';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    
    table.appendChild(thead);
    table.appendChild(tbody);
    tableContainer.innerHTML = '';
    tableContainer.appendChild(table);
}

function setupColumnSelect(sheet) {
    const columnSelect = document.getElementById('columns');
    columnSelect.innerHTML = '';

    const keys = Object.keys(sheet[0]);
    keys.forEach(key => {
        const option = document.createElement('option');
        option.value = key;
        option.innerText = key;
        columnSelect.appendChild(option);
    });
}

function compareSheets(sheet1, sheet2, selectedColumns) {
    const result = [];
    const sheet2Strings = sheet2.map(row => JSON.stringify(selectedColumns.reduce((obj, key) => {
        obj[key] = row[key];
        return obj;
    }, {})));

    sheet1.forEach(row => {
        const rowString = JSON.stringify(selectedColumns.reduce((obj, key) => {
            obj[key] = row[key];
            return obj;
        }, {}));
        if (!sheet2Strings.includes(rowString)) {
            result.push(row);
        }
    });

    displayResults(result);
}

function displayResults(result) {
    if (result.length > 0) {
        document.getElementById('statusMessage').innerText = `Se encontraron ${result.length} filas no repetidas.`;
        document.getElementById('resultSection').style.display = 'block';
        document.getElementById('downloadReport').style.display = 'block';
        generateReport(result);
    } else {
        document.getElementById('statusMessage').innerText = "No se encontraron filas no repetidas.";
    }
}

function generateReport(data) {
    document.getElementById('downloadReport').addEventListener('click', function() {
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Reporte');
        XLSX.writeFile(wb, 'reporte_no_repetidos.xlsx');
    });
}
