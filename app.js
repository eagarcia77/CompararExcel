/**
 * @author Eduardo Augusto García Rodríguez
 * @contact eagarcia77@gmail.com
 * @description Este archivo compara dos archivos Excel y muestra las filas no repetidas.
 * @copyright 2024
 * 
 * @brief Compares two Excel files, showing rows that don't match and allowing users to download a report.
 */

document.getElementById('excelForm').addEventListener('submit', function(event) {
    event.preventDefault();
    const file1 = document.getElementById('file1').files[0];
    const file2 = document.getElementById('file2').files[0];

    if (file1 && file2 && validateFiles(file1, file2)) {
        document.getElementById('loadingSpinner').style.display = 'block';
        document.getElementById('progressBar').style.display = 'block';
        readExcelFiles(file1, file2);
    }
});

/**
 * Función para validar que los archivos sean de tipo Excel (.xlsx, .xls).
 * Checks if the files uploaded are Excel files (.xlsx or .xls).
 */
function validateFiles(file1, file2) {
    const validExtensions = ['xlsx', 'xls'];
    const file1Extension = file1.name.split('.').pop();
    const file2Extension = file2.name.split('.').pop();

    if (!validExtensions.includes(file1Extension) || !validExtensions.includes(file2Extension)) {
        alert('Por favor, sube archivos en formato Excel (.xlsx o .xls).');
        return false;
    }

    return true;
}

/**
 * Lee los archivos Excel y prepara los datos para la comparación.
 * Reads the Excel files and prepares the data for comparison.
 */
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

/**
 * Muestra una vista previa de los primeros 10 registros de la hoja de Excel.
 * Shows a preview of the first 10 rows from the Excel sheet.
 */
function showPreview(sheet, tableId) {
    const tableContainer = document.getElementById(tableId);
    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');
    
    const keys = Object.keys(sheet[0]);
    
    // Crear encabezado de tabla / Create table headers
    const trHead = document.createElement('tr');
    keys.forEach(key => {
        const th = document.createElement('th');
        th.innerText = key;
        trHead.appendChild(th);
    });
    thead.appendChild(trHead);
    
    // Crear cuerpo de tabla / Create table body
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

/**
 * Configura las opciones para seleccionar las columnas que se van a comparar.
 * Sets up options for selecting columns to be compared.
 */
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

/**
 * Actualiza la barra de progreso mientras se procesa la comparación.
 * Updates the progress bar during the comparison process.
 */
function updateProgress(percent) {
    document.getElementById('progressPercent').innerText = `${percent}%`;
    document.getElementById('progressBarFill').style.width = `${percent}%`;
}

/**
 * Compara las hojas de Excel y muestra los resultados.
 * Compares the Excel sheets and displays the results.
 */
function compareSheets(sheet1, sheet2, selectedColumns) {
    const result = [];
    const sheet2Map = new Set(sheet2.map(row => JSON.stringify(selectedColumns.reduce((obj, key) => {
        obj[key] = row[key];
        return obj;
    }, {}))));
    
    sheet1.forEach((row, index) => {
        const rowString = JSON.stringify(selectedColumns.reduce((obj, key) => {
            obj[key] = row[key];
            return obj;
        }, {}));
        if (!sheet2Map.has(rowString)) {
            result.push(row);
        }

        // Actualizar progreso / Update progress
        const percent = Math.floor((index / sheet1.length) * 100);
        updateProgress(percent);
    });

    displayResults(result);
}

/**
 * Muestra los resultados de las filas no repetidas en una tabla.
 * Displays the non-repeated rows in a table.
 */
function displayResults(result) {
    if (result.length > 0) {
        document.getElementById('statusMessage').innerText = `Se encontraron ${result.length} filas no repetidas.`;
        document.getElementById('resultSection').style.display = 'block';
        
        // Crear tabla de resultados / Create results table
        const table = document.createElement('table');
        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');

        const keys = Object.keys(result[0]);
        const trHead = document.createElement('tr');
        keys.forEach(key => {
            const th = document.createElement('th');
            th.innerText = key;
            trHead.appendChild(th);
        });
        thead.appendChild(trHead);

        result.forEach(row => {
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
        document.getElementById('resultTable').innerHTML = '';
        document.getElementById('resultTable').appendChild(table);

        document.getElementById('downloadReport').style.display = 'block';
        generateReport(result);
    } else {
        document.getElementById('statusMessage').innerText = "No se encontraron filas no repetidas.";
    }
}

/**
 * Genera un archivo Excel con los resultados de las filas no repetidas.
 * Generates an Excel file with the non-repeated rows.
 */
function generateReport(data) {
    document.getElementById('downloadReport').addEventListener('click', function() {
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Reporte');
        XLSX.writeFile(wb, 'reporte_no_repetidos.xlsx');
    });
}
