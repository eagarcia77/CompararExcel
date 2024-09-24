/**
 * Autor: Eduardo Augusto García Rodríguez
 * Contacto: eagarcia77@gmail.com
 * Descripción: Este archivo compara dos archivos Excel permitiendo seleccionar las columnas a comparar y muestra las filas no repetidas.
 * Copyright 2024
 * 
 * This file compares two Excel files, allowing the user to select the columns to compare and shows the non-repeated rows.
 */

document.getElementById('excelForm').addEventListener('submit', function(event) {
    event.preventDefault(); // Previene que el formulario recargue la página / Prevents form submission from reloading the page
    const file1 = document.getElementById('file1').files[0];
    const file2 = document.getElementById('file2').files[0];

    if (file1 && file2 && validateFiles(file1, file2)) {
        document.getElementById('loadingSpinner').style.display = 'block'; // Muestra el spinner / Displays loading spinner
        document.getElementById('progressBar').style.display = 'block'; // Muestra la barra de progreso / Displays progress bar
        readExcelFiles(file1, file2); // Leer archivos Excel / Reads the Excel files
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

    // Validación del tipo de archivo / File type validation
    if (!validExtensions.includes(file1Extension) || !validExtensions.includes(file2Extension)) {
        alert('Por favor, sube archivos en formato Excel (.xlsx o .xls).');
        return false;
    }

    return true;
}

/**
 * Lee los archivos Excel y muestra las columnas disponibles para que el usuario seleccione cuál comparar.
 * Reads the Excel files and displays available columns for the user to select.
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

            // Ocultar spinner de carga / Hide loading spinner
            document.getElementById('loadingSpinner').style.display = 'none';
            // Mostrar vista previa de los archivos / Show file preview
            showPreview(sheet1, 'previewTable1');
            showPreview(sheet2, 'previewTable2');

            // Configura las columnas a comparar para cada archivo / Setup column selection for each file
            setupColumnSelect(sheet1, 'column1', 'selectColumn1');
            setupColumnSelect(sheet2, 'column2', 'selectColumn2');
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
    
    const keys = Object.keys(sheet[0]); // Obtener claves de las columnas / Get column keys
    
    // Crear encabezado de tabla / Create table headers
    const trHead = document.createElement('tr');
    keys.forEach(key => {
        const th = document.createElement('th');
        th.innerText = key;
        trHead.appendChild(th);
    });
    thead.appendChild(trHead);
    
    // Crear cuerpo de tabla con las primeras 10 filas / Create table body with first 10 rows
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
 * Configura las opciones de la lista desplegable para seleccionar la columna de comparación.
 * Sets up the dropdown options to select which column to compare.
 */
function setupColumnSelect(sheet, columnSelectId, containerId) {
    const columnSelect = document.getElementById(columnSelectId);
    columnSelect.innerHTML = '';

    const keys = Object.keys(sheet[0]);
    keys.forEach(key => {
        const option = document.createElement('option');
        option.value = key;
        option.innerText = key;
        columnSelect.appendChild(option);
    });

    document.getElementById(containerId).style.display = 'block'; // Mostrar la selección de columnas / Show column selection
}

/**
 * Función para comparar las hojas Excel utilizando las columnas seleccionadas.
 * Compares the Excel sheets using the selected columns.
 */
function compareSheets(sheet1, sheet2) {
    const selectedColumn1 = document.getElementById('column1').value;
    const selectedColumn2 = document.getElementById('column2').value;

    const result = [];
    const sheet2Map = new Set(sheet2.map(row => row[selectedColumn2])); // Crear un mapa de la segunda hoja / Create a map from sheet2

    // Comparar fila por fila de la primera hoja / Compare row by row from sheet1
    sheet1.forEach((row, index) => {
        const rowValue = row[selectedColumn1];
        if (!sheet2Map.has(rowValue)) {
            result.push(row); // Añadir fila si no se encuentra en sheet2 / Add row if not found in sheet2
        }

        const percent = Math.floor((index / sheet1.length) * 100); // Calcular progreso / Calculate progress
        updateProgress(percent); // Actualizar barra de progreso / Update progress bar
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

        document.getElementById('downloadReport').style.display = 'block'; // Mostrar botón de descarga / Show download button
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
        XLSX.writeFile(wb, 'reporte_no_repetidos.xlsx'); // Descargar archivo Excel / Download Excel file
    });
}

/**
 * Actualiza la barra de progreso mientras se procesa la comparación.
 * Updates the progress bar during the comparison process.
 */
function updateProgress(percent) {
    document.getElementById('progressPercent').innerText = `${percent}%`; // Mostrar porcentaje / Show percentage
    document.getElementById('progressBarFill').style.width = `${percent}%`; // Actualizar ancho de la barra / Update progress bar width
}
