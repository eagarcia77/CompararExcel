/**
 * Autor: Eduardo Augusto García Rodríguez
 * Contacto: eagarcia77@gmail.com
 * Descripción: Este archivo compara dos archivos Excel permitiendo seleccionar las columnas a comparar y muestra las filas no repetidas.
 * Copyright 2023
 */

// Evento para el primer archivo
document.getElementById('file1').addEventListener('change', function(event) {
    const file1 = event.target.files[0];
    if (file1 && validateFiles(file1)) {
        readExcelFile(file1, 'column1', 'selectColumn1'); // Leer y mostrar columnas para el primer archivo
    }
});

// Evento para el segundo archivo
document.getElementById('file2').addEventListener('change', function(event) {
    const file2 = event.target.files[0];
    if (file2 && validateFiles(file2)) {
        readExcelFile(file2, 'column2', 'selectColumn2'); // Leer y mostrar columnas para el segundo archivo
    }
});

/**
 * Validación de los archivos Excel.
 */
function validateFiles(file) {
    const validExtensions = ['xlsx', 'xls'];
    const fileExtension = file.name.split('.').pop().toLowerCase();

    if (!validExtensions.includes(fileExtension)) {
        alert('Por favor, sube archivos en formato Excel (.xlsx o .xls).');
        return false;
    }

    return true;
}

/**
 * Lee un archivo Excel y muestra las columnas disponibles.
 */
function readExcelFile(file, columnSelectId, containerId) {
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

        setupColumnSelect(sheet, columnSelectId, containerId); // Mostrar las columnas del archivo
    };

    reader.readAsArrayBuffer(file); // Leer el archivo Excel
}

/**
 * Configura las opciones de selección de columnas para cada archivo.
 */
function setupColumnSelect(sheet, columnSelectId, containerId) {
    const columnSelect = document.getElementById(columnSelectId);
    columnSelect.innerHTML = ''; // Limpiar cualquier opción previa

    const keys = Object.keys(sheet[0]); // Obtener las claves (nombres de columnas)

    // Rellenar el select con las opciones de las columnas
    keys.forEach(key => {
        const option = document.createElement('option');
        option.value = key;
        option.innerText = key;
        columnSelect.appendChild(option);
    });

    document.getElementById(containerId).style.display = 'block'; // Mostrar el contenedor del select
}

/**
 * Comparar archivos al enviar el formulario.
 */
document.getElementById('excelForm').addEventListener('submit', function(event) {
    event.preventDefault();

    const file1 = document.getElementById('file1').files[0];
    const file2 = document.getElementById('file2').files[0];

    if (file1 && file2) {
        const selectedColumn1 = document.getElementById('column1').value;
        const selectedColumn2 = document.getElementById('column2').value;

        if (selectedColumn1 && selectedColumn2) {
            readAndCompareExcelFiles(file1, file2, selectedColumn1, selectedColumn2); // Realizar la comparación
        } else {
            alert('Por favor, selecciona una columna de cada archivo.');
        }
    } else {
        alert('Por favor, sube los dos archivos Excel.');
    }
});

/**
 * Lee ambos archivos Excel y compara las columnas seleccionadas.
 */
function readAndCompareExcelFiles(file1, file2, selectedColumn1, selectedColumn2) {
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

            compareSheets(sheet1, sheet2, selectedColumn1, selectedColumn2); // Comparar las hojas
        };

        reader2.readAsArrayBuffer(file2);
    };

    reader1.readAsArrayBuffer(file1);
}

/**
 * Compara las columnas seleccionadas de ambos archivos Excel.
 */
function compareSheets(sheet1, sheet2, selectedColumn1, selectedColumn2) {
    const result = [];
    const sheet2Map = new Set(sheet2.map(row => row[selectedColumn2]));

    // Comparar los valores de la columna seleccionada en sheet1
    sheet1.forEach(row => {
        const rowValue = row[selectedColumn1];
        if (!sheet2Map.has(rowValue)) {
            result.push(row); // Si no se encuentra en sheet2, añadir al resultado
        }
    });

    // Mostrar los resultados
    displayResults(result);
}

/**
 * Muestra los resultados de la comparación y permite descargar el reporte.
 */
function displayResults(result) {
    const resultSection = document.getElementById('resultSection');
    const resultTable = document.getElementById('resultTable');
    
    resultTable.innerHTML = ''; // Limpiar la tabla previa

    if (result.length > 0) {
        document.getElementById('statusMessage').innerText = `Se encontraron ${result.length} filas no repetidas.`;
        resultSection.style.display = 'block';

        // Crear una tabla de resultados
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
        resultTable.appendChild(table);

        document.getElementById('downloadReport').style.display = 'block'; // Mostrar botón de descarga
        generateReport(result); // Generar el reporte descargable
    } else {
        document.getElementById('statusMessage').innerText = "No se encontraron filas no repetidas.";
        resultSection.style.display = 'none';
    }
}

/**
 * Genera un archivo Excel con los resultados de las filas no repetidas.
 */
function generateReport(data) {
    document.getElementById('downloadReport').addEventListener('click', function() {
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Reporte');
        XLSX.writeFile(wb, 'reporte_no_repetidos.xlsx'); // Descargar archivo Excel
    });
}
