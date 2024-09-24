/**
 * Autor: Eduardo Augusto García Rodríguez
 * Contacto: eagarcia77@gmail.com
 * Descripción: Este archivo compara dos archivos Excel permitiendo seleccionar las columnas a comparar y muestra las filas no repetidas.
 * Copyright 2024
 */

document.getElementById('excelForm').addEventListener('submit', function(event) {
    event.preventDefault(); // Evitar que el formulario recargue la página
    
    const file1 = document.getElementById('file1').files[0];
    const file2 = document.getElementById('file2').files[0];

    // Validar archivos antes de procesar
    if (file1 && file2 && validateFiles(file1, file2)) {
        document.getElementById('loadingSpinner').style.display = 'block';
        readExcelFiles(file1, file2);
    } else {
        alert('Por favor, sube dos archivos Excel válidos.');
    }
});

/**
 * Función para validar que los archivos sean de tipo Excel (.xlsx, .xls).
 */
function validateFiles(file1, file2) {
    const validExtensions = ['xlsx', 'xls'];
    const file1Extension = file1.name.split('.').pop();
    const file2Extension = file2.name.split('.').pop();

    // Validación del tipo de archivo
    if (!validExtensions.includes(file1Extension) || !validExtensions.includes(file2Extension)) {
        alert('Por favor, sube archivos en formato Excel (.xlsx o .xls).');
        return false;
    }

    return true;
}

/**
 * Función para leer los archivos Excel y mostrar las columnas disponibles para seleccionar.
 */
function readExcelFiles(file1, file2) {
    const reader1 = new FileReader();
    const reader2 = new FileReader();

    // Leer el primer archivo
    reader1.onload = function(e) {
        const data1 = new Uint8Array(e.target.result);
        const workbook1 = XLSX.read(data1, { type: 'array' });
        const sheet1 = XLSX.utils.sheet_to_json(workbook1.Sheets[workbook1.SheetNames[0]]);

        // Leer el segundo archivo
        reader2.onload = function(e) {
            const data2 = new Uint8Array(e.target.result);
            const workbook2 = XLSX.read(data2, { type: 'array' });
            const sheet2 = XLSX.utils.sheet_to_json(workbook2.Sheets[workbook2.SheetNames[0]]);

            // Mostrar las opciones de columnas disponibles
            setupColumnSelect(sheet1, 'column1', 'selectColumn1');
            setupColumnSelect(sheet2, 'column2', 'selectColumn2');
            
            document.getElementById('loadingSpinner').style.display = 'none'; // Ocultar el spinner de carga

            // Agregar evento de submit para la comparación
            document.getElementById('excelForm').addEventListener('submit', function(event) {
                event.preventDefault();
                compareSheets(sheet1, sheet2);
            });
        };

        reader2.readAsArrayBuffer(file2); // Leer el segundo archivo
    };

    reader1.readAsArrayBuffer(file1); // Leer el primer archivo
}

/**
 * Función para configurar las opciones de selección de columnas para cada archivo.
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
 * Función para comparar las columnas seleccionadas de los dos archivos Excel.
 */
function compareSheets(sheet1, sheet2) {
    const selectedColumn1 = document.getElementById('column1').value;
    const selectedColumn2 = document.getElementById('column2').value;

    if (!selectedColumn1 || !selectedColumn2) {
        alert('Por favor, selecciona una columna de cada archivo.');
        return;
    }

    const result = [];
    const sheet2Map = new Set(sheet2.map(row => row[selectedColumn2])); // Crear un Set con los valores de la columna seleccionada en sheet2

    // Comparar los valores de la columna seleccionada en sheet1
    sheet1.forEach(row => {
        const rowValue = row[selectedColumn1]; // Obtener el valor de la columna seleccionada en sheet1
        if (!sheet2Map.has(rowValue)) {
            result.push(row); // Si no se encuentra en sheet2, añadir al resultado
        }
    });

    // Mostrar los resultados
    displayResults(result);
}

/**
 * Muestra los resultados de la comparación en una tabla.
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
    } else {
        document.getElementById('statusMessage').innerText = "No se encontraron filas no repetidas.";
        resultSection.style.display = 'none';
    }
}
