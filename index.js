const xlsx = require('xlsx');
const fs = require('fs');

// 1. Leer el archivo XLS
const workbook = xlsx.readFile('documentos.xlsx');
const sheetName = workbook.SheetNames[0]; // Primera hoja
const worksheet = workbook.Sheets[sheetName];

// 2. Convertir la hoja a JSON
const rows = xlsx.utils.sheet_to_json(worksheet);
console.log('Datos obtenidos del archivo XLS:', rows);

// 3. Construir las consultas SQL
let insertStatements = '';
rows.forEach((row) => {
    const query = `
    INSERT INTO dbo.documentos_docuware (
        tipo_documento,
        tipo_id_causante,
        no_id_causante,
        tipo_id_beneficiario,
        no_id_beneficiario,
        estado_documento,
        nombre_documento,
        document_id,
        contexto
    ) VALUES (
        '${row.tipo_documento || ''}',
        '${row.tipo_id_causante || ''}',
        '${row.no_id_causante || ''}',
        '${row.tipo_id_beneficiario || ''}',
        '${row.no_id_beneficiario || ''}',
        '${row.estado_documento || ''}',
        '${row.nombre_documento || ''}',
        '${row.document_id || ''}',
        '${row.contexto || ''}'
    );\n`;
    insertStatements += query;
});

// 4. Guardar las consultas en un archivo .sql
fs.writeFileSync('insert_documentos.sql', insertStatements);
console.log('Consultas SQL generadas y guardadas en "insert_documentos.sql".');

