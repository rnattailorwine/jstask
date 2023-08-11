const fs = require('fs');
const xlsx = require('xlsx');

// Dateiname für den Kunden "schenke"
const fileName = 'schenke.xlsx';

// Name der Spalte mit den GTINs
const columnName = 'GTIN';

// Lese die Excel-Datei ein
const workbook = xlsx.readFile(fileName);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

console.log('Inhalt des Arbeitsblatts:');
console.log(sheet);

// Extrahiere die GTINs aus der Excel-Tabelle
const gtins = [];
for (const cellAddress in sheet) {
  if (cellAddress.includes(columnName)) {
    const cellValue = sheet[cellAddress].v;
    if (typeof cellValue === 'string') {
      gtins.push(cellValue);
    }
  }
}

console.log('Extrahierte GTINs für Schenke:');
console.log(gtins);

if (gtins.length > 0) {
  console.log(`Extrahierte GTINs für Schenke: ${gtins.join(', ')}`);
} else {
  console.log('Keine GTINs für Schenke gefunden.');
}
