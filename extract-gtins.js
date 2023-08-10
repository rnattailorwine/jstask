const ExcelJS = require('exceljs');
const fs = require('fs');

const workbook = new ExcelJS.Workbook();

async function extractGTINs(filePath) {
  try {
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.getWorksheet(1);

    const gtins = [];

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber !== 1) { 
        gtins.push(row.getCell(1).value); 
      }
    });

    return gtins;
  } catch (error) {
    console.error('Error extracting GTINs:', error);
    return [];
  }
}

const filePath = '/Users/finjamanski/Desktop/jstask/prechtel/prechtel.xlsx';

extractGTINs(filePath).then((gtins) => {
  console.log('Extrahierte GTINs:', gtins);
});
