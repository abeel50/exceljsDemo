const ExcelJS = require('exceljs');

// Create a workbook and add a worksheet
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Sheet1');

// Define data for the table
const data = [
  ['Name', 'Age', 'Country'],
  ['John', 30, 'USA'],
  ['Alice', 25, 'UK'],
  ['Bob', 35, 'Canada']
];


// Add data to the worksheet
worksheet.addTable({
  name: 'MyTable',
  ref: 'A1',
  headerRow: true,
  totalsRow: false,
  columns: [
    { name: 'Name' ,filterButton:true},
    { name: 'Age', filterButton:true},
    { name: 'Country',filterButton:true }
  ],
  rows: data.slice(1) // Exclude the header row from data
});
// Adjust column widths
worksheet.columns.forEach((column, index) => {
    column.width = 15; // Set width to 15 units for all columns
  });

  worksheet.getRow(1).font = { name: 'Calibri', size: 14, bold: true,color: { argb: 'FFFFFF' } };


  worksheet.getRow(1).fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'ff1b73f0'},
  };
  

// Freeze the header row
worksheet.views = [
    { state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'B2' }
  ];


// Save the workbook to a file
workbook.xlsx.writeFile('data.xlsx')
  .then(() => {
    console.log('Excel file created successfully.');
  })
  .catch((error) => {
    console.error('Error occurred:', error);
  });
