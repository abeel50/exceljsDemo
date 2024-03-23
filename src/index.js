const ExcelJS = require('exceljs');
fs = require('fs');

// Create a workbook and add a worksheet
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Sheet1');

// Define data for the table
const data = [
  ['Name', 'Age', 'Country','Description'],
  ['John', 30, 'USA',`Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries,`],
  ['Alice', 25, 'UK',`Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.`],
  ['Bob', 35, 'Canada',`Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, `]
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
    { name: 'Country',filterButton:true },
    {name:'Description', filterButton:true}
  ],
  rows: data.slice(1) // Exclude the header row from data
});
// Adjust column widths
worksheet.getColumn(4).width = 35;

const defaultHeight = 15;

worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
  // Initialize the maximum height for this row
  let maxHeight = 0;
  // Iterate over cells in the row
  row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    cell.alignment = { wrapText: true ,horizontal: 'left', vertical: 'middle' }; // Wrap text for each cell
    // Calculate the height required for the content in this cell
    const contentHeight = cell.value ? (cell.text.length / 2) + 8 : 1; // You can adjust the calculation as needed
    // Update the maximum height if needed
    if (contentHeight > maxHeight) {
      maxHeight = contentHeight;
    }
  });
  // Set the row height based on the maximum height
  row.height = Math.max(defaultHeight ,maxHeight);
});

  // Adjust column widths
  // worksheet.eachRow((row, index) => {
  //   row.alignment = { horizontal: 'left', vertical: 'middle' }; // Center and middle align the content
  // });
  
  const firstRow = worksheet.getRow(1);
  firstRow.font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFFFF' } };
  firstRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'ff1b73f0' } };
  firstRow.alignment = { horizontal: 'center', vertical: 'middle' }; // Center and middle align the content
  firstRow.height = 30;
  

// Freeze the header row
worksheet.views = [
    { state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'B2' }
  ];


// Write the workbook to a buffer
workbook.xlsx.writeBuffer()
  .then(buffer => {
    // Now you can use the buffer as needed
    // For example, you can send it as a response in an HTTP server
      fs.writeFileSync('C:/Users/TK-LPT-533/Documents/output.xlsx', buffer);

    console.log('✅✅✅ Excel buffer created successfully.');
  })
  .catch(error => {
    console.error('❌❌❌ Error occurred:', error);
  });
