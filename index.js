// const XlsxPopulate = require("xlsx-populate");

// XlsxPopulate.fromFileAsync("Data.xlsx").then((workbook) => {
//   const values = workbook.sheet("Customers").usedRange().value();
//   console.log(values);
// });

// function sortGoldCustomer(values){
// return{

// }}

// XlsxPopulate.fromFileAsync("Data.xlsx").then((workbook) => {
//   const newSheet = workbook.addSheet("Info");

//   return workbook.toFileAsync("new.xlsx");
//   console.log(values);
// });

//To change the sheet name
//const newSheet = workbook.sheet(0).name("Gold_Customers")
//const newSheet = workbook.sheet(0).name("Bronze_Customers")

//   workbook.sheet("Sheet1").cell("A1").value("Name");

//   return workbook.toFileAsync("result.xlsx");
// });

const fs = require("fs");
const XLSX = require("xlsx");
const { parse, format, addDays } = require("date-fns");

// Load the Excel file
const workbook = XLSX.readFile("./Data.xlsx");

// Get the "Customers" sheet
const sheetName = "Customers";
const sheet = workbook.Sheets[sheetName];

// Initialize arrays for bronze and gold rows
const bronzeRows = [];
const goldRows = [];
const customerRows = [];

// Process each row (assuming the header is in the first row)
const range = XLSX.utils.decode_range(sheet["!ref"]);
const lastRow = range.e.r + 1; // Adding 1 because it's a zero-based index
const lastColumn = range.e.c;

for (let row = 2; row <= lastRow; row++) {
  const rowData = [];

  for (
    let col = "A";
    col <= XLSX.utils.encode_col(lastColumn);
    col = String.fromCharCode(col.charCodeAt(0) + 1)
  ) {
    const cell = sheet[col + row];

    if (cell) {
      if (col === "H") {
        // Handle the date column (Column H)
        let formattedDate;

        if (typeof cell.v === "number") {
          // Numeric date value (Excel date serial number)
          const excelDateValue = cell.v;
          const jsDate = new Date(
            (excelDateValue - (25567 + 1)) * 86400 * 1000
          );
          const adjustedJsDate = addDays(jsDate, -1); // Adjust by subtracting one day
          formattedDate = format(adjustedJsDate, "dd/MM/yyyy");
        } else if (typeof cell.v === "string") {
          // String date value in MM/DD/YYYY format
          const dateRegex = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;

          if (dateRegex.test(cell.v)) {
            const [, month, day, year] = cell.v.match(dateRegex);
            formattedDate = `${day}/${month}/${year}`;
          } else {
            // Invalid date format, push as is
            formattedDate = cell.v;
          }
        } else {
          // Unsupported date format, push as is
          formattedDate = cell.v;
        }

        rowData.push(formattedDate);
      } else {
        rowData.push(cell.v);
      }
    } else {
      rowData.push(undefined); // Push undefined when the cell is undefined
    }
  }

  if (rowData.length > 0) {
    const hasEmptyCell = rowData.some(
      (cellValue) =>
        cellValue === null || cellValue === undefined || cellValue === ""
    );

    if (hasEmptyCell) {
      bronzeRows.push(rowData);
    } else {
      goldRows.push(rowData);
    }
    customerRows.push(rowData);
  }
}

// Create a new workbook
const newWorkbook = XLSX.utils.book_new();

// Create bronze and gold sheets
const header = [
  "Customer_ID",
  "Name",
  "Age",
  "Gender",
  "Email",
  "Phone",
  "Address",
  "Collection_Date",
];
const bronzeSheetData = [header, ...bronzeRows];
const goldSheetData = [header, ...goldRows];
const customersSheetData = [header, ...customerRows];

const bronzeSheet = XLSX.utils.aoa_to_sheet(bronzeSheetData);
const goldSheet = XLSX.utils.aoa_to_sheet(goldSheetData);
const customersSheet = XLSX.utils.aoa_to_sheet(customersSheetData);

// Set the date format for the "Collection_Date" column (Column H) in both sheets
const dateFormat = { numFmt: "dd/mm/yyyy" };

// Add bronze and gold sheets to the workbook
XLSX.utils.book_append_sheet(newWorkbook, bronzeSheet, "bronze");
XLSX.utils.book_append_sheet(newWorkbook, goldSheet, "gold");
XLSX.utils.book_append_sheet(newWorkbook, customersSheet, "Customers");

// Apply the date format to the "Collection_Date" column in both sheets
bronzeSheet["H1"].z = dateFormat.numFmt;
goldSheet["H1"].z = dateFormat.numFmt;

// Write the new workbook to a file
XLSX.writeFile(newWorkbook, "output_file.xlsx");

console.log("File processing complete.");
