const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const os = require('os');

// Function to extract all string values from a specified row, excluding values containing "DUMMY"
function extractRowValues(worksheet, rowNumber) {
  const rowValues = [];
  const range = xlsx.utils.decode_range(worksheet['!ref']);

  for (let col = range.s.c; col <= range.e.c; col++) {
    const cellAddress = xlsx.utils.encode_cell({ c: col, r: rowNumber });
    const cell = worksheet[cellAddress];
    const value = cell ? cell.v : null;

    if (value !== null && typeof value === 'string' && !value.includes('DUMMY')) {
      rowValues.push({ value, col });
    }
  }

  return rowValues;
}

// Function to extract data from a specific column starting from a given row for a certain number of rows
function extractColumnData(worksheet, col, startRow, numRows) {
  const columnData = [];
  for (let row = startRow; row < startRow + numRows; row++) {
    const cellAddress = xlsx.utils.encode_cell({ c: col, r: row });
    const cell = worksheet[cellAddress];
    const value = cell ? cell.v : 'N/A';
    columnData.push(value);
  }
  return columnData;
}

// Function to generate time intervals for 1440 rows
function generateTimeIntervals(numRows) {
  const timeIntervals = [];
  const startDate = new Date(2000, 0, 1, 0, 0, 0);

  for (let i = 0; i < numRows; i++) {
    const endDate = new Date(startDate);
    endDate.setMinutes(startDate.getMinutes() + 1);
    const startTime = startDate.toTimeString().slice(0, 5);
    const endTime = endDate.toTimeString().slice(0, 5);
    timeIntervals.push(`${startTime} - ${endTime}`);
    startDate.setMinutes(startDate.getMinutes() + 1);
  }

  return timeIntervals;
}

function processFilesBatch(inputFolder, inputFiles, batchNumber, outputFolder, zone) {
  // Initializing the output data with headers
  const outputData = [['Zone', 'Name of Station', 'Date', 'Time', 'SCADA Tag', 'Data', 'Range']];

  inputFiles.forEach(file => {
    const inputFilePath = path.join(inputFolder, file);
    processFile(inputFilePath, outputData, zone);
  });

  // Create a new workbook and worksheet for the consolidated output data
  const outputWorkbook = xlsx.utils.book_new();
  const outputWorksheet = xlsx.utils.aoa_to_sheet(outputData);
  xlsx.utils.book_append_sheet(outputWorkbook, outputWorksheet, 'Sheet1');

  // Save the output workbook to an Excel file in the corresponding output folder
  const outputFilePath = path.join(outputFolder, `Consolidated_SCADA_Tag_Data_Batch_${batchNumber}.xlsx`);
  xlsx.writeFile(outputWorkbook, outputFilePath);

  console.log(`Batch ${batchNumber}: Consolidated SCADA Tag data for zone ${zone} has been successfully written to ${outputFilePath}`);
}

function processFile(inputFilePath, outputData, zone) {
  // Read the Excel file
  const workbook = xlsx.readFile(inputFilePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Extract string values from the specified rows
  const thirdRowValues = extractRowValues(worksheet, 2);
  console.log(`String values in the third row (excluding values containing "DUMMY") from ${inputFilePath}:`, thirdRowValues);

  const twoThousandThirdRowValues = extractRowValues(worksheet, 2002);
  console.log(`String values in the 2003rd row (excluding values containing "DUMMY") from ${inputFilePath}:`, twoThousandThirdRowValues);

  const fourThousandThirdRowValues = extractRowValues(worksheet, 4002);
  console.log(`String values in the 4003rd row (excluding values containing "DUMMY") from ${inputFilePath}:`, fourThousandThirdRowValues);

  const sixThousandThirdRowValues = extractRowValues(worksheet, 6002);
  console.log(`String values in the 6003rd row (excluding values containing "DUMMY") from ${inputFilePath}:`, sixThousandThirdRowValues);

  const eightThousandThirdRowValues = extractRowValues(worksheet, 8002);
  console.log(`String values in the 8003rd row (excluding values containing "DUMMY") from ${inputFilePath}:`, eightThousandThirdRowValues);

  // Extraction of Name of the Station from cell A1
  const nameOfStation = worksheet['A1'] ? worksheet['A1'].v : 'N/A';

  // Extraction of Date from cell B6
  let date = worksheet['B6'] ? worksheet['B6'].v : 'N/A';
  if (!isNaN(date)) {
    date = xlsx.SSF.format('yyyy-mm-dd', date);
  }

  // Generating time intervals for 1440 rows
  const timeIntervals = generateTimeIntervals(1440);

  function addDataToOutput(rowValues, rangeStart) {
    rowValues.forEach(({ value, col }) => {
      const columnData = extractColumnData(worksheet, col, rangeStart + 5, 1440); // Extract data from row 4 in range
      columnData.forEach((dataValue, index) => {
        outputData.push([zone, nameOfStation, date, timeIntervals[index], value, dataValue, `${rangeStart}-${rangeStart + 2000}`]);
      });
    });
  }

  // Extract data for each SCADA tag from the rows and add to the outputData
  addDataToOutput(thirdRowValues, 0);

  if (twoThousandThirdRowValues.length > 0) {
    for (let i = 0; i < 4; i++) {
      outputData.push([]);
    }
    addDataToOutput(twoThousandThirdRowValues, 2000);
  }

  if (fourThousandThirdRowValues.length > 0) {
    for (let i = 0; i < 4; i++) {
      outputData.push([]);
    }
    addDataToOutput(fourThousandThirdRowValues, 4000);
  }

  if (sixThousandThirdRowValues.length > 0) {
    for (let i = 0; i < 4; i++) {
      outputData.push([]);
    }
    addDataToOutput(sixThousandThirdRowValues, 6000);
  }

  if (eightThousandThirdRowValues.length > 0) {
    for (let i = 0; i < 4; i++) {
      outputData.push([]);
    }
    addDataToOutput(eightThousandThirdRowValues, 8000);
  }
}

// Array of input folders to process
const inputFolders = ['./BGK_testing', './BGM_testing', './HSN_testing'];

// Start monitoring performance
const startTime = Date.now();
let fileCount = 0;

// Process each input folder
inputFolders.forEach(inputFolder => {
  // Extract the zone from the folder name
  const zoneMatch = path.basename(inputFolder).match(/^(\w{3})_/);
  const zone = zoneMatch ? zoneMatch[1] : 'Unknown';

  // Determine the corresponding output folder
  const outputFolder = path.join(inputFolder, 'output');
  if (!fs.existsSync(outputFolder)) {
    fs.mkdirSync(outputFolder, { recursive: true });
  }

  fs.readdir(inputFolder, (err, files) => {
    if (err) {
      console.error(`Error reading the folder ${inputFolder}:`, err);
      return;
    }

    let batchFiles = [];
    let batchNumber = 1;

    files.forEach(file => {
      if (path.extname(file) === '.xls' || path.extname(file) === '.xlsx') {
        batchFiles.push(file);
        if (batchFiles.length === 5) {
          processFilesBatch(inputFolder, batchFiles, batchNumber, outputFolder, zone);
          batchFiles = [];
          batchNumber++;
          fileCount += 5;
        }
      }
    });

    // Process any remaining files
    if (batchFiles.length > 0) {
      processFilesBatch(inputFolder, batchFiles, batchNumber, outputFolder, zone);
      fileCount += batchFiles.length;
    }
  });
});

// End monitoring performance
const endTime = Date.now();
const duration = (endTime - startTime) / 1000; // Duration in seconds

console.log(`Processed ${fileCount} files in ${duration} seconds`);
console.log(`System Memory Usage: ${(process.memoryUsage().heapUsed / 1024 / 1024).toFixed(2)} MB`);
console.log(`CPU Load: ${os.loadavg().map(load => load.toFixed(2)).join(', ')}`);
