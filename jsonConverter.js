const fs = require("fs");
const XLSX = require("xlsx");

function convertJSONtoExcel(jsonData, sheet, parentKey = "") {
  if (typeof jsonData === "object" && !Array.isArray(jsonData)) {
    for (const key in jsonData) {
      const newKey = parentKey ? `${parentKey}_${key}` : key;
      convertJSONtoExcel(jsonData[key], sheet, newKey);
    }
  } else {
    sheet[parentKey] = jsonData;
  }
}

function jsonToExcel(inputFile, outputFile) {
  try {
    const jsonData = JSON.parse(fs.readFileSync(inputFile, "utf8"));
    const workbook = XLSX.utils.book_new();

    for (const key in jsonData) {
      const sheetData = {};
      convertJSONtoExcel(jsonData[key], sheetData, key);
      const worksheet = XLSX.utils.json_to_sheet([sheetData]);
      XLSX.utils.book_append_sheet(workbook, worksheet, key);
    }

    XLSX.writeFile(workbook, outputFile);

    console.log(`Conversion successful! Excel file saved at: ${outputFile}`);
  } catch (error) {
    console.error("Error:", error.message);
  }
}

jsonToExcel("./data1.json", "output.xlsx");
