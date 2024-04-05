const xlsx = require("xlsx");
const sqlite3 = require("sqlite3").verbose();

// Function to convert Excel to JSON
function readExcelFile(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetNames = workbook.SheetNames;

  const dataSheets = {};

  sheetNames.forEach((sheet) => {
    const jsonData = xlsx.utils.sheet_to_json(workbook.Sheets[sheet]);
    dataSheets[sheet] = jsonData;
  });

  return dataSheets;
}

// Replace 'your_excel_file.xlsx' with the path to your Excel file
const excelData = readExcelFile("assignment_data.xlsx");

// Establish a connection to the SQLite database
const db = new sqlite3.Database("UnGrey.db", sqlite3.OPEN_READWRITE, (err) => {
  if (err) {
    console.error(err.message);
  } else {
    console.log("Connected to the SQLite database.");
  }
});

// Function to insert data into the specified table
function insertDataIntoTable(tableName, data) {
  db.serialize(() => {
    db.run("BEGIN TRANSACTION");
    console.log(tableName);

    // Prepare SQL statement to insert data
    console.log(Object.keys(data));
    const placeholders = Object.keys(data[0])
      .map(() => "?")
      .join(",");
    const sql = `INSERT INTO ${tableName} (${Object.keys(data[0]).join(
      ","
    )}) VALUES (${placeholders})`;

    const stmt = db.prepare(sql);

    // Insert each row into the table
    data.forEach((row) => {
      stmt.run(Object.values(row), (err) => {
        if (err) {
          console.error(`Insert error on table ${tableName}:`, err.message);
        }
      });
    });

    stmt.finalize();

    db.run("COMMIT");
  });
}

// Assuming the Excel file has three sheets named 'Sheet1', 'Sheet2', 'Sheet3'
// and your SQLite tables are named 'table1', 'table2', 'table3'
insertDataIntoTable("table1", excelData["2"]);
insertDataIntoTable("table2", excelData["4"]);
insertDataIntoTable("table3", excelData["6"]);

// Close the database connection
db.close((err) => {
  if (err) {
    console.error(err.message);
  } else {
    console.log("Closed the SQLite database connection.");
  }
});
