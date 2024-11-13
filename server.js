import express from "express";
import { uploadExcel } from "./fileUpload.js";
import ExcelJS from "exceljs";
import bodyParser from "body-parser";
const app = express();

app.use(bodyParser.json());

app.use(express.json());

app.listen(5005, (req, res) => {
  console.log("SERVER RUNNING ON PORT 5005");
});
app.get("/", async (req, res) => {
  res.json("WORKING");
});
app.post("/exceltojson", uploadExcel.single("excel"), async (req, res) => {
  const { key, value, sheet } = req.body;
  async function parseExcel() {
    try {
      // Use exceljs to read the workbook from the buffer
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(req.file.buffer);
      const worksheet = workbook.getWorksheet(+sheet); // Assuming the data is in the first sheet

      // Get the data from the worksheet
      const data = {}; // Create an empty object to store the key-value pairs
      console.log(worksheet, "SELECTED WORKSHEET");

      worksheet.eachRow({ includeEmpty: true }, (row) => {
        if (row.number === 1) return; // Skip the first row (header row)

        const thirdColumnValue = row.getCell(+key).text;
        const fourthColumnValue = row.getCell(+value).text;

        if (
          thirdColumnValue &&
          thirdColumnValue !== null &&
          thirdColumnValue !== ""
        ) {
          data[thirdColumnValue] = fourthColumnValue;
        }
      });

      const jsonResult = JSON.stringify(data, null, 2); // Convert the object to a nicely formatted JSON string
      console.log(jsonResult); // Print or use the JSON string as needed

      return data;
    } catch (error) {
      console.error("Error parsing Excel file:", error);
    }
  }
  const ExcelData = await parseExcel();

  res.json(ExcelData);
});

app.post(
  "/exceltojsonMultiple",
  uploadExcel.single("excel"),
  async (req, res) => {
    const { key, value, sheet } = req.body;
    async function parseExcel() {
      try {
        // Use exceljs to read the workbook from the buffer
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.getWorksheet(+sheet); // Assuming the data is in the first sheet

        // Get the data from the worksheet
        const data = []; // Create an empty object to store the key-value pairs
        console.log(worksheet, "SELECTED WORKSHEET");

        worksheet.eachRow({ includeEmpty: true }, (row) => {
          if (row.number === 1) return; // Skip the first row (header row)

          const thirdColumnValue = row.getCell(+key).text;
          const results = value;
          var returnValues = [];
          if (results && results.length) {
            results.map((item) => {
              returnValues.push(row.getCell(+item).text);
            });

            // const fourthColumnValue = row.getCell(+value).text;
          }

          if (
            thirdColumnValue &&
            thirdColumnValue !== null &&
            thirdColumnValue !== ""
          ) {
            if (
              returnValues[0] !== "" &&
              returnValues[1] !== "" &&
              returnValues[0] !== "Currency"
            ) {
              data.push({
                country: thirdColumnValue,
                currency_name: returnValues[0],
                currency: returnValues[1],
              });
            }

            // data[thirdColumnValue] = returnValues;
          }
        });

        const jsonResult = JSON.stringify(data, null, 2); // Convert the object to a nicely formatted JSON string
        console.log(jsonResult); // Print or use the JSON string as needed

        return data;
      } catch (error) {
        console.error("Error parsing Excel file:", error);
      }
    }
    const ExcelData = await parseExcel();

    res.json(ExcelData);
  }
);

app.post("/jsontoexcel", async (req, res) => {
  try {
    const { json } = req.body;

    // Create a new workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("excel_file");

    // Convert JSON data to Excel rows
    Object.entries(json).forEach(([key, value]) => {
      worksheet.addRow([key, value]);
    });

    // Set column headers
    worksheet.getRow(1).values = ["Key", "Value"];

    // Adjust column width to fit the content
    worksheet.columns.forEach((column) => {
      // let maxColumnWidth = column.header.length;
      column.eachCell({ includeEmpty: true }, (cell) => {
        const columnWidth = cell.value ? String(cell.value).length : 0;
        // maxColumnWidth = Math.max(maxColumnWidth, columnWidth);
      });
      // column.width = Math.min(30, maxColumnWidth); // Limit the maximum column width to 30 characters
    });

    // Prepare and send the Excel file as response
    res.setHeader("Content-Disposition", "attachment; filename=file.xlsx");
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    const buffer = await workbook.xlsx.writeBuffer();
    res.send(buffer);
  } catch (error) {
    console.error("Error:", error);
    res.status(500).send("Internal Server Error");
  }
});
