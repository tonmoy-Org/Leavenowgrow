const path = require("path");
const ejs = require("ejs");
const excel = require("exceljs");
const separateTagValues = require("./separateTagValues");
const { arrayToObject } = require("./arrayToObject");

const generateGoalsWorksheetExcel = async (goalsWorksheets, cName) => {
  const templatePathDocx = path.join(
    __dirname,
    "../views",
    "goals-worksheets-pdf.ejs"
  );

  let htmlPageDocx = "";
  htmlPageDocx = await ejs.renderFile(templatePathDocx, {
    goalsWorksheets: goalsWorksheets,
    print: true,
    company: cName,
  });

  let tValues = separateTagValues(htmlPageDocx);

  const workbook = new excel.Workbook();
  const worksheet = workbook.addWorksheet("Sheet 1");

  const childArrays = goalsWorksheets.map((item) => {
    const worksheets = arrayToObject(item?.worksheets);

    return { ...worksheets, name: item?.name };
  });

  worksheet.getColumn(1).alignment = { wrapText: true, vertical: "top" };

  worksheet.getColumn(1).font = { bold: true };
  worksheet.getColumn(1).values = [
    "Name",
    "Goal #1",
    "Benefit #1",
    "Metric #1",
    "Action #1",
    "Goal #2",
    "Benefit #2",
    "Metric #2",
    "Action #2",
    "Goal #3",
    "Benefit #3",
    "Metric #3",
    "Action #3",
    "Goal #4",
    "Benefit #4",
    "Metric #4",
    "Action #4",
  ];

  // Get the maximum length of child arrays
  const maxChildArrayLength = childArrays.length;
  // console.log(maxChildArrayLength);

  let $a = 2;

  // console.log(childArrays)
  // process.abort()
  childArrays.forEach((obj) => {
    //  arr.forEach((obj) => {
    //  worksheet.getColumn($a).values = values;
    const data = [
      obj.name,
      obj.goal0,
      obj.benefit0,
      obj.metric0,
      obj.action0,
      obj.goal1,
      obj.benefit1,
      obj.metric1,
      obj.action1,
      obj.goal2,
      obj.benefit2,
      obj.metric2,
      obj.action2,
      obj.goal3,
      obj.benefit3,
      obj.metric3,
      obj.action3,
    ];

    worksheet.getColumn($a).alignment = { wrapText: true, vertical: "top" };

    worksheet.getColumn($a).values = data;

    // })
    $a++;
  });

  // console.log(childArrays)
  worksheet.getRow(1).font = { bold: true };
  worksheet.getColumn(1).eachCell((cell, rowNumber) => {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "CCFFFF" },
    };
  });

  worksheet.getRow(1).eachCell((cell, colNumber) => {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "CCFFFF" },
    };
  });

  // Set borders for all cells with values
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      if (cell.value) {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      }
    });
  });

  worksheet.columns.forEach((column) => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const columnLength = cell.value ? cell.value.length : 0;
      maxLength = Math.max(maxLength, columnLength);
    });
    column.width = 30;
  });

  const excelBuffer = await workbook.xlsx.writeBuffer();

  return excelBuffer;
};

function reshapeArray(arr, rows, cols) {
  if (arr.length !== rows * cols) {
    throw new Error("Array length does not match the dimensions.");
  }

  const newArray = [];
  for (let i = 0; i < rows; i++) {
    newArray.push(arr.slice(i * cols, (i + 1) * cols));
  }
  return newArray;
}

const generateGoalsWorksheetExcelV2 = async (goalsWorksheet, cName) => {
  const templatePathDocx = path.join(
    __dirname,
    "../views",
    "goals-worksheets-pdf-v2.ejs"
  );

  let htmlPageDocx = "";
  htmlPageDocx = await ejs.renderFile(templatePathDocx, {
    goalsWorksheet: goalsWorksheet,
    print: true,
    company: cName,
  });

  let tValues = separateTagValues(htmlPageDocx);

  const reshapedTdValues = reshapeArray(tValues?.tdValues, 5, 5);

  const workbook = new excel.Workbook();
  const worksheet = workbook.addWorksheet("Sheet 1");

  const childArrays = goalsWorksheet.worksheets;

  // the header row will be colored, bold etc.
  worksheet.getRow(1).values = tValues["thValues"];
  worksheet.getRow(1).alignment = { wrapText: true, vertical: "top" };
  worksheet.getRow(1).font = { bold: true };
  worksheet.getRow(1).eachCell((cell, colNumber) => {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "CCFFFF" },
    };
  });

  // the rest of the rows will folloe this configurations
  let $a = 2;
  reshapedTdValues.forEach((tdValue) => {
    worksheet.getRow($a).values = tdValue;
    worksheet.getRow($a).alignment = { wrapText: true, vertical: "top" };
    $a++;
  });

  // worksheet.getColumn(1).eachCell((cell, rowNumber) => {
  //   cell.fill = {
  //     type: "pattern",
  //     pattern: "solid",
  //     fgColor: { argb: "CCFFFF" },
  //   };
  // });

  // Set borders for all cells with values
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      if (cell.value) {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      }
    });
  });

  worksheet.columns.forEach((column) => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const columnLength = cell.value ? cell.value.length : 0;
      maxLength = Math.max(maxLength, columnLength);
    });
    column.width = 30;
  });

  const excelBuffer = await workbook.xlsx.writeBuffer();

  return excelBuffer;
};

module.exports = {
  generateGoalsWorksheetExcel,
  generateGoalsWorksheetExcelV2,
};
