const path = require("path");
const ejs = require("ejs");
const excel = require("exceljs");
const separateTagValues = require("./separateTagValues");

const generateExcelFormdata = async (formDataItems, cName) => {
  const templatePathDocx = path.join(
    __dirname,
    "../views",
    "employees-docx.ejs"
  );

  let htmlPageDocx = "";
  htmlPageDocx = await ejs.renderFile(templatePathDocx, {
    formDataItems: formDataItems,
    print: true,
    company: cName,
  });

  let tValues = separateTagValues(htmlPageDocx);

  const workbook = new excel.Workbook();
  const worksheet = workbook.addWorksheet("Sheet 1");

  const childArrays = formDataItems.map((doc) => doc.toObject());

  worksheet.getColumn(1).alignment = { wrapText: true, vertical: "top" };

  worksheet.getColumn(1).font = { bold: true };
  worksheet.getColumn(1).values = tValues["thValues"];

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
      obj.title,
      obj.jobDescription,
      obj.managerSupervisor,
      obj.directReports,
      obj.accountable,
      obj.participate,
      obj.metrics,
      obj.positionalObjectives,
      obj.personalObjectives,
      obj.authorityLevels,
      obj.delegationOfAuthority,
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

const generateAllExcel = async (formDataItems) => {
  const templatePath = path.join(
    __dirname,
    "../views",
    "all-evaluations-pdf.ejs"
  );

  const htmlPage = await ejs.renderFile(templatePath, {
    formDataItems: formDataItems,
    print: true,
  });

  let tValues = separateTagValues(htmlPage);

  const workbook = new excel.Workbook();
  const worksheet = workbook.addWorksheet("Sheet 1");

  worksheet.getColumn(1).values = tValues["thValues"];

  let $a = 2;

  formDataItems.forEach((obj) => {
    const data = [
      obj?.evaluation,
      obj?.title,
      obj?.jobDescription,
      obj?.managerSupervisor,
      obj?.directReports,
      obj?.accountable,
      obj?.participate,
      obj?.metrics,
      obj?.positionalObjectives,
      obj?.personalObjectives,
      obj?.authorityLevels,
      obj?.delegationOfAuthority,
      obj?.managerNotes,
      obj?.managerReview,
    ];

    worksheet.getColumn($a).alignment = { wrapText: true, vertical: "top" };
    worksheet.getColumn($a).values = data;
    $a++;
  });

  worksheet.getColumn(1).eachCell((cell, rowNumber) => {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "CCFFFF" },
    };
    cell.alignment = { wrapText: true, vertical: "top" };
    cell.font = { name: "Arial", bold: true };
  });

  worksheet.getRow(1).eachCell((cell, colNumber) => {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "CCFFFF" },
    };
    cell.font = { name: "Arial", bold: true };
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
    column.eachCell({ includeEmpty: true }, (cell) => {
      cell.value = cell.value ? cell.value.trim() : "";
    });
    column.width = 75;
  });

  const excelBuffer = await workbook.xlsx.writeBuffer();

  return excelBuffer;
};

const generateSelfExcel = async (formDataItems) => {
  const templatePathDocx = path.join(
    __dirname,
    "../views",
    "employees-docx.ejs"
  );

  const htmlPageDocx = await ejs.renderFile(templatePathDocx, {
    formDataItems: formDataItems,
    print: true,
    company: "",
    manager: false,
    quarter: formDataItems?.quarter,
    year: formDataItems?.year,
  });

  let tValues = separateTagValues(htmlPageDocx);

  const workbook = new excel.Workbook();
  const worksheet = workbook.addWorksheet("Sheet 1");

  const childArrays = formDataItems.map((doc) => doc.toObject());

  worksheet.getColumn(1).alignment = { wrapText: true, vertical: "top" };

  worksheet.getColumn(1).font = { bold: true };
  worksheet.getColumn(1).values = tValues["thValues"];

  let $a = 2;

  childArrays.forEach((obj) => {
    const data = [
      obj.name,
      obj.title,
      obj.jobDescription,
      obj.managerSupervisor,
      obj.directReports,
      obj.accountable,
      obj.participate,
      obj.metrics,
      obj.positionalObjectives,
      obj.personalObjectives,
      obj.authorityLevels,
      obj.delegationOfAuthority,
      obj?.managerReview,
    ];

    worksheet.getColumn($a).alignment = { wrapText: true, vertical: "top" };
    worksheet.getColumn($a).values = data;
    $a++;
  });

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

const generateManagerExcel = async (formDataItems) => {
  const templatePathDocx = path.join(
    __dirname,
    "../views",
    "evaluation-form-pdf.ejs"
  );

  const htmlPageDocx = await ejs.renderFile(templatePathDocx, {
    formDataItems: formDataItems,
    print: true,
    company: "",
    manager: true,
    quarter: formDataItems?.quarter,
    year: formDataItems?.year,
  });

  let tValues = separateTagValues(htmlPageDocx);

  const workbook = new excel.Workbook();
  const worksheet = workbook.addWorksheet("Sheet 1");

  const childArrays = formDataItems.map((doc) => doc.toObject());

  worksheet.getColumn(1).alignment = { wrapText: true, vertical: "top" };

  worksheet.getColumn(1).font = { bold: true };
  worksheet.getColumn(1).values = tValues["thValues"];

  let $a = 2;

  childArrays.forEach((obj) => {
    const data = [
      obj.name,
      obj.title,
      obj.jobDescription,
      obj.managerSupervisor,
      obj.directReports,
      obj.accountable,
      obj.participate,
      obj.metrics,
      obj.positionalObjectives,
      obj.personalObjectives,
      obj.authorityLevels,
      obj.delegationOfAuthority,
      obj.managerNotes,
      obj?.managerReview,
    ];

    worksheet.getColumn($a).alignment = { wrapText: true, vertical: "top" };
    worksheet.getColumn($a).values = data;
    $a++;
  });

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

module.exports = {
  generateExcelFormdata,
  generateAllExcel,
  generateSelfExcel,
  generateManagerExcel,
};
