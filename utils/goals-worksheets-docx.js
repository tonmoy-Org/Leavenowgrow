const docx = require("docx");
const fs = require("fs");
const { arrayToObject } = require("./arrayToObject");

const TABLE_WIDTH = 100;
const ARIAL_FAMILY = "Arial";
const CELL_COLOR = "CCFFFF";
const BORDER_STROKE = 12;
const FIRST_COLUMN_WIDTH = 1.5;
const PAGE_MARGIN = 0.5;

function selectPageSize(dataLength) {
  let PAGE_WIDTH;
  let PAGE_HEIGHT;
  let LANDSCAPE;
  let columnWidth;
  let TABLE_COLUMN_WIDTHS = [];
  // if (dataLength <= 2) {
  //   // A4 portrait page
  //   PAGE_WIDTH = 8.27;
  //   PAGE_HEIGHT = 11.69;
  //   LANDSCAPE = false;
  //   TABLE_COLUMN_WIDTHS.push(FIRST_COLUMN_WIDTH * 1440);
  //   const columns =
  //     (PAGE_WIDTH - 2 * PAGE_MARGIN - FIRST_COLUMN_WIDTH) / dataLength;
  //   for (let i = 0; i < dataLength; i++) {
  //     TABLE_COLUMN_WIDTHS.push(columns * 1440);
  //   }
  // } else if (dataLength > 2 && dataLength <= 5) {
  //   // A4 landscape page
  //   PAGE_WIDTH = 8.27;
  //   PAGE_HEIGHT = 11.69;
  //   LANDSCAPE = true;
  //   //TABLE_COLUMN_WIDTHS.push(FIRST_COLUMN_WIDTH * 1440);
  //   const columns = (PAGE_HEIGHT - 2 * PAGE_MARGIN) / (dataLength + 1);
  //   for (let i = 0; i < dataLength + 1; i++) {
  //     TABLE_COLUMN_WIDTHS.push(columns * 1440);
  //   }
  // } else {
  //   // Tabloid 11 * 17 landscape page
  //   PAGE_WIDTH = 11;
  //   PAGE_HEIGHT = 17;
  //   LANDSCAPE = true;
  //   // TABLE_COLUMN_WIDTHS.push(FIRST_COLUMN_WIDTH * 1440);
  //   const columns = (PAGE_HEIGHT - 2 * PAGE_MARGIN) / (dataLength + 1);
  //   for (let i = 0; i < dataLength + 1; i++) {
  //     TABLE_COLUMN_WIDTHS.push(columns * 1440);
  //   }
  // }

  if (dataLength === 1) {
    // A4 portrait page
    PAGE_WIDTH = 8.5;
    PAGE_HEIGHT = 11;
    LANDSCAPE = false;
    TABLE_COLUMN_WIDTHS.push(FIRST_COLUMN_WIDTH * 1440);
    columnWidth =
      (PAGE_WIDTH - 2 * PAGE_MARGIN - FIRST_COLUMN_WIDTH) / dataLength;
    for (let i = 0; i < dataLength; i++) {
      TABLE_COLUMN_WIDTHS.push(columnWidth * 1440);
    }
  } else {
    // Tabloid 11 * 17 landscape page
    PAGE_WIDTH = 11;
    PAGE_HEIGHT = 17;
    LANDSCAPE = true;
    // TABLE_COLUMN_WIDTHS.push(FIRST_COLUMN_WIDTH * 1440);
    columnWidth = (PAGE_HEIGHT - 2 * PAGE_MARGIN) / (dataLength + 1);
    for (let i = 0; i < dataLength + 1; i++) {
      TABLE_COLUMN_WIDTHS.push(columnWidth * 1440);
    }
  }

  return {
    PAGE_WIDTH,
    PAGE_HEIGHT,
    LANDSCAPE,
    TABLE_COLUMN_WIDTHS,
    columnWidth,
  };
}

const blueShade = {
  fill: CELL_COLOR,
};
const borders = {
  top: {
    style: docx.BorderStyle.SINGLE,
    size: BORDER_STROKE,
    color: "000000",
  },
  bottom: {
    style: docx.BorderStyle.SINGLE,
    size: BORDER_STROKE,
    color: "000000",
  },
  left: {
    style: docx.BorderStyle.SINGLE,
    size: BORDER_STROKE,
    color: "000000",
  },
  right: {
    style: docx.BorderStyle.SINGLE,
    size: BORDER_STROKE,
    color: "000000",
  },
};
const marginStyle = {
  top: docx.convertInchesToTwip(0.1),
  bottom: docx.convertInchesToTwip(0.1),
  right: docx.convertInchesToTwip(0.0),
  left: docx.convertInchesToTwip(0.0),
};

function createRow(
  data,
  keyName,
  rowName,
  shadingProperty,
  isHeader = false,
  description = "",
  cellWidth
) {
  const nameCells = data.map((d) => {
    if (!d[keyName]) {
      return new docx.TableCell({
        width: {
          size: docx.convertInchesToTwip(cellWidth),
          type: docx.WidthType.DXA,
        },
        shading: rowName === "Name" ? shadingProperty : undefined,
        borders: borders,
        margins: marginStyle,
        children: "",
      });
    }

    const polishedItem = d[keyName]
      .replace(/\t/g, "")
      .split("\r\n")
      .map((subItem) => {
        if (subItem.length > 0) {
          return new docx.Paragraph({
            indent: {
              left: "0.25in",
              right: 0,
            },
            spacing: {
              beforeAutoSpacing: true,
              after: 200,
            },
            children: [
              new docx.TextRun({
                text: subItem,
                font: ARIAL_FAMILY,
                bold: keyName === "name",
                size: 20,
              }),
            ],
          });
        }
      });

    return new docx.TableCell({
      width: {
        size: docx.convertInchesToTwip(cellWidth),
        type: docx.WidthType.DXA,
      },
      shading: rowName === "Name" ? shadingProperty : undefined,
      borders: borders,
      margins: marginStyle,
      children: polishedItem,
    });
  });

  const tableRow = new docx.TableRow({
    tableHeader: isHeader,
    cantSplit: false,
    children: [
      new docx.TableCell({
        width: {
          size: docx.convertInchesToTwip(FIRST_COLUMN_WIDTH),
          type: docx.WidthType.DXA,
        },
        margins: marginStyle,
        borders: borders,
        shading: shadingProperty,
        children: [
          new docx.Paragraph({
            indent: {
              left: "0.25in",
              right: 0,
            },
            spacing: {
              beforeAutoSpacing: true,
              afterAutoSpacing: true,
            },
            children: [
              new docx.TextRun({
                text: rowName,
                font: ARIAL_FAMILY,
                size: 20,
                bold: true,
              }),
            ],
          }),
          new docx.Paragraph({
            indent: {
              left: "0.25in",
              right: 0,
            },
            spacing: {
              beforeAutoSpacing: true,
              afterAutoSpacing: true,
            },
            children: [
              new docx.TextRun({
                text: description,
                font: ARIAL_FAMILY,
                italics: true,
                size: 14,
              }),
            ],
          }),
        ],
      }),
      ...nameCells,
    ],
  });

  return tableRow;
}

// generate docx file from employee data
async function generateGoalsWorksheetDocx(goalsWorksheets, companyName) {
  const {
    PAGE_WIDTH,
    PAGE_HEIGHT,
    LANDSCAPE,
    columnWidth,
    TABLE_COLUMN_WIDTHS,
  } = selectPageSize(goalsWorksheets.length);
  let fullTable = [];

  const worksheetsData = goalsWorksheets.map((item) => {
    const worksheets = arrayToObject(item?.worksheets);

    return { ...worksheets, name: item?.name };
  });

  const nameRow = createRow(
    worksheetsData,
    "name",
    "Name",
    blueShade,
    true,
    "",
    columnWidth
  );
  const goalRow1 = createRow(
    worksheetsData,
    "goal0",
    "Goal #1",
    blueShade,
    false,
    "",
    columnWidth
  );
  const benefitRow1 = createRow(
    worksheetsData,
    "benefit0",
    "Benefit #1",
    blueShade,
    false,
    "",
    columnWidth
  );
  const metricRow1 = createRow(
    worksheetsData,
    "metric0",
    "Metric #1",
    blueShade,
    false,
    "",
    columnWidth
  );
  const actionRow1 = createRow(
    worksheetsData,
    "action0",
    "Action #1",
    blueShade,
    false,
    "",
    columnWidth
  );
  const goalRow2 = createRow(
    worksheetsData,
    "goal1",
    "Goal #2",
    blueShade,
    false,
    "",
    columnWidth
  );
  const benefitRow2 = createRow(
    worksheetsData,
    "benefit1",
    "Benefit #2",
    blueShade,
    false,
    "",
    columnWidth
  );
  const metricRow2 = createRow(
    worksheetsData,
    "metric1",
    "Metric #2",
    blueShade,
    false,
    "",
    columnWidth
  );
  const actionRow2 = createRow(
    worksheetsData,
    "action1",
    "Action #2",
    blueShade,
    false,
    "",
    columnWidth
  );
  const goalRow3 = createRow(
    worksheetsData,
    "goal2",
    "Goal #3",
    blueShade,
    false,
    "",
    columnWidth
  );
  const benefitRow3 = createRow(
    worksheetsData,
    "benefit2",
    "Benefit #3",
    blueShade,
    false,
    "",
    columnWidth
  );
  const metricRow3 = createRow(
    worksheetsData,
    "metric2",
    "Metric #3",
    blueShade,
    false,
    "",
    columnWidth
  );
  const actionRow3 = createRow(
    worksheetsData,
    "action2",
    "Action #3",
    blueShade,
    false,
    "",
    columnWidth
  );
  const goalRow4 = createRow(
    worksheetsData,
    "goal3",
    "Goal #4",
    blueShade,
    false,
    "",
    columnWidth
  );
  const benefitRow4 = createRow(
    worksheetsData,
    "benefit3",
    "Benefit #4",
    blueShade,
    false,
    "",
    columnWidth
  );
  const metricRow4 = createRow(
    worksheetsData,
    "metric3",
    "Metric #4",
    blueShade,
    false,
    "",
    columnWidth
  );
  const actionRow4 = createRow(
    worksheetsData,
    "action3",
    "Action #4",
    blueShade,
    false,
    "",
    columnWidth
  );

  fullTable = [
    nameRow,
    goalRow1,
    benefitRow1,
    metricRow1,
    actionRow1,
    goalRow2,
    benefitRow2,
    metricRow2,
    actionRow2,
    goalRow3,
    benefitRow3,
    metricRow3,
    actionRow3,
    goalRow4,
    benefitRow4,
    metricRow4,
    actionRow4,
  ];

  const table = new docx.Table({
    width: {
      size: 100,
      type: docx.WidthType.PERCENTAGE,
    },
    columnWidths: TABLE_COLUMN_WIDTHS,
    rows: fullTable,
  });

  const heading = new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: {
      after: 0,
    },
    children: [
      new docx.TextRun({
        text: "Goals Worksheet",
        font: ARIAL_FAMILY,
        size: 14,
      }),
    ],
  });

  if (companyName?.length === 0) {
    companyName = "";
  }

  const subHeading = new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: {
      after: 500,
    },
    children: [
      new docx.TextRun({
        text: `${companyName}`,
        font: ARIAL_FAMILY,
        size: 21,
      }),
    ],
  });

  const logo = new docx.Paragraph({
    children: [
      new docx.ImageRun({
        data: fs.readFileSync("./public/img/photos/symbol.png"),
        transformation: {
          width: 50,
          height: 50,
        },
        floating: {
          horizontalPosition: {
            relative: docx.HorizontalPositionRelativeFrom.PAGE,
            align: docx.HorizontalPositionAlign.RIGHT,
          },
          verticalPosition: {
            relative: docx.VerticalPositionRelativeFrom.PAGE,
            align: docx.VerticalPositionAlign.BOTTOM,
          },
        },
      }),
    ],
  });

  const doc = new docx.Document({
    creator: "LeaveNowGrow",
    description: "This is the goals worksheet of employees.",
    title: "LeaveNowGrow",
    background: {
      color: "FFFFFF",
    },
    sections: [
      {
        properties: {
          page: {
            size: {
              orientation:
                LANDSCAPE === true
                  ? docx.PageOrientation.LANDSCAPE
                  : docx.PageOrientation.PORTRAIT,
              width: docx.convertInchesToTwip(PAGE_WIDTH),
              height: docx.convertInchesToTwip(PAGE_HEIGHT),
            },
            margin: {
              top: docx.convertInchesToTwip(PAGE_MARGIN),
              bottom: docx.convertInchesToTwip(PAGE_MARGIN),
              right: docx.convertInchesToTwip(PAGE_MARGIN),
              left: docx.convertInchesToTwip(PAGE_MARGIN),
            },
          },
        },
        headers: {
          default: new docx.Header({
            children: [],
          }),
        },
        footers: {
          default: new docx.Footer({
            children: [logo],
          }),
        },
        children: [heading, subHeading, table],
      },
    ],
  });
  const docxBuffer = await docx.Packer.toBuffer(doc);
  // const docxBuffer = await docx.Packer.toBase64String(doc);

  return docxBuffer;
}

module.exports = { generateGoalsWorksheetDocx };
