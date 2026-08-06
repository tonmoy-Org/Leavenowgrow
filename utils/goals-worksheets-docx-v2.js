const docx = require("docx");
const fs = require("fs");

const TABLE_WIDTH = 100;
const ARIAL_FAMILY = "Arial";
const CELL_COLOR = "CCFFFF";
const BORDER_STROKE = 12;
const FIRST_COLUMN_WIDTH = 0.6;
const PAGE_MARGIN = 0.5;

function selectPageSize(dataLength) {
  let PAGE_WIDTH;
  let PAGE_HEIGHT;
  let LANDSCAPE;
  let columnWidth;
  let TABLE_COLUMN_WIDTHS = [];

  // A4 landscape page
  PAGE_WIDTH = 11;
  PAGE_HEIGHT = 8.5;
  LANDSCAPE = false;
  TABLE_COLUMN_WIDTHS.push(FIRST_COLUMN_WIDTH * 1440);
  columnWidth =
    (PAGE_WIDTH - 2 * PAGE_MARGIN - FIRST_COLUMN_WIDTH) / dataLength;
  for (let i = 0; i < dataLength; i++) {
    TABLE_COLUMN_WIDTHS.push(columnWidth * 1440);
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

function createCell(data, isBold, isColored, columnWidth) {
  const polishedItem = data
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
              bold: isBold,
              size: 20,
            }),
          ],
        });
      }
    });

  return new docx.TableCell({
    width: {
      size: docx.convertInchesToTwip(columnWidth),
      type: docx.WidthType.DXA,
    },
    margins: marginStyle,
    borders: borders,
    shading: isColored ? blueShade : undefined,
    children: polishedItem,
  });
}

function createHeading(content, size) {
  return new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: {
      after: 100,
    },
    children: [
      new docx.TextRun({
        text: content,
        font: ARIAL_FAMILY,
        size: size,
      }),
    ],
  });
}

// generate docx file from employee data
async function generateGoalsWorksheetDocxV2(goalsWorksheet, companyName) {
  let fullTable = [];

  const worksheetsData = goalsWorksheet?.worksheets;

  const {
    PAGE_WIDTH,
    PAGE_HEIGHT,
    LANDSCAPE,
    columnWidth,
    TABLE_COLUMN_WIDTHS,
  } = selectPageSize(worksheetsData?.length);

  const headerWidths = [0.5, 2.375, 2.375, 2.375, 2.375];
  const headerCells = ["#", "Goals", "Benefits", "Metrics", "Actions"].map(
    (header, idx) => {
      return createCell(header, true, true, headerWidths[idx]);
    }
  );

  const headerRow = new docx.TableRow({
    tableHeader: true,
    cantSplit: false,
    children: headerCells,
  });

  const worksheetRows = worksheetsData.map((data, index) => {
    const numberCell = createCell(
      (index + 1).toString(),
      false,
      false,
      headerWidths[0]
    );
    const goalCell = createCell(data?.goal, false, false, headerWidths[1]);
    const benefitCell = createCell(
      data?.benefit,
      false,
      false,
      headerWidths[2]
    );
    const metricCell = createCell(data?.metric, false, false, headerWidths[3]);
    const actionCell = createCell(data?.action, false, false, headerWidths[4]);

    return new docx.TableRow({
      children: [numberCell, goalCell, benefitCell, metricCell, actionCell],
    });
  });

  fullTable = [headerRow, ...worksheetRows];

  const table = new docx.Table({
    width: {
      size: 100,
      type: docx.WidthType.PERCENTAGE,
    },
    columnWidths: headerWidths.map((width) => width * 1440),
    rows: fullTable,
  });

  const heading = new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: {
      after: 100,
    },
    children: [
      new docx.TextRun({
        text: "Goals Worksheet",
        font: ARIAL_FAMILY,
        size: 30,
      }),
    ],
  });

  if (companyName?.length === 0) {
    companyName = "";
  }

  const subHeading1 = createHeading(companyName, 24);
  const subHeading2 = createHeading(goalsWorksheet?.name, 20);
  const subHeading3 = createHeading(
    `${goalsWorksheet?.quarter}, ${goalsWorksheet?.year}`,
    20
  );

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
        children: [heading, subHeading1, subHeading2, subHeading3, table],
      },
    ],
  });
  const docxBuffer = await docx.Packer.toBuffer(doc);
  // const docxBuffer = await docx.Packer.toBase64String(doc);

  return docxBuffer;
}

module.exports = { generateGoalsWorksheetDocxV2 };
