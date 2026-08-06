const docx = require("docx");
const fs = require("fs");

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
        shading:
          rowName === "Name" || rowName === "Evaluation"
            ? shadingProperty
            : undefined,
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
      shading:
        rowName === "Name" || rowName === "Evaluation"
          ? shadingProperty
          : undefined,
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
async function generate(employees, companyName) {
  const {
    PAGE_WIDTH,
    PAGE_HEIGHT,
    LANDSCAPE,
    columnWidth,
    TABLE_COLUMN_WIDTHS,
  } = selectPageSize(employees.length);
  let fullTable = [];

  const nameRow = createRow(
    employees,
    "name",
    "Name",
    blueShade,
    true,
    "First and last name.",
    columnWidth
  );
  const titleRow = createRow(
    employees,
    "title",
    "Title",
    blueShade,
    false,
    "EXACT current job title.",
    columnWidth
  );
  const jobDescriptionRow = createRow(
    employees,
    "jobDescription",
    "Job Description",
    blueShade,
    false,
    "Describe your key job responsibilities in your current position. (3-5 bulleted statements, ranked in order of importance).",
    columnWidth
  );
  const managerSupervisorRow = createRow(
    employees,
    "managerSupervisor",
    "Manager/ Supervisor",
    blueShade,
    false,
    "Name and EXACT title of the one person to whom you directly report; Who is responsible for your formal performance reviews?",
    columnWidth
  );
  const directReportsRow = createRow(
    employees,
    "directReports",
    "Direct Reports",
    blueShade,
    false,
    "The names and EXACT titles of those persons who report directly to you; you are responsible for their formal performance reviews.",
    columnWidth
  );
  const accountableRow = createRow(
    employees,
    "accountable",
    "Accountable",
    blueShade,
    false,
    "Areas of focus for which YOU, and ONLY YOU, are Accountable.",
    columnWidth
  );
  const participateRow = createRow(
    employees,
    "participate",
    "Participate",
    blueShade,
    false,
    "Activities you have responsibility to participate in but for which someone else has the 'A'",
    columnWidth
  );
  const compensationRow = createRow(
    employees,
    "compensation",
    "Compensation",
    blueShade,
    false,
    "How are you compensated, other than salary (types and structure of compensation, not amounts: e.g., commissions, bonuses, gain-sharing, ownership interest)?",
    columnWidth
  );
  const metricsRow = createRow(
    employees,
    "metrics",
    "Metrics",
    blueShade,
    false,
    "How does your direct supervisor measure your success?",
    columnWidth
  );
  const positionalObjectivesRow = createRow(
    employees,
    "positionalObjectives",
    "Positional Objectives",
    blueShade,
    false,
    "What are your key objectives for this fiscal year in your current position? These are likely objectives that are currently in written form.",
    columnWidth
  );
  const personalObjectivesRow = createRow(
    employees,
    "personalObjectives",
    "Personal Objectives",
    blueShade,
    false,
    "What are your personal (not positional) objectives over the next 2-3 years and beyond?",
    columnWidth
  );
  const authorityLevelsRow = createRow(
    employees,
    "authorityLevels",
    "Authority Levels",
    blueShade,
    false,
    "What authority levels or decision rights do you have for planning, spending, hiring/firing, contractual relationships, structural changes, etc.?",
    columnWidth
  );
  const delegationOfAuthorityRow = createRow(
    employees,
    "delegationOfAuthority",
    "Delegation Of Authority",
    blueShade,
    false,
    "What freedom do you give to your subordinates to make decisions on their own, e.g. authority levels, requesting permission, submitting data, etc. Think about which of your accountabilities you might be able to delegate.",
    columnWidth
  );

  fullTable = [
    nameRow,
    titleRow,
    jobDescriptionRow,
    managerSupervisorRow,
    directReportsRow,
    accountableRow,
    participateRow,
    metricsRow,
    positionalObjectivesRow,
    personalObjectivesRow,
    authorityLevelsRow,
    delegationOfAuthorityRow,
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
        text: "ASSIGNMENT OF ACCOUNTABILITY",
        font: ARIAL_FAMILY,
        size: 14,
      }),
    ],
  });

  if (companyName.length === 0) {
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
    description: "This is an AOA report of employees.",
    title: "Leave Now Grow",
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

// generate docx file from self evaluation data
async function generateSelf(employees) {
  const {
    PAGE_WIDTH,
    PAGE_HEIGHT,
    LANDSCAPE,
    columnWidth,
    TABLE_COLUMN_WIDTHS,
  } = selectPageSize(employees.length);
  let fullTable = [];

  const nameRow = createRow(
    employees,
    "name",
    "Name",
    blueShade,
    true,
    "First and last name.",
    columnWidth
  );
  const titleRow = createRow(
    employees,
    "title",
    "Title",
    blueShade,
    false,
    "EXACT current job title.",
    columnWidth
  );
  const jobDescriptionRow = createRow(
    employees,
    "jobDescription",
    "Job Description",
    blueShade,
    false,
    "Describe your key job responsibilities in your current position. (3-5 bulleted statements, ranked in order of importance).",
    columnWidth
  );
  const managerSupervisorRow = createRow(
    employees,
    "managerSupervisor",
    "Manager/ Supervisor",
    blueShade,
    false,
    "Name and EXACT title of the one person to whom you directly report; Who is responsible for your formal performance reviews?",
    columnWidth
  );
  const directReportsRow = createRow(
    employees,
    "directReports",
    "Direct Reports",
    blueShade,
    false,
    "The names and EXACT titles of those persons who report directly to you; you are responsible for their formal performance reviews.",
    columnWidth
  );
  const accountableRow = createRow(
    employees,
    "accountable",
    "Accountable",
    blueShade,
    false,
    "Areas of focus for which YOU, and ONLY YOU, are Accountable.",
    columnWidth
  );
  const participateRow = createRow(
    employees,
    "participate",
    "Participate",
    blueShade,
    false,
    "Activities you have responsibility to participate in but for which someone else has the 'A'",
    columnWidth
  );
  const compensationRow = createRow(
    employees,
    "compensation",
    "Compensation",
    blueShade,
    false,
    "How are you compensated, other than salary (types and structure of compensation, not amounts: e.g., commissions, bonuses, gain-sharing, ownership interest)?",
    columnWidth
  );
  const metricsRow = createRow(
    employees,
    "metrics",
    "Metrics",
    blueShade,
    false,
    "How does your direct supervisor measure your success?",
    columnWidth
  );
  const positionalObjectivesRow = createRow(
    employees,
    "positionalObjectives",
    "Positional Objectives",
    blueShade,
    false,
    "What are your key objectives for this fiscal year in your current position? These are likely objectives that are currently in written form.",
    columnWidth
  );
  const personalObjectivesRow = createRow(
    employees,
    "personalObjectives",
    "Personal Objectives",
    blueShade,
    false,
    "What are your personal (not positional) objectives over the next 2-3 years and beyond?",
    columnWidth
  );
  const authorityLevelsRow = createRow(
    employees,
    "authorityLevels",
    "Authority Levels",
    blueShade,
    false,
    "What authority levels or decision rights do you have for planning, spending, hiring/firing, contractual relationships, structural changes, etc.?",
    columnWidth
  );
  const delegationOfAuthorityRow = createRow(
    employees,
    "delegationOfAuthority",
    "Delegation Of Authority",
    blueShade,
    false,
    "What freedom do you give to your subordinates to make decisions on their own, e.g. authority levels, requesting permission, submitting data, etc. Think about which of your accountabilities you might be able to delegate.",
    columnWidth
  );
  const managerReviewSelfRow = createRow(
    employees,
    "managerReview",
    "Manager Review",
    blueShade,
    false,
    "2 things my manager is doing well that support my success and two areas that could be improved to better support me.",
    columnWidth
  );

  fullTable = [
    nameRow,
    titleRow,
    jobDescriptionRow,
    managerSupervisorRow,
    directReportsRow,
    accountableRow,
    participateRow,
    metricsRow,
    positionalObjectivesRow,
    personalObjectivesRow,
    authorityLevelsRow,
    delegationOfAuthorityRow,
    managerReviewSelfRow,
  ];

  const table = new docx.Table({
    width: {
      size: 100,
      type: docx.WidthType.PERCENTAGE,
    },
    columnWidths: TABLE_COLUMN_WIDTHS,
    rows: fullTable,
  });

  const heading1 = new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: {
      after: 0,
    },
    children: [
      new docx.TextRun({
        text: "ASSIGNMENT OF ACCOUNTABILITY",
        font: ARIAL_FAMILY,
        size: 14,
      }),
    ],
  });

  const heading2 = new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: {
      after: 0,
    },
    children: [
      new docx.TextRun({
        text: "Periodic Performance Discussion - Self Evaluation",
        font: ARIAL_FAMILY,
        size: 21,
      }),
    ],
  });

  const subHeading = new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: {
      after: 500,
    },
    children: [
      new docx.TextRun({
        text: `${employees[0]?.quarter}, ${employees[0]?.year}`,
        font: ARIAL_FAMILY,
        size: 14,
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
    description: "This is an AOA report of employees.",
    title: "Leave Now Grow",
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
        children: [heading1, heading2, subHeading, table],
      },
    ],
  });
  const docxBuffer = await docx.Packer.toBuffer(doc);
  // const docxBuffer = await docx.Packer.toBase64String(doc);

  return docxBuffer;
}

// generate manager docx file from manager evaluation data
async function generateManager(employees) {
  const {
    PAGE_WIDTH,
    PAGE_HEIGHT,
    LANDSCAPE,
    columnWidth,
    TABLE_COLUMN_WIDTHS,
  } = selectPageSize(employees.length);
  let fullTable = [];

  const nameRow = createRow(
    employees,
    "name",
    "Name",
    blueShade,
    true,
    "First and last name.",
    columnWidth
  );
  const titleRow = createRow(
    employees,
    "title",
    "Title",
    blueShade,
    false,
    "EXACT current job title.",
    columnWidth
  );
  const jobDescriptionRow = createRow(
    employees,
    "jobDescription",
    "Job Description",
    blueShade,
    false,
    "Describe your key job responsibilities in your current position. (3-5 bulleted statements, ranked in order of importance).",
    columnWidth
  );
  const managerSupervisorRow = createRow(
    employees,
    "managerSupervisor",
    "Manager/ Supervisor",
    blueShade,
    false,
    "Name and EXACT title of the one person to whom you directly report; Who is responsible for your formal performance reviews?",
    columnWidth
  );
  const directReportsRow = createRow(
    employees,
    "directReports",
    "Direct Reports",
    blueShade,
    false,
    "The names and EXACT titles of those persons who report directly to you; you are responsible for their formal performance reviews.",
    columnWidth
  );
  const accountableRow = createRow(
    employees,
    "accountable",
    "Accountable",
    blueShade,
    false,
    "Areas of focus for which YOU, and ONLY YOU, are Accountable.",
    columnWidth
  );
  const participateRow = createRow(
    employees,
    "participate",
    "Participate",
    blueShade,
    false,
    "Activities you have responsibility to participate in but for which someone else has the 'A'",
    columnWidth
  );
  const compensationRow = createRow(
    employees,
    "compensation",
    "Compensation",
    blueShade,
    false,
    "How are you compensated, other than salary (types and structure of compensation, not amounts: e.g., commissions, bonuses, gain-sharing, ownership interest)?",
    columnWidth
  );
  const metricsRow = createRow(
    employees,
    "metrics",
    "Metrics",
    blueShade,
    false,
    "How does your direct supervisor measure your success?",
    columnWidth
  );
  const positionalObjectivesRow = createRow(
    employees,
    "positionalObjectives",
    "Positional Objectives",
    blueShade,
    false,
    "What are your key objectives for this fiscal year in your current position? These are likely objectives that are currently in written form.",
    columnWidth
  );
  const personalObjectivesRow = createRow(
    employees,
    "personalObjectives",
    "Personal Objectives",
    blueShade,
    false,
    "What are your personal (not positional) objectives over the next 2-3 years and beyond?",
    columnWidth
  );
  const authorityLevelsRow = createRow(
    employees,
    "authorityLevels",
    "Authority Levels",
    blueShade,
    false,
    "What authority levels or decision rights do you have for planning, spending, hiring/firing, contractual relationships, structural changes, etc.?",
    columnWidth
  );
  const delegationOfAuthorityRow = createRow(
    employees,
    "delegationOfAuthority",
    "Delegation Of Authority",
    blueShade,
    false,
    "What freedom do you give to your subordinates to make decisions on their own, e.g. authority levels, requesting permission, submitting data, etc. Think about which of your accountabilities you might be able to delegate.",
    columnWidth
  );
  const managerNotesRow = createRow(
    employees,
    "managerNotes",
    "Manager Notes and Next Steps",
    blueShade,
    false,
    "",
    columnWidth
  );
  const managerReviewManagerRow = createRow(
    employees,
    "managerReview",
    "Manager Review",
    blueShade,
    false,
    "",
    columnWidth
  );

  fullTable = [
    nameRow,
    titleRow,
    jobDescriptionRow,
    managerSupervisorRow,
    directReportsRow,
    accountableRow,
    participateRow,
    metricsRow,
    positionalObjectivesRow,
    personalObjectivesRow,
    authorityLevelsRow,
    delegationOfAuthorityRow,
    managerNotesRow,
    managerReviewManagerRow,
  ];

  const table = new docx.Table({
    width: {
      size: 100,
      type: docx.WidthType.PERCENTAGE,
    },
    columnWidths: TABLE_COLUMN_WIDTHS,
    rows: fullTable,
  });

  const heading1 = new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: {
      after: 0,
    },
    children: [
      new docx.TextRun({
        text: "ASSIGNMENT OF ACCOUNTABILITY",
        font: ARIAL_FAMILY,
        size: 14,
      }),
    ],
  });

  const heading2 = new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: {
      after: 0,
    },
    children: [
      new docx.TextRun({
        text: "Periodic Performance Discussion - Manager Evaluation",
        font: ARIAL_FAMILY,
        size: 21,
      }),
    ],
  });

  const subHeading = new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: {
      after: 500,
    },
    children: [
      new docx.TextRun({
        text: `${employees[0]?.quarter}, ${employees[0]?.year}`,
        font: ARIAL_FAMILY,
        size: 14,
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
    description: "This is an AOA report of employees.",
    title: "Leave Now Grow",
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
        children: [heading1, heading2, subHeading, table],
      },
    ],
  });
  const docxBuffer = await docx.Packer.toBuffer(doc);
  // const docxBuffer = await docx.Packer.toBase64String(doc);

  return docxBuffer;
}

// generate all evaluations aoa, self and manager
async function generateAll(employees) {
  const {
    PAGE_WIDTH,
    PAGE_HEIGHT,
    LANDSCAPE,
    columnWidth,
    TABLE_COLUMN_WIDTHS,
  } = selectPageSize(employees.length);
  let fullTable = [];

  const evaluationRow = createRow(
    employees,
    "evaluation",
    "Evaluation",
    blueShade,
    true,
    "",
    columnWidth
  );
  const titleRow = createRow(
    employees,
    "title",
    "Title",
    blueShade,
    false,
    "EXACT current job title.",
    columnWidth
  );
  const jobDescriptionRow = createRow(
    employees,
    "jobDescription",
    "Job Description",
    blueShade,
    false,
    "Describe your key job responsibilities in your current position. (3-5 bulleted statements, ranked in order of importance).",
    columnWidth
  );
  const managerSupervisorRow = createRow(
    employees,
    "managerSupervisor",
    "Manager/ Supervisor",
    blueShade,
    false,
    "Name and EXACT title of the one person to whom you directly report; Who is responsible for your formal performance reviews?",
    columnWidth
  );
  const directReportsRow = createRow(
    employees,
    "directReports",
    "Direct Reports",
    blueShade,
    false,
    "The names and EXACT titles of those persons who report directly to you; you are responsible for their formal performance reviews.",
    columnWidth
  );
  const accountableRow = createRow(
    employees,
    "accountable",
    "Accountable",
    blueShade,
    false,
    "Areas of focus for which YOU, and ONLY YOU, are Accountable.",
    columnWidth
  );
  const participateRow = createRow(
    employees,
    "participate",
    "Participate",
    blueShade,
    false,
    "Activities you have responsibility to participate in but for which someone else has the 'A'",
    columnWidth
  );
  const compensationRow = createRow(
    employees,
    "compensation",
    "Compensation",
    blueShade,
    false,
    "How are you compensated, other than salary (types and structure of compensation, not amounts: e.g., commissions, bonuses, gain-sharing, ownership interest)?",
    columnWidth
  );
  const metricsRow = createRow(
    employees,
    "metrics",
    "Metrics",
    blueShade,
    false,
    "How does your direct supervisor measure your success?",
    columnWidth
  );
  const positionalObjectivesRow = createRow(
    employees,
    "positionalObjectives",
    "Positional Objectives",
    blueShade,
    false,
    "What are your key objectives for this fiscal year in your current position? These are likely objectives that are currently in written form.",
    columnWidth
  );
  const personalObjectivesRow = createRow(
    employees,
    "personalObjectives",
    "Personal Objectives",
    blueShade,
    false,
    "What are your personal (not positional) objectives over the next 2-3 years and beyond?",
    columnWidth
  );
  const authorityLevelsRow = createRow(
    employees,
    "authorityLevels",
    "Authority Levels",
    blueShade,
    false,
    "What authority levels or decision rights do you have for planning, spending, hiring/firing, contractual relationships, structural changes, etc.?",
    columnWidth
  );
  const delegationOfAuthorityRow = createRow(
    employees,
    "delegationOfAuthority",
    "Delegation Of Authority",
    blueShade,
    false,
    "What freedom do you give to your subordinates to make decisions on their own, e.g. authority levels, requesting permission, submitting data, etc. Think about which of your accountabilities you might be able to delegate.",
    columnWidth
  );
  const managerNotesRow = createRow(
    employees,
    "managerNotes",
    "Manager Notes and Next Steps",
    blueShade,
    false,
    "",
    columnWidth
  );
  const managerReviewRow = createRow(
    employees,
    "managerReview",
    "Manager Review",
    blueShade,
    false,
    "",
    columnWidth
  );

  fullTable = [
    evaluationRow,
    titleRow,
    jobDescriptionRow,
    managerSupervisorRow,
    directReportsRow,
    accountableRow,
    participateRow,
    metricsRow,
    positionalObjectivesRow,
    personalObjectivesRow,
    authorityLevelsRow,
    delegationOfAuthorityRow,
    managerNotesRow,
    managerReviewRow,
  ];

  const table = new docx.Table({
    width: {
      size: 100,
      type: docx.WidthType.PERCENTAGE,
    },
    columnWidths: TABLE_COLUMN_WIDTHS,
    rows: fullTable,
  });

  const heading1 = new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: {
      after: 250,
    },
    children: [
      new docx.TextRun({
        text: "ASSIGNMENT OF ACCOUNTABILITY",
        font: ARIAL_FAMILY,
        size: 14,
      }),
    ],
  });

  const heading2 = new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: {
      after: 250,
    },
    children: [
      new docx.TextRun({
        text: `Periodic Performance Discussion - ${employees[0]?.name}`,
        font: ARIAL_FAMILY,
        size: 21,
      }),
    ],
  });

  const heading3 = new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: {
      after: 500,
    },
    children: [
      new docx.TextRun({
        text: `${employees[1]?.quarter}, ${employees[1]?.year}`,
        font: ARIAL_FAMILY,
        size: 14,
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
    description: "This is an AOA report of employees.",
    title: "Leave Now Grow",
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
        children: [heading1, heading2, heading3, table],
      },
    ],
  });
  const docxBuffer = await docx.Packer.toBuffer(doc);
  // const docxBuffer = await docx.Packer.toBase64String(doc);

  return docxBuffer;
}

module.exports = { generate, generateSelf, generateManager, generateAll };
