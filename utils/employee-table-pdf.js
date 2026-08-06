const ejs = require("ejs");
const fs = require("fs");
const launchPuppeteer = require("./launchPuppeteer");
const { consoleMe } = require("./consoleMe");

async function generatePDF(formDataItems, cName, templatePath) {
  try {
    const htmlPage = await ejs.renderFile(templatePath, {
      formDataItems: formDataItems,
      print: true,
      cName: cName,
    });

    const browser = await launchPuppeteer();
    const page = await browser.newPage();

    await page.setContent(htmlPage, { waitUntil: "networkidle0" });

    const headerTemplate = `
      <div style="width: 100%; text-align: center; margin: 0px;">
        <h3>
          <span style="font-size: 9px; font-weight: normal">
            ASSIGNMENT OF ACCOUNTABILITY
          </span><br>
          <span style="font-size: 14px; font-weight: 300">
            ${cName}
          </span>
        </h3>
      </div>`;

    const footerTemplate = `
      <div style="width: 100%; display: flex; justify-content: flex-end; align-items: flex-end">
        <img src="data:image/png;base64,${fs
          .readFileSync("./public/img/photos/symbol.png")
          .toString("base64")}" width="50" height="50" style="display: block">
      </div>`;

    const pdfGenerated = await page.pdf({
      format: "TABLOID",
      landscape: true,
      printBackground: true,
      displayHeaderFooter: true,
      headerTemplate,
      footerTemplate,
      margin: {
        top: "2cm",
        bottom: "2cm",
        left: "2cm",
        right: "2cm",
      },
    });

    const pdfBuffer = Buffer.from(pdfGenerated);

    await browser.close();

    return pdfBuffer;
  } catch (error) {
    consoleMe(error);
  }
}

async function generatePDFonA4(formDataItems, cName, templatePath) {
  try {
    const htmlPage = await ejs.renderFile(templatePath, {
      formDataItems: formDataItems,
      print: true,
      cName: cName,
    });

    const browser = await launchPuppeteer();
    const page = await browser.newPage();

    await page.setContent(htmlPage, { waitUntil: "networkidle0" });

    const headerTemplate = `
      <div style="width: 100%; text-align: center; margin: 0px;">
        <h3>
          <span style="font-size: 9px; font-weight: normal">
            ASSIGNMENT OF ACCOUNTABILITY
          </span><br>
          <span style="font-size: 14px; font-weight: 300">
            ${cName}
          </span>
        </h3>
      </div>`;

    const footerTemplate = `
      <div style="width: 100%; display: flex; justify-content: flex-end; align-items: flex-end">
        <img src="data:image/png;base64,${fs
          .readFileSync("./public/img/photos/symbol.png")
          .toString("base64")}" width="50" height="50" style="display: block">
      </div>`;

    const pdfGenerated = await page.pdf({
      format: "A4", // Use A4 paper size
      printBackground: true,
      displayHeaderFooter: true,
      headerTemplate,
      footerTemplate,
      margin: {
        top: "2cm", // Increase top margin
        bottom: "2cm", // Increase bottom margin
        left: "2cm", // Increase left margin
        right: "2cm", // Increase right margin
      },
    });

    const pdfBuffer = Buffer.from(pdfGenerated);

    await browser.close();

    return pdfBuffer;
  } catch (error) {
    consoleMe(error);
  }
}

async function generatePDFonTabloid(formDataItems, cName, templatePath) {
  try {
    const htmlPage = await ejs.renderFile(templatePath, {
      formDataItems: formDataItems,
      print: true,
      company: cName,
    });

    const browser = await launchPuppeteer();
    const page = await browser.newPage();

    await page.setContent(htmlPage, { waitUntil: "networkidle0" });

    const headerTemplate = `
      <div style="width: 100%; text-align: center; margin: 0px;">
        <h3>
          <span style="font-size: 9px; font-weight: normal">
            ASSIGNMENT OF ACCOUNTABILITY
          </span><br>
          <span style="font-size: 14px; font-weight: 300">
            ${cName}
          </span>
        </h3>
      </div>`;

    const footerTemplate = `
      <div style="width: 100%; display: flex; justify-content: flex-end; align-items: flex-end">
        <img src="data:image/png;base64,${fs
          .readFileSync("./public/img/photos/symbol.png")
          .toString("base64")}" width="50" height="50" style="display: block">
      </div>`;

    const pdfGenerated = await page.pdf({
      format: "TABLOID", // Use Tabloid paper size
      landscape: true,
      printBackground: true,
      displayHeaderFooter: true,
      headerTemplate,
      footerTemplate,
      margin: {
        top: "2cm", // Increase top margin
        bottom: "2cm", // Increase bottom margin
        left: "2cm", // Increase left margin
        right: "2cm", // Increase right margin
      },
    });

    const pdfBuffer = Buffer.from(pdfGenerated);

    await browser.close();

    return pdfBuffer;
  } catch (error) {
    consoleMe(error);
  }
}

async function generateEvaluationsPDF(formDataItems, manager) {
  try {
    const templatePath = path.join(
      __dirname,
      "../views",
      "evaluation-form-pdf.ejs"
    );

    const htmlPage = await ejs.renderFile(templatePath, {
      formDataItems: formDataItems,
      print: true,
      company: "",
      manager: manager,
      quarter: selfDataItem?.quarter,
      year: selfDataItem?.year,
    });

    const browser = await launchPuppeteer();
    const page = await browser.newPage();

    await page.setContent(htmlPage, { waitUntil: "networkidle0" });

    const headerTemplate = `
      <div style="width: 100%; text-align: center; margin: 0px;">
        <h3>
          <span style="font-size: 9px; font-weight: normal">
            ASSIGNMENT OF ACCOUNTABILITY
          </span><br>
        </h3>
      </div>`;

    const footerTemplate = `
      <div style="width: 100%; display: flex; justify-content: flex-end; align-items: flex-end">
        <img src="data:image/png;base64,${fs
          .readFileSync("./public/img/photos/symbol.png")
          .toString("base64")}" width="50" height="50" style="display: block">
      </div>`;

    const pdfGenerated = await page.pdf({
      format: "TABLOID", // Use Tabloid paper size
      landscape: true,
      printBackground: true,
      displayHeaderFooter: true,
      headerTemplate,
      footerTemplate,
      margin: {
        top: "2cm", // Increase top margin
        bottom: "2cm", // Increase bottom margin
        left: "2cm", // Increase left margin
        right: "2cm", // Increase right margin
      },
    });

    const pdfBuffer = Buffer.from(pdfGenerated);

    await browser.close();

    return pdfBuffer;
  } catch (error) {
    consoleMe(error);
  }
}

module.exports = {
  generatePDF,
  generatePDFonA4,
  generatePDFonTabloid,
  generateEvaluationsPDF,
};
