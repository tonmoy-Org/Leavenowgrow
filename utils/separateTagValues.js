const cheerio = require("cheerio");

function separateTagValues(htmlString) {
  const $ = cheerio.load(htmlString);

  const thValues = [];
  const tdValues = [];

  $("table")
    .find("tr")
    .each((rowIndex, rowElement) => {
      const thElements = $(rowElement).find("th");
      const tdElements = $(rowElement).find("td");

      thElements.each((index, element) => {
        const thValue = $(element).text().trim();
        thValues.push(thValue);
      });

      tdElements.each((index, element) => {
        const tdValue = $(element).text().trim();
        tdValues.push(tdValue);
      });
    });

  return { thValues, tdValues };
}


module.exports = separateTagValues;