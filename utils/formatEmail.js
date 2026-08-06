const formatEmail = (content = "") => {
  const contentArray = content.split("\n");

  let htmlContent = ["<div>"];
  contentArray.forEach((string) => {
    htmlContent.push(`<p>${string}</p>`);
  });
  htmlContent.push("</div>");

  const htmlString = htmlContent.join("");

  return htmlString;
};

module.exports = { formatEmail };
