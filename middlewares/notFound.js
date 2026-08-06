const notFound = (req, res, next) => {
  res.render("404", { content: "Requested path not found." });
  // res.status(403).json({
  //   success: false,
  //   message: "Requested path not found.",
  //   error: "Not found error.",
  // });
};

module.exports = { notFound };
