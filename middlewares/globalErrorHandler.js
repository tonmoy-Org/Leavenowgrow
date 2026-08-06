const { configEnv } = require("../config");
const { handleCastError } = require("../errors/handleCastError");
const { handleDuplicateError } = require("../errors/handleDuplicateError");
const { handleValidationError } = require("../errors/handleValidationError");
const { consoleMe } = require("../utils/consoleMe");

const globalErrorHandler = (err, req, res, next) => {
  let statusCode = 500;
  let message = "Internal Server Error. Try again later.";

  let errorSources = [
    {
      path: "",
      message: "Something went wrong.",
    },
  ];

  if (err?.name === "ValidationError") {
    const simplifiedError = handleValidationError(err);
    statusCode = simplifiedError?.statusCode;
    message = simplifiedError?.message;
    errorSources = simplifiedError?.errorSources;
  } else if (err?.name === "CastError") {
    const simplifiedError = handleCastError(err);
    statusCode = simplifiedError?.statusCode;
    message = simplifiedError?.message;
    errorSources = simplifiedError?.errorSources;
  } else if (err?.code === 11000) {
    const simplifiedError = handleDuplicateError(err);
    statusCode = simplifiedError?.statusCode;
    message = simplifiedError?.message;
    errorSources = simplifiedError?.errorSources;
  } else if (
    err?.cause?.name === "TokenExpiredError" ||
    err?.cause?.name === "JsonWebTokenError" ||
    err?.cause?.name === "NotBeforeError"
  ) {
    res.render("verification-email", { verified: false });
    return;
  } else if (err instanceof Error) {
    message = err?.message;
    errorSources = [
      {
        path: "",
        message: err?.message,
      },
    ];
  }

  // consoleMe(err?.name);
  // consoleMe(typeof err);

  if (configEnv.mode === "prod") {
    res.render("500", { content: message });
  } else {
    res.status(statusCode).json({
      success: false,
      message: message,
      errorSources: errorSources,
      err: err?.stack,
    });
  }
};

module.exports = {
  globalErrorHandler,
};
