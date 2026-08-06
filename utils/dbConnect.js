const mongoose = require("mongoose");
const { consoleMe } = require("./consoleMe");

// Connect to MongoDB using Mongoose with optimized connection pool settings
const dbConnect = (mongoURL) => {
  mongoose
    .connect(mongoURL, {
      useNewUrlParser: true,
      useUnifiedTopology: true,
      useBigInt64: false,
      maxPoolSize: 50,
      minPoolSize: 5,
      serverSelectionTimeoutMS: 5000,
      socketTimeoutMS: 45000,
    })
    .then(() => {
      consoleMe("Connected to MongoDB");
    })
    .catch((error) => {
      consoleMe("Failed to connect to MongoDB:", error);
    });
};

module.exports = { dbConnect };
