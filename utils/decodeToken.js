const jwt = require("jsonwebtoken");
const { configEnv } = require("../config");

const decodeToken = (token) => {
  let decodedData;

  jwt.verify(token, configEnv.jwtSecret, function (err, decoded) {
    if (err) {
      throw new Error("Your token is invalid or expired.", { cause: err });
    }
    decodedData = decoded;
  });

  return decodedData;
};

module.exports = { decodeToken };
