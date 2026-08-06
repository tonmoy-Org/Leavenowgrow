const jwt = require("jsonwebtoken");
const { configEnv } = require("../config");

const getToken = (data) => {
  const token = jwt.sign(data, configEnv.jwtSecret, {
    expiresIn: configEnv.jwtExpires,
  });

  return token;
};

module.exports = { getToken };
