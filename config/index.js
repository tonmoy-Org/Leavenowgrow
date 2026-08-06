const path = require("path");

require("dotenv").config({ path: path.join(process.cwd(), ".env") });

const configEnv = {
  tenantId: process.env.TENANT_ID,
  clientId: process.env.CLIENT_ID,
  userEmail: process.env.USER_EMAIL,
  clientSecret: process.env.CLIENT_SECRET,
  aadEndpoint: process.env.AAD_ENDPOINT,
  graphEndpoint: process.env.GRAPH_ENDPOINT,
  mode: process.env.MODE,
  port: process.env.PORT,
  mongodbUri: process.env.MONGODB_URI,
  prodUrl: process.env.PROD_URL,
  secretKey: process.env.SECRET_KEY,
  jwtSecret: process.env.JWT_SECRET,
  jwtExpires: process.env.JWT_EXPIRES,
  defaultPass: process.env.DEFAULT_PASS,
};

module.exports = { configEnv };
