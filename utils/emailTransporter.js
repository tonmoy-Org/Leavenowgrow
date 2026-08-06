const nodemailer = require("nodemailer");
const path = require("path");
const { ConfidentialClientApplication } = require("@azure/msal-node");

require("dotenv").config({ path: path.resolve(process.cwd(), "../.env") });

const config = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
};

const cca = new ConfidentialClientApplication(config);

const getAccessToken = async () => {
  const result = await cca.acquireTokenByClientCredential({
    scopes: ["https://outlook.office365.com/.default"],
  });
  return result.accessToken;
};

const createEmailTransporter = async () => {
  const accessToken = await getAccessToken();

  const transporter = nodemailer.createTransport({
    service: "Outlook365",
    auth: {
      type: "OAuth2",
      user: process.env.USER_EMAIL,
      accessToken: accessToken,
    },
  });

  return transporter;
  // Mail options should be like this
  //   const mailOptions = {
  //     from: `LeaveNowGrow <${process.env.USER_EMAIL}>`,
  //     to: "miraj2465@gmail.com",
  //     subject: "Test Email",
  //     text: "Hello, this is a test email sent using Nodemailer and OAuth2!",
  //   };

  //   transporter.sendMail(mailOptions, (error, info) => {
  //     if (error) {
  //       return console.log(error);
  //     }
  //     console.log("Email sent: " + info.response);
  //   });
};

module.exports = { createEmailTransporter };
