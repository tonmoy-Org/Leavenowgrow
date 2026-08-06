const nodemailer = require("nodemailer");
const path = require("path");
const { ConfidentialClientApplication } = require("@azure/msal-node");
const { consoleMe } = require("./consoleMe");

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

const sendEmail = async () => {
  const accessToken = await getAccessToken();

  const transporter = nodemailer.createTransport({
    service: "Outlook365",
    auth: {
      type: "OAuth2",
      user: process.env.USER_EMAIL,
      accessToken: accessToken,
    },
  });

  const mailOptions = {
    from: `LeaveNowGrow <${process.env.USER_EMAIL}>`,
    to: "miraj2465@gmail.com",
    subject: "Test Email",
    text: "Hello, this is a test email sent using Nodemailer and OAuth2!",
  };

    transporter.sendMail(mailOptions, (error, info) => {
      if (error) {
        return consoleMe(error);
      }
      consoleMe("Email sent: " + info.response);
    });

  // verify connection configuration
  // transporter.verify(function (error, success) {
  //   if (error) {
  //     console.log(error);
  //   } else {
  //     console.log("Server is ready to take our messages");
  //   }
  // });
};

sendEmail().catch(console.error);

// testing the web api, getting sign in page
// const axios = require("axios");

// async function callApi(endpoint) {
//   const accessToken = await getAccessToken();

//   const options = {
//     headers: {
//       Authorization: `Bearer ${accessToken}`,
//     },
//   };

//   console.log("request made to web API at: " + new Date().toString());

//   try {
//     const response = await axios.default.get(endpoint, options);
//     console.log(response.data);
//     return response.data;
//   } catch (error) {
//     console.log(error);
//     return error;
//   }
// }

// callApi("https://outlook.office365.com/.default");
