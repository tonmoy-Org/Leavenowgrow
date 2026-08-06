const nodemailer = require("nodemailer");
const { formatEmail } = require("./formatEmail");
const { createEmailTransporter } = require("./emailTransporter");
const path = require("path");
require("dotenv").config({ path: path.resolve(process.cwd(), "../.env") });

const sendEmailAOALinkFromMS365 = async (emails, subject, content) => {
  const message = content;

  // Send the PDF to the provided email address
  const transporter = await createEmailTransporter();

  const mailOptions = {
    from: `LeaveNowGrow <${process.env.USER_EMAIL}>`,
    to: emails,
    subject: subject,
    text: content,
    html: message,
  };

  await transporter.sendMail(mailOptions);
};

const sendEmailAOALinkFromGmail = async (emails, subject, content) => {
  const message = formatEmail(content);

  // Send the PDF to the provided email address
  const transporter = nodemailer.createTransport({
    service: "gmail",
    port: 465,
    secure: true,
    logger: true,
    debug: true,
    secureConnection: false,
    auth: {
      user: "leavenowgrow@gmail.com",
      pass: "pczt ypxq cwpz frvw",
    },
    tls: {
      rejectUnauthorized: true,
    },
  });

  const mailOptions = {
    from: "leavenowgrow@gmail.com",
    to: emails,
    subject: subject,
    text: content,
    html: message,
  };

  await transporter.sendMail(mailOptions);
};

module.exports = { sendEmailAOALinkFromGmail, sendEmailAOALinkFromMS365 };
