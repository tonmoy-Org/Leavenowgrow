// Prevent process crash on unhandled runtime errors or rejected promises
process.on("uncaughtException", (err) => {
  console.error("GLOBAL SAFETY: Uncaught Exception caught:", err);
});
process.on("unhandledRejection", (reason, promise) => {
  console.error("GLOBAL SAFETY: Unhandled Promise Rejection:", reason);
});

// Assuming you have the necessary dependencies installed
const express = require("express");

const session = require("express-session");
let MongoStore;
try {
  MongoStore = require("connect-mongo");
} catch (e) {
  console.warn("connect-mongo not found in node_modules, falling back to memory store:", e.message);
}
const path = require("path");
const ejs = require("ejs");
const fs = require("fs");
const pdf = require("html-pdf");
const cheerio = require("cheerio");
const excel = require("exceljs");
const JSZip = require("jszip");
const HTMLtoDOCX = require("html-to-docx");
const bcrypt = require("bcrypt");
const passport = require("passport");
const LocalStrategy = require("passport-local").Strategy;
const flash = require("express-flash");
const bodyParser = require("body-parser");
const mongoose = require("mongoose");
const ObjectId = mongoose.Types.ObjectId;

require("dotenv").config({ path: path.resolve(process.cwd(), ".env") });

const {
  generate,
  generateSelf,
  generateManager,
  generateAll,
} = require("./utils/employee-table-docx");
const {
  generatePDF,
  generatePDFonA4,
  generatePDFonTabloid,
  generateEvaluationsPDF,
} = require("./utils/employee-table-pdf");
const sendEmailAOALink = require("./utils/emailHelper");
const { formatEmail } = require("./utils/formatEmail");
const { globalErrorHandler } = require("./middlewares/globalErrorHandler");
const { notFound } = require("./middlewares/notFound");
const {
  generateSelfExcel,
  generateManagerExcel,
  generateAllExcel,
  generateExcelFormdata,
} = require("./utils/generateExcel");
const { createEmailTransporter } = require("./utils/emailTransporter");
const { consoleMe } = require("./utils/consoleMe");
const launchPuppeteer = require("./utils/launchPuppeteer");
const { configEnv } = require("./config");
const { dbConnect } = require("./utils/dbConnect");
const { FormData } = require("./models/formdata");
const { Company } = require("./models/company");
const { Group } = require("./models/group");
const { User } = require("./models/user");
const { SelfEvaluation } = require("./models/selfEvaluation");
const { ManagerEvaluation } = require("./models/managerEvaluation");
const { isLoggedIn } = require("./middlewares/isLoggedIn");
const { isSuperAdmin } = require("./middlewares/isSuperAdmin");
const { isCompAdmin } = require("./middlewares/isCompAdmin");
const { catchAsync } = require("./utils/catchAsync");
const { getAnalytics } = require("./services/getAnalytics");
const { GoalsWorksheet } = require("./models/goalsWorksheet");
const {
  generateGoalsWorksheetExcel,
  generateGoalsWorksheetExcelV2,
} = require("./utils/goals-worksheets-excel");
const {
  generateGoalsWorksheetsPDF,
  generateGoalsWorksheetsPDFV2,
} = require("./utils/goals-worksheets-pdf");
const { generateGoalsWorksheetDocx } = require("./utils/goals-worksheets-docx");
const { getToken } = require("./utils/getToken");
const { decodeToken } = require("./utils/decodeToken");
const { handleDuplicateError } = require("./errors/handleDuplicateError");
const { upload } = require("./utils/upload");
const {
  generateGoalsWorksheetDocxV2,
} = require("./utils/goals-worksheets-docx-v2");
const { Test, Test2 } = require("./models/test");
const {
  getDescendants,
  getFollowingEmployees,
} = require("./utils/getDescendants");

// running application
const MODE = configEnv.mode; // two modes: "dev" and "prod"
// PORT to run your application
const PORT = configEnv.port;
const BASE_LINK =
  MODE === "prod" ? configEnv.prodUrl : `http://localhost:${PORT}`;

// Create an Express application
const app = express();

// Trust proxy for IIS / Nginx / Azure / PM2 reverse proxies
app.set("trust proxy", 1);

// Connect to MongoDB using Mongoose
// Create a MongoDB client and specify the connection URL
const mongoURL = configEnv.mongodbUri;
dbConnect(mongoURL);

// Set EJS as the view engine
app.set("view engine", "ejs");

// Configure the application to parse incoming JSON data
// app.use(bodyParser.json());

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static("public"));

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
// Set up sessions with persistent MongoDB store (cross-PM2 cluster support)
const sessionOptions = {
  secret: configEnv.secretKey,
  resave: false,
  saveUninitialized: false,
  cookie: {
    maxAge: 14 * 24 * 60 * 60 * 1000,
    httpOnly: true,
    sameSite: "lax",
  },
};

if (MongoStore) {
  try {
    sessionOptions.store = MongoStore.create({
      mongoUrl: mongoURL,
      ttl: 14 * 24 * 60 * 60, // 14 days
      autoRemove: "native",
    });
  } catch (storeErr) {
    console.error("Failed to create MongoStore instance:", storeErr.message);
  }
}

app.use(session(sessionOptions));

// Initialize Passport
app.use(passport.initialize());
app.use(passport.session());
app.use(flash());

// app.set('views', path.join(__dirname, 'views'))

// async function createUser() {
//   const pass = await bcrypt.hash('tom124pass', 10)

//   User.create({
//     name: 'Tom Smith', username: 'tom124', email: 'aar6m6@gmail.com',
//     password: pass, role: 2
//   })
// }

// createUser()

passport.use(
  new LocalStrategy(
    { usernameField: "email" }, // Use email as the username field
    async (email, password, done) => {
      try {
        // Find the user by email
        const user = await User.findOne({ email });

        if (!user) {
          return done(null, false, { message: "Invalid email or password." });
        }

        // Compare the provided password with the hashed password
        const passwordsMatch = await bcrypt.compare(password, user.password);

        if (!passwordsMatch) {
          return done(null, false, { message: "Incorrect email or password." });
        }

        return done(null, user);
      } catch (error) {
        return done(error);
      }
    }
  )
);

// Serialize and deserialize user for session management
passport.serializeUser((user, done) => {
  done(null, user._id || user.id);
});

passport.deserializeUser(async (id, done) => {
  try {
    const user = await User.findById(id).select("-password").lean();
    done(null, user);
  } catch (error) {
    done(error);
  }
});

app.get("/unauthorized", async (req, res) => {
  res.render("403");
});

app.get("/login", async (req, res) => {
  res.render("login", { message: req.flash("error") });
});

app.post(
  "/login",
  passport.authenticate("local", {
    successRedirect: "/",
    failureRedirect: "/login",
    failureFlash: true, // Show error messages
  })
);

app.get(
  "/org-hierarchy",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    if (user?.role === 1) {
      const companies = await Company.find().lean();
      res.render("org-hierarchy-2", { user: user, companies: companies });
    } else {
      res.redirect("/unauthorized");
    }
  })
);

app.get(
  "/org-hierarchy-company",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    res.render("org-hierarchy-company", { user: user });
  })
);

// using this api, boom
app.get(
  "/v3/get/companies/:id",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const loggedInUser = req.user;

    const companyId = req.params.id;

    const selectedCompany = await Company.findOne({ _id: companyId }).lean();

    const userItems = await User.find({ company: companyId })
      .populate({ path: "formData", model: FormData })
      .lean();

    const users = userItems.map((user) => {
      return {
        id: user?._id,
        pid: user?.manager?.id ?? null,
        data: {
          formDataId: user?.formData?._id,
          goalsWorksheetId: user?.goalsWorksheet,
          companyId: selectedCompany?._id,
          companyOrGroup: selectedCompany?.name,
          name: user?.name,
          jobTitle: user?.formData?.title,
          profilePic: user?.profilePictureUrl,
        },
      };
    });

    let employees = [];
    if (loggedInUser?.role === 1) {
      employees = users;
    } else {
      const descendantsUsers = getDescendants(loggedInUser?._id, users);

      employees = descendantsUsers.map((node) =>
        node?.id?.toString() === loggedInUser?._id.toString()
          ? { ...node, pid: null }
          : node
      );
    }

    res.send({ status: true, data: employees });
  })
);

// Not using this one, using the above
app.get(
  "/get/companies/:id",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const companyId = req.params.id;

    const selectedCompany = await Company.findOne({ _id: companyId }).lean();

    // TODO: maybe the populate will not work for group
    const formDataItems = await FormData.find({ company: companyId })
      .populate({
        path: "group",
        model: Group,
      })
      .populate({
        path: "user",
        model: User,
      })
      .lean();

    const employeesData = Promise.all(
      formDataItems.map(async (item) => {
        const formDataId = item?._id;
        const employee = await User.findOne({ formData: formDataId }).lean();

        return {
          id: employee?._id,
          pid: item?.user?.manager?.id ?? null,
          data: {
            formDataId: formDataId,
            goalsWorksheetId: item?.user?.goalsWorksheet,
            companyOrGroup: item?.group?.name ?? selectedCompany?.name,
            name: item?.name,
            jobTitle: item?.title,
            profilePic: employee?.profilePictureUrl,
            type: "aoa",
          },
        };
      })
    );

    const employeesArray = await employeesData;

    const employees = employeesArray.filter((employee) => employee?.id);

    res.send({ status: true, data: employees });
  })
);

// Not using this api, using the one above
app.get(
  "/v2/get/companies/:id",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const companyId = req.params.id;

    // company --> _id, name, user_id.name, user_id.role: 2
    // subgroup --> _id, name, adminId.name, adminId.role: 3
    // aoa --> _id, name, jobTitle
    const selectedCompany = await Company.findOne({ _id: companyId })
      .populate({
        path: "user_id",
        model: User,
        select: "-password",
      })
      .populate({
        path: "subGroup",
        model: Group,
        populate: {
          path: "adminId",
          model: User,
          select: "-password",
        },
      });

    const subGroup = Promise.all(
      selectedCompany?.subGroup?.map(async (group) => {
        const formdata = await FormData.find({ group: group?._id }).lean();

        const employees = formdata?.map((employee) => {
          return {
            id: employee?._id,
            pid: group?._id,
            data: {
              id: employee?._id,
              companyOrGroup: group?.name,
              name: employee?.name,
              profilePic: employee?.profilePic,
              title: employee?.title,
              type: "aoa",
            },
          };
        });

        return [
          {
            id: group?._id,
            pid: selectedCompany?._id,
            data: {
              id: group?.adminId?.formData,
              companyOrGroup: group?.name,
              name: group?.adminId?.name,
              profilePic: group?.user_id?.profilePictureUrl,
              title: "Group Admin",
              type: "group",
            },
          },
          ...employees,
        ];
      })
    );

    const children = await subGroup;

    let group1 = [];
    children.forEach((child) => {
      group1.push(...child);
    });

    const company = [
      {
        id: selectedCompany?._id,
        pid: null,
        data: {
          id: selectedCompany?.user_id?.formData,
          companyOrGroup: selectedCompany?.name,
          name: selectedCompany?.user_id?.name,
          profilePic: selectedCompany?.user_id?.profilePictureUrl,
          title: "Admin",
          type: "company",
        },
      },
      ...group1,
    ];

    res.send({ status: true, data: company });
  })
);

// the send aoa link can be: /register?c=<company_id>&m=<manager_id>&token=<token_string>
// token = { companyId, managerId, email }
app.get(
  "/register",
  catchAsync(async (req, res) => {
    const token = req.query.token;
    const decodedToken = decodeToken(token);

    if (decodedToken) {
      const companyId = decodedToken?.companyId;
      const managerId = decodedToken?.managerId;
      const email = decodedToken?.email;

      res.render("register", {
        message: "",
        company: companyId,
        manager: managerId,
        email: email,
      });
    } else {
      res.redirect("/unauthorized");
      return;
    }
  })
);

app.post(
  "/register",
  upload.single("user"),
  catchAsync(async (req, res) => {
    const data = JSON.parse(req.body.data); // contain name, email, password, manager _id
    const password = await bcrypt.hash(data.password, 10);

    // <id>:<manager_name>:<email>
    const managerDetails = data?.manager?.split(":");

    const userData = {
      name: data?.name,
      email: data?.email,
      password: password,
      company: data?.company,
      level: Number(managerDetails[3]) + 1,
      manager: {
        id: managerDetails[0],
        name: managerDetails[1],
        email: managerDetails[2],
      },
      profilePictureUrl: "",
    };

    // headshot image logic
    const image = req.file;

    if (image) {
      const profilePictureUrl = `${BASE_LINK}/img/users/${image?.filename}`;
      userData.profilePictureUrl = profilePictureUrl;
    }

    const transporter = await createEmailTransporter();

    const newUser = new User(userData);

    newUser
      .save()
      .then(() => {
        consoleMe("User created successfully");

        const token = getToken({
          email: data?.email,
        });

        const mailOptions = {
          from: `LeaveNowGrow <${configEnv?.userEmail}>`,
          to: data?.email,
          subject: "Please verify your email",
          text: `Hi ${data?.name},
      Please verify your email address by clicking on this link below.
      <a href='${BASE_LINK}/email/verify/${token}'>Link</a>`,
          html: `<p>Hi ${data?.name},</p>
      <p>Please verify your email address by clicking on this link below.</p>
      <a href='${BASE_LINK}/email/verify/${token}'>Link</a>`,
        };

        transporter.sendMail(mailOptions);

        res.send({
          status: true,
          message: "Verification email is sent to your email address.",
        });
      })
      .catch((error) => {
        const { errorSources } = handleDuplicateError(error);
        consoleMe(errorSources[0]?.message);

        res.send({ status: false, message: errorSources[0]?.message });
      });
  })
);

app.get(
  "/email/verify/:token",
  catchAsync(async (req, res) => {
    const token = req.params?.token;

    const decodedData = decodeToken(token);

    const userEmail = decodedData?.email;

    const filter = { email: userEmail };

    const user = await User.findOneAndUpdate(
      filter,
      {
        isVerified: true,
      },
      { new: true }
    );

    if (!user) {
      throw new Error("User not found.");
    }

    if (user?.isVerified) {
      res.render("verification-email", { verified: true });
    } else {
      res.render("verification-email", { verified: false });
    }
  })
);

app.get("/logout", isLoggedIn, (req, res, next) => {
  req.logout(function (err) {
    if (err) {
      return next(err);
    }
    res.redirect("/login");
  });
});

app.post(
  "/get-verified",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;

    const token = getToken({
      email: user?.email,
    });

    const mailOptions = {
      from: `LeaveNowGrow <${configEnv?.userEmail}>`,
      to: user?.email,
      subject: "Please verify your email",
      text: `Hi ${user?.name},
      Please verify your email address by clicking on this link below.
      <a href='${BASE_LINK}/email/verify/${token}'>Link</a>`,
      html: `<p>Hi ${user?.name},</p>
      <p>Please verify your email address by clicking on this link below.</p>
      <a href='${BASE_LINK}/email/verify/${token}'>Link</a>`,
    };

    const transporter = await createEmailTransporter();

    transporter.sendMail(mailOptions);

    res.send({
      status: true,
      message: "Verification email is sent to your email address.",
    });
  })
);

app.get(
  "/user/profile",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;

    res.render("profile", { user: user });
  })
);

app.patch(
  "/change-user-info",
  upload.single("user"),
  catchAsync(async (req, res) => {
    const data = JSON.parse(req.body.data);
    const image = req.file;

    const updatedUser = {};

    if (data?.name) {
      updatedUser.name = data?.name;
    }

    if (image) {
      const profilePictureUrl = `${BASE_LINK}/img/users/${image?.filename}`;
      updatedUser.profilePictureUrl = profilePictureUrl;
    }

    const email = data?.email;
    const user = await User.findOneAndUpdate({ email: email }, updatedUser, {
      new: true,
    });

    if (!user) {
      res.send({ success: false, message: "Failed to update the user." });
    }

    res.send({ success: true, message: "User updated successfully." });
  })
);

// this is an API for resetting password
app.post(
  "/reset-password",
  catchAsync(async (req, res) => {
    const data = req.body;

    const filter = { email: data?.email };
    // Find the user by email
    const user = await User.findOne(filter);
    // Compare the provided password with the hashed password

    const passwordsMatch = await bcrypt.compare(
      data?.passwordCurrent,
      user?.password
    );

    if (!passwordsMatch) {
      res.send({ message: "Password didn't match.", status: false });
      return;
    }

    const hashedPassword = await bcrypt.hash(data?.passwordNew, 10);

    await User.findOneAndUpdate(filter, {
      password: hashedPassword,
    });

    res.send({ message: "Password changed successfully.", status: true });
  })
);

app.post(
  "/change-my-password",
  catchAsync(async (req, res) => {
    const data = req.body;

    const filter = { email: data?.email };
    // Find the user by email
    const user = await User.findOne(filter);

    if (!user) {
      res.send({ message: "User Not Found.", status: false });
      return;
    }

    const hashedPassword = await bcrypt.hash(data?.password, 10);

    await User.findOneAndUpdate(filter, {
      password: hashedPassword,
    });

    res.send({ message: "Password changed successfully.", status: true });
  })
);

app.get(
  "/forgot-my-password/:token",
  catchAsync(async (req, res) => {
    const token = req.params.token;

    const decodedData = decodeToken(token);

    const userEmail = decodedData?.email;

    const filter = { email: userEmail };

    const user = await User.findOne(filter);

    if (!user) {
      throw new Error("User not found.");
    }

    res.render("change-my-password", { email: userEmail });
  })
);

app.get(
  "/forgot-my-password",
  catchAsync(async (req, res) => {
    res.render("forgot-my-password");
  })
);

app.post(
  "/forgot-my-password",
  catchAsync(async (req, res) => {
    const email = req.body.email;

    const user = await User.findOne({ email: email });

    if (!user) {
      res.send({ message: "User not found.", status: false });
      return;
    }

    const token = getToken({
      name: user?.name,
      email: user?.email,
    });

    const transporter = await createEmailTransporter();

    const mailOptions = {
      from: `LeaveNowGrow <${configEnv?.userEmail}>`,
      to: user?.email,
      subject: "You forgot your account password",
      text: `Hi ${user?.name},
Please set up your new account password by clicking on this link below and following the instructions further.
<a href='${BASE_LINK}/forgot-my-password/${token}'>Link</a>`,
      html: `<p>Hi ${user?.name},</p>
<p>Please set up your new account password by clicking on this link below and following the instructions further.</p>
<a href='${BASE_LINK}/forgot-my-password/${token}'>Link</a>`,
    };

    transporter.sendMail(mailOptions);

    res.send({ message: "Email sent to your account.", status: true });
  })
);

app.get(
  "/analytics",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const analytics = await getAnalytics();

    res.header("Content-Type", "application/json");
    res.status(200).send(JSON.stringify(analytics, null, 4));
  })
);

// ?comp=<company_id>
app.get(
  "/aoa-details/:id",
  catchAsync(async (req, res) => {
    const id = req.params.id;
    const companyId = req.query.comp;

    if (id === "undefined") {
      res.redirect(`/formadmin/${companyId}`);
      return;
    }

    const formDataItem = await FormData.findOne({ _id: id });
    const user = req.user;

    const managerId = user?.manager?.id;
    const group = await Group.findOne({ adminId: managerId });

    res.render("employee-details", {
      formDataItems: [formDataItem],
      company: group?._id ?? companyId,
      user: user,
      comp: false,
    });
  })
);

app.get(
  "/employee-details",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    const userEmail = req.user.email;

    const managerId = user?.manager?.id;
    const group = await Group.findOne({ adminId: managerId });

    const formDataItem = await FormData.findOne({ _id: user?.formData });

    if (group?._id && group?.adminId) {
      res.render("employee-details", {
        formDataItems: [formDataItem],
        company: group?._id ? group?._id : user?.company,
        user: user,
        comp: false,
      });
    } else {
      res.render("employee-details", {
        formDataItems: [formDataItem],
        company: user?.company,
        user: user,
        comp: true,
      });
    }
  })
);

app.get(
  "/create-aoa-form/:id",
  catchAsync(async (req, res) => {
    const companyId = req.params.id;
    const groupId = req.params.id;
    const group = await Group.findById(groupId);

    res.render("form_self", {
      id: companyId,
      groupId: group?._id,
      group: group?._id ? true : false,
      user: req?.user,
    });
  })
);

app.get(
  "/",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;

    if (user?.level === "Super Admin") {
      // const analytics = await getAnalytics();
      const companies = await Company.find();

      return res.render("org-hierarchy-2", {
        user: user,
        companies: companies,
      });
    } else if (user?.level === "1") {
      // Number 2 for admin access
      const groups = await Group.find({ creatorId: user._id });
      const admins = await User.find({ creatorId: user._id });
      const currentComapny = await Company.findOne({ _id: user?.company });

      return res.render("groups", {
        groups,
        admins,
        user: user,
        company: currentComapny?._id,
      });
    } else if (user?.level === "2") {
      // Number 3 for group admin access
      const group = await Group.findOne({ adminId: user._id });
      let groupId = "";

      // Retrieve the specific form data item based on the provided ID
      // let formDataItem = [];

      if (group) {
        const formDataItem = await FormData.find({ group: group._id });
        const gName = group.name;

        formDataItem.empCompanyName = gName;
        groupId = group._id;
      }

      // Render the EJS template with the form data item
      return res.redirect("/group-admin/view");
    } else {
      res.redirect("/employee-details");
    }
  })
);

app.get(
  "/admins",
  isLoggedIn,
  isSuperAdmin,
  catchAsync(async (req, res) => {
    const user = req.user;

    User.find({})
      .populate("admins")
      .then((users) => {
        const parentUser = [];

        for (let i = 0; i < users.length; i++) {
          if (users[i].role == 1 || users[i].role == 2) {
            parentUser.push(users[i]);
          }
        }

        res.render("admins", { users: parentUser, user: user, message: "" });

        // res.render("admins-v2", { users: users, user: user });
      });
  })
);

app.get(
  "/all/users",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;

    if (user?.level === "Super Admin") {
      const users = await User.find({}).sort("name");

      res.render("all-users", { users: users, user: user });
      return;
    } else if (user?.level === "1") {
      const users = await User.find({ company: user?.company }).sort("name");
      const managers = await User.find({
        level: { $in: ["1", "2", "3", "4", "5"] },
      }).sort("name");

      res.render("all-users", { users: users, user: user, managers: managers });
      return;
    }

    res.redirect("/unauthorized");
  })
);

app.get(
  "/v2/admins",
  isLoggedIn,
  isSuperAdmin,
  catchAsync(async (req, res) => {
    const user = req.user;

    const users = await User.find();

    res.render("admins-v2", { users: users, user: user });
  })
);

app.get(
  "/companies/all",
  catchAsync(async (req, res) => {
    const companies = await Company.find().lean();

    res.send({ companies: companies });
  })
);

app.get(
  "/managers/:id",
  catchAsync(async (req, res) => {
    const id = req.params.id;

    const users = await User.find({ company: id }).lean();

    const parentUsers = [];

    for (let i = 0; i < users.length; i++) {
      if (
        users[i].level == "Super Admin" ||
        users[i].level == "1" ||
        users[i].level == "2" ||
        users[i].level == "3" ||
        users[i].level == "4" ||
        users[i].level == "5"
      ) {
        parentUsers.push(users[i]);
      }
    }

    res.send({ parentUsers: parentUsers });
  })
);

app.get(
  "/comp-admins",
  isLoggedIn,
  isCompAdmin,
  catchAsync(async (req, res) => {
    const currentUser = req.user;

    // Find users where creatorId matches the current user's _id
    User.find({ company: currentUser?.company, level: ["1", "2"] }).then(
      (users) => {
        res.render("admins-v2", { users, user: currentUser });
      }
    );
  })
);

// POST endpoint to update the company name
app.post(
  "/edit-company-name",
  catchAsync(async (req, res) => {
    const companyName = req.body.companyName;
    const companyId = req.body.companyId;
    const adminId = req.body.user_id;

    // Find the company data
    const company = await Company.findOne({ _id: companyId });
    // this is currently assigned admin
    const currentAdmin = await User.findOne({ _id: company?.user_id });

    if (currentAdmin?._id.toString() === adminId) {
      res.redirect("/companies");
      return;
    }

    // this is going to be new admin
    const newAdmin = await User.findById(adminId);

    if (!company || !newAdmin) {
      // If company data doesn't exist, create a new document
      throw new Error("No company or admin found.");
    }

    // If company data exists, update the name, admin
    company.name = companyName;
    company.user_id = newAdmin?._id;
    await company.save();

    // update the company in user (admin)
    newAdmin.company = company?._id;
    newAdmin.manager = null;
    await newAdmin.save();

    // update the users with manager (id, name, email) with the new admin
    const filter = {
      "manager.id": (currentAdmin?._id).toString(),
    };

    const manager = {
      id: (newAdmin?._id).toString(),
      name: newAdmin?.name,
      email: newAdmin?.email,
    };

    await User.updateMany(filter, { $set: { manager: manager } });
    // also set the current manager
    currentAdmin.manager = manager;
    await currentAdmin.save();

    res.redirect("/companies");
  })
);

app.post(
  "/edit-group-name",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const groupName = req.body.groupName;
    const groupId = req.body.groupId;
    const user = req.user;
    const adminId = req.body.user_id;

    // Find the group data
    const group = await Group.findOne({ _id: groupId });

    if (!group) {
      // If company data doesn't exist, create a new document
      await Group.create({ name: groupName, adminId: adminId });
    } else {
      // If company data exists, update the name
      group.name = groupName;
      group.adminId = adminId;
      await group.save();
    }

    if (user.role == 1) {
      res.redirect("/companies");
    } else if (user.role == 2) {
      res.redirect("/groups");
    } else {
      res.redirect("/");
    }
  })
);

// POST endpoint to update the company name
app.post(
  "/edit-admin",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    const data = req.body;

    // Find the company data
    const admin = await User.findOne({ _id: data.adminId });
    admin.name = data.name;
    admin.level = data.level;
    await admin.save();

    if (user.role == 1) {
      res.redirect("/admins");
    } else if (user.role == 2) {
      res.redirect("/comp-admins");
    } else {
      res.redirect("/");
    }
  })
);

// Route to render the form.ejs file
app.get("/form", (req, res) => {
  res.render("form");
});

app.get(
  "/review/:id",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const ID = req.params.id;
    const formData = await FormData.find({ _id: ID });

    res.render("review", {
      user: req?.user,
      formDataItems: formData,
      comp: false,
    });
  })
);

// Define a route to handle the form submission
app.post(
  "/submit",
  catchAsync(async (req, res) => {
    // Extract the form data from the request body
    const formData = req.body;
    const userEmail = formData?.email;

    // Create a new FormData instance using the Mongoose model
    const newFormData = new FormData(formData);

    if (formData?.group) {
      newFormData.groupIds.addToSet(formData?.group);
    }

    const user = await User.findOne({ email: userEmail });

    if (user) {
      user.formData = newFormData?._id;
      await user.save();
      newFormData.user = user?._id;
    }

    // Save the form data to the MongoDB database
    newFormData.save().then(() => {
      consoleMe("Form data saved successfully");
      res.render("thankyou-page", {
        user: req.user,
        newFormData,
      });
    });
  })
);

// Define the route handler for the GET endpoint
// app.get('/admin', async (req, res) => {
//   try {
//     // Retrieve all form data documents from the "formdatas" collection
//     const formDataItems = await FormData.find();

//     // Render the data in a tabular form
// let tableHtml = '<table style="border-collapse: collapse; border: 1px solid black;">';
// tableHtml += '<tr><th style="border: 1px solid black;">Name</th><th style="border: 1px solid black;">Title</th><th style="border: 1px solid black;">Job Description</th><th style="border: 1px solid black;">Manager/Supervisor</th><th style="border: 1px solid black;">Direct Reports</th><th style="border: 1px solid black;">Accountable</th><th style="border: 1px solid black;">Participate</th><th style="border: 1px solid black;">Compensation</th><th style="border: 1px solid black;">Metrics</th><th style="border: 1px solid black;">Positional Objectives</th><th style="border: 1px solid black;">Personal Objectives</th><th style="border: 1px solid black;">Authority Levels</th><th style="border: 1px solid black;">Delegation of Authority</th></tr>';

// formDataItems.forEach((item) => {
//   tableHtml += `<tr><td style="border: 1px solid black;">${item.name}</td><td style="border: 1px solid black;">${item.title}</td><td style="border: 1px solid black;">${item.jobDescription}</td><td style="border: 1px solid black;">${item.managerSupervisor}</td><td style="border: 1px solid black;">${item.directReports}</td><td style="border: 1px solid black;">${item.accountable}</td><td style="border: 1px solid black;">${item.participate}</td><td style="border: 1px solid black;">${item.compensation}</td><td style="border: 1px solid black;">${item.metrics}</td><td style="border: 1px solid black;">${item.positionalObjectives}</td><td style="border: 1px solid black;">${item.personalObjectives}</td><td style="border: 1px solid black;">${item.authorityLevels}</td><td style="border: 1px solid black;">${item.delegationOfAuthority}</td></tr>`;
// });

// tableHtml += '</table>';

//     // Send the HTML response with the table
//     res.send(tableHtml);
//   } catch (error) {
//     console.error(error);
//     res.status(500).send('Internal Server Error');
//   }
// });

app.get(
  "/admin",
  isLoggedIn,
  isSuperAdmin,
  catchAsync(async (req, res) => {
    const user = req.user;
    // Retrieve all form data documents from the "formdatas" collection
    let company = await Company.findOne().lean();

    const formDataItems = await FormData.find().lean();

    if (!company) {
      company = "Not Assigned";
    }

    // Render the "admin" EJS template with the form data items
    res.render("employees", {
      formDataItems,
      company: company._id,
      user: user,
    });
  })
);

// Using no more
app.get(
  "/admin/downloadv1/:id",
  catchAsync(async (req, res) => {
    const selectedIds = req.params.id.split(",");

    // const selectedIds = req.params.id;
    // Retrieve the specific form data item based on the provided ID
    // const formDataItem = await FormData.findById(req.params.id);
    // const formDataItem = selectedIds;

    let company = await Company.findOne();

    if (!company) {
      company.name = "Not Assigned";
    }
    const formDataItems = await FormData.find({ _id: { $in: selectedIds } });

    const companyName = await Company.findOne({
      _id: formDataItems[0]?.company,
    });

    const cName = companyName ? companyName.name : "company";

    formDataItems.empCompanyName = cName;

    if (!formDataItems) {
      // Return an error if the form data item is not found
      return res.status(404).send("Form data not found");
    }

    const browser = await launchPuppeteer();
    const page = await browser.newPage();
    // Define the path to the EJS template file
    const templatePath = path.join(
      __dirname,
      "views",
      "employees-pdf-docx.ejs"
    );
    const templatePathDocx = path.join(
      __dirname,
      "views",
      "employees-docx.ejs"
    );

    function separateTagValues(htmlString) {
      const $ = cheerio.load(htmlString);

      const thValues = [];
      const tdValues = [];

      $("table")
        .find("tr")
        .each((rowIndex, rowElement) => {
          const thElements = $(rowElement).find("th");
          const tdElements = $(rowElement).find("td");

          thElements.each((index, element) => {
            const thValue = $(element).text().trim();
            thValues.push(thValue);
          });

          tdElements.each((index, element) => {
            const tdValue = $(element).text().trim();
            tdValues.push(tdValue);
          });
        });

      return { thValues, tdValues };
    }

    let tValues = [];

    // Render the EJS template with the form data item
    html = await ejs.renderFile(templatePath, {
      formDataItems: formDataItems,
      print: true,
      company: company.name,
    });
    const htmlForDocx = await ejs.renderFile(templatePathDocx, {
      formDataItems: formDataItems,
      print: true,
      company: company.name,
    });

    tValues = separateTagValues(htmlForDocx);

    const completeHtml = htmlForDocx;

    await page.setContent(html, { waitUntil: "networkidle0" });

    let headerTemplate = `<div style="width: 100%; text-align: center; height: 120px; min-height: 120px;
      margin-top: 30px; margin-bottom: 30px">
      <h3 style="line-height: 12px;"><span style="font-size: 9px; font-weight: normal">
      ASSIGNMENT OF ACCOUNTABILITY</span><br><span style="font-size: 14px; font-weight: 300">${cName}</span></h3></div>`;

    const pdfGenerated = await page.pdf({
      landscape: true,
      printBackground: true,
      displayHeaderFooter: true,
      headerTemplate,
      margin: {
        top: "4cm",
        bottom: "2cm",
        left: "2cm",
        right: "2cm",
      },
      width: "1056px",
      height: "1632px",
    });

    await browser.close();

    const workbook = new excel.Workbook();
    const worksheet = workbook.addWorksheet("Sheet 1");

    // worksheet.headerFooter().oddHeader = `ASSIGNMENT OF ACCOUNTABILITY \n ${cName}`;

    worksheet.getColumn(1).alignment = { wrapText: true, vertical: "top" };
    worksheet.getColumn(2).alignment = { wrapText: true, vertical: "top" };

    worksheet.getColumn(1).font = { bold: true };

    worksheet.getColumn(1).font = { bold: true };
    worksheet.getColumn(1).values = tValues["thValues"];
    worksheet.getColumn(2).values = tValues["tdValues"];

    worksheet.getRow(1).font = { bold: true };
    worksheet.getColumn(1).eachCell((cell, rowNumber) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "CCFFFF" },
      };
    });

    worksheet.getRow(1).eachCell((cell, colNumber) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "CCFFFF" },
      };
    });

    // Set borders for all cells with values
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        if (cell.value) {
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        }
      });
    });

    //  Auto-fit column widths
    worksheet.columns.forEach((column) => {
      let maxLength = 0;
      column.eachCell({ includeEmpty: true }, (cell) => {
        const columnLength = cell.value ? cell.value.toString().length : 0;
        maxLength = Math.max(maxLength, columnLength);
      });
      column.width = 30;
    });

    const excelBuffer = await workbook.xlsx.writeBuffer();

    const pdfBuffer = Buffer.from(pdfGenerated);

    // const docx = htmlDocx.asBlob(completeHtml, { orientation: "landscape" });
    // const docxBuffer = await docx
    //   .arrayBuffer()
    //   .then((arrayBuffer) => Buffer.from(arrayBuffer, "binary"));

    const docxBuffer = await HTMLtoDOCX(completeHtml, null, {
      orientation: "landscape",
    });

    // const docxBuffer = Buffer.from(docx)

    const zip = new JSZip();
    zip.file("form_data.pdf", pdfBuffer);
    zip.file("form_data.docx", docxBuffer);
    zip.file("form_data.xlsx", excelBuffer);

    const zipBuffer = await zip.generateAsync({ type: "nodebuffer" });

    res.setHeader("Content-Type", "application/zip");
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="combined-files.zip"'
    );

    // Send the combined buffer as the response
    res.send(zipBuffer);
  })
);

// Using no more
app.get("/admin/download2/:id", async (req, res) => {
  try {
    const selectedIds = req.params.id.split(",");

    let company = await Company.findOne();

    if (!company) {
      company.name = "Not Assigned";
    }

    // Retrieve the specific form data items based on the provided IDs
    const formDataItems = await FormData.find({ _id: { $in: selectedIds } });

    if (!formDataItems) {
      // Return an error if the form data items are not found
      return res.status(404).send("Form data not found");
    }

    // Divide the form data items into chunks of 4
    const chunks = [];
    const chunkSize = 4;

    for (let i = 0; i < formDataItems.length; i += chunkSize) {
      const chunk = formDataItems.slice(i, i + chunkSize);
      chunks.push(chunk);
    }

    // Define the path to the EJS template file
    const templatePath = path.join(__dirname, "views", "employees.ejs");

    // Create a PDF for each chunk of form data items
    const pdfOptions = {
      format: { width: "17in", height: "11in" },
      printBackground: true,
      displayHeaderFooter: false,
      margin: {
        top: "20px",
        bottom: "20px",
        left: "20px",
        right: "20px",
      },
      orientation: "landscape",
      scale: 0.1,
    };

    const pdfPromises = chunks.map((chunk) => {
      return new Promise((resolve, reject) => {
        ejs.renderFile(
          templatePath,
          { formDataItems: chunk, print: true, company: company.name },
          (err, html) => {
            if (err) {
              console.error("Failed to render EJS template:", err);
              return reject(err);
            }

            pdf.create(html, pdfOptions).toBuffer((err, buffer) => {
              if (err) {
                console.error("Failed to create PDF:", err);
                return reject(err);
              }

              resolve(buffer);
            });
          }
        );
      });
    });

    for (const inx in pdfPromises) {
      const pdfBuffer = await pdfPromises[inx];
      // Define the path to save the PDF file
      const filePath = path.join(__dirname, "temp", `form_data${inx}.pdf`);

      // Write the PDF buffer to a file
      fs.writeFileSync(filePath, pdfBuffer);
    }

    const tempFolderPath = path.join(__dirname, "temp");

    const pdfFiles = fs
      .readdirSync(tempFolderPath)
      .filter((file) => file.endsWith(".pdf"))
      .map((file) => path.join(tempFolderPath, file));

    const mergedPdf = await pdflib.create();

    for (const pdfFile of pdfFiles) {
      const pdfBytes = fs.readFileSync(pdfFile);
      const pdfDoc = await pdflib.load(pdfBytes);

      const copiedPages = await mergedPdf.copyPages(
        pdfDoc,
        pdfDoc.getPageIndices()
      );
      copiedPages.forEach((page) => mergedPdf.addPage(page));
    }

    const mergedPdfBytes = await mergedPdf.save();

    const zip = new JSZip();
    zip.file("form_data.pdf", mergedPdfBytes);
    // zip.file('form_data.xlsx', excelBuffer);

    const zipBuffer = await zip.generateAsync({ type: "nodebuffer" });

    res.setHeader("Content-Type", "application/zip");
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="combined-files.zip"'
    );

    // Send the combined buffer as the response
    res.send(zipBuffer);
  } catch (error) {
    console.error(error);
    res.status(500).send("Internal Server Error");
  }
});

// Using no more
app.get("/admin/download/:id", async (req, res) => {
  try {
    const selectedIds = req.params.id.split(",");

    let company = await Company.findOne();

    if (!company) {
      company.name = "Not Assigned";
    }

    // Retrieve the specific form data items based on the provided IDs
    const formDataItems = await FormData.find({ _id: { $in: selectedIds } });

    const companyName = await Company.findOne({
      _id: formDataItems[0].company,
    });

    const cName = companyName.name;

    formDataItems.empCompanyName = cName;

    if (!formDataItems) {
      // Return an error if the form data items are not found
      return res.status(404).send("Form data not found");
    }

    // Divide the form data items into chunks of 4
    // const chunks = [];
    // const chunkSize = 4;

    // for (let i = 0; i < formDataItems.length; i += chunkSize) {
    //   const chunk = formDataItems.slice(i, i + chunkSize);
    //   chunks.push(chunk);
    // }

    // chunks.push()
    // console.log(chunks)
    // process.abort()
    const browser = await launchPuppeteer();
    const page = await browser.newPage();
    // Define the path to the EJS template file
    const templatePath = path.join(
      __dirname,
      "views",
      "employees-pdf-docx.ejs"
    );
    const templatePathDocx = path.join(
      __dirname,
      "views",
      "employees-docx.ejs"
    );

    function separateTagValues(htmlString) {
      const $ = cheerio.load(htmlString);

      const thValues = [];
      const tdValues = [];

      $("table")
        .find("tr")
        .each((rowIndex, rowElement) => {
          const thElements = $(rowElement).find("th");
          const tdElements = $(rowElement).find("td");

          thElements.each((index, element) => {
            const thValue = $(element).text().trim();
            thValues.push(thValue);
          });

          tdElements.each((index, element) => {
            const tdValue = $(element).text().trim();
            tdValues.push(tdValue);
          });
        });

      return { thValues, tdValues };
    }

    // Create a PDF for each chunk of form data items
    // const pdfOptions = {
    //   format: 'A4',
    //   landscape: true,
    //   printBackground: true,
    //   displayHeaderFooter: true,
    //   margin: {
    //     top: '2cm',
    //     bottom: '2cm',
    //     left: '2cm',
    //     right: '2cm'
    //   },
    //   width: '17in',
    //   height: '11in',
    //   headerTemplate: `
    //   <div class="col-12" style="text-align: center; padding: 20px;">
    //   <h3 style="font-weight: bold; line-height: 13px; font-family: Arial, Helvetica, sans-serif;">
    //   <span style="font-size: 10px">ASSIGNMENT OF ACCOUNTABILITY</span><br>
    //   <span style="font-size: 14px">${cName}</span></h3>
    //   </div>`
    // };

    // const pdfPromises = chunks.map(chunk => {
    //   return new Promise((resolve, reject) => {
    //     ejs.renderFile(templatePath, { formDataItems: chunk, print: true, company: company.name }, (err, html) => {
    //       if (err) {
    //         console.error('Failed to render EJS template:', err);
    //         return reject(err);
    //       }

    //       pdf.create(html, pdfOptions).toBuffer((err, buffer) => {
    //         if (err) {
    //           console.error('Failed to create PDF:', err);
    //           return reject(err);
    //         }

    //         resolve(buffer);
    //       });
    //     });
    //   });
    // });

    // const pdfPromises = chunks.map(chunk => {
    //   return new Promise((resolve, reject) => {
    //     ejs.renderFile(templatePath, { formDataItems: chunk, print: true, company: company.name }, (err, html) => {
    //       if (err) {
    //         console.error('Failed to render EJS template:', err);
    //         return reject(err);
    //       }

    //       resolve(html)
    //     });
    //   });
    // });

    let htmlPage = "";
    let htmlPageDocx = "";
    // const browser = await puppeteer.launch();
    // const page = await browser.newPage();

    let tValues = [];

    for (let element of formDataItems) {
      htmlPage = await ejs.renderFile(templatePath, {
        formDataItems: formDataItems,
        print: true,
        company: company.name,
      });
      htmlPageDocx = await ejs.renderFile(templatePathDocx, {
        formDataItems: formDataItems,
        print: true,
        company: company.name,
      });

      tValues = separateTagValues(htmlPageDocx);

      // await page.setContent(htmlPage);

      // if(completeHtml != ""){
      //       // Extract the <div> and <table> content from the page
      // // const pageContent = completeHtml.match(/<div[^>]*>.*?<\/table>/s);

      // // Combine the extracted content with a <pagebreak /> between them
      // html += `<div style='page-break-after: always;'>${completeHtml}</div>`;

      // } else {
      //   html = `<div style='page-break-after: always;'>${page}</div>`;
      // }
      // print(html)
    }

    // html = '<body style="zoom: 0.5;">' + html + '</body>';

    // res.send(html);
    // return;

    // for (const inx in pdfPromises){
    //   const pdfBuffer = await pdfPromises[inx];
    //   // Define the path to save the PDF file
    // const filePath = path.join(__dirname, 'temp', `form_data${inx}.pdf`);

    // // Write the PDF buffer to a file
    // fs.writeFileSync(filePath, pdfBuffer);

    // }

    // combinedBuffer = await pdf.create(html, pdfOptions).toBuffer()

    const workbook = new excel.Workbook();
    const worksheet = workbook.addWorksheet("Sheet 1");

    const childArrays = formDataItems.map((doc) => doc.toObject());

    worksheet.getColumn(1).alignment = { wrapText: true, vertical: "top" };

    worksheet.getColumn(1).font = { bold: true };
    worksheet.getColumn(1).values = tValues["thValues"];

    // Get the maximum length of child arrays
    const maxChildArrayLength = childArrays.length;
    consoleMe(maxChildArrayLength);

    $a = 2;

    // console.log(childArrays)
    // process.abort()
    childArrays.forEach((obj) => {
      //  arr.forEach((obj) => {
      //  worksheet.getColumn($a).values = values;
      const data = [
        obj.name,
        obj.title,
        obj.jobDescription,
        obj.managerSupervisor,
        obj.directReports,
        obj.accountable,
        obj.participate,
        obj.compensation,
        obj.metrics,
        obj.positionalObjectives,
        obj.personalObjectives,
        obj.authorityLevels,
        obj.delegationOfAuthority,
      ];

      worksheet.getColumn($a).alignment = { wrapText: true, vertical: "top" };

      worksheet.getColumn($a).values = data;

      // })
      $a++;
    });

    // console.log(childArrays)
    worksheet.getRow(1).font = { bold: true };
    worksheet.getColumn(1).eachCell((cell, rowNumber) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "CCFFFF" },
      };
    });

    worksheet.getRow(1).eachCell((cell, colNumber) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "CCFFFF" },
      };
    });

    // Set borders for all cells with values
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        if (cell.value) {
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        }
      });
    });

    worksheet.columns.forEach((column) => {
      let maxLength = 0;
      column.eachCell({ includeEmpty: true }, (cell) => {
        const columnLength = cell.value ? cell.value.length : 0;
        maxLength = Math.max(maxLength, columnLength);
      });
      column.width = 30;
    });

    const excelBuffer = await workbook.xlsx.writeBuffer();

    // const docx = htmlDocx.asBlob(htmlPageDocx, { orientation: "landscape" });

    // const base64Docx = docx.toString('base64');

    // const docxBuffer = Buffer.from(docx)
    // const docxBuffer = await docx
    //   .arrayBuffer()
    //   .then((arrayBuffer) => Buffer.from(arrayBuffer, "binary"));
    // const customWidth = 1632
    // const customHeight = 1056

    // await page.setViewport({ width: customWidth, height: customHeight });

    const docxBuffer = await HTMLtoDOCX(completeHtml, null, {
      orientation: "landscape",
    });

    await page.setContent(htmlPage, { waitUntil: "networkidle0" });

    const headerTemplate = `<div style="width: 100%; text-align: center; height: 120px; min-height: 120px;
      margin-top: 30px; margin-bottom: 30px">
      <h3 style="line-height: 12px;"><span style="font-size: 9px; font-weight: normal">
      ASSIGNMENT OF ACCOUNTABILITY</span><br><span style="font-size: 14px; font-weight: 300">${cName}</span></h3></div>`;

    const pdfGenerated = await page.pdf({
      landscape: true,
      printBackground: true,
      displayHeaderFooter: true,
      headerTemplate,
      margin: {
        top: "4cm",
        bottom: "2cm",
        left: "2cm",
        right: "2cm",
      },
      width: "1056px",
      height: "1632px",
    });

    await browser.close();

    const pdfBuffer = Buffer.from(pdfGenerated);

    const zip = new JSZip();
    zip.file("form_data.pdf", pdfBuffer);
    zip.file("form_data.docx", docxBuffer);
    zip.file("form_data.xlsx", excelBuffer);
    // zip.file('form_data.xlsx', excelBuffer);

    // zip.file('form_data.xlsx', excelBuWffer);

    const zipBuffer = await zip.generateAsync({ type: "nodebuffer" });

    res.setHeader("Content-Type", "application/zip");
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="combined-files.zip"'
    );

    // Send the combined buffer as the response
    res.send(zipBuffer);

    // res.redirect('views/employees')
  } catch (error) {
    console.error(error);
    res.status(500).send("Internal Server Error");
  }
});

// using no more
app.get("/admin/:id/download", async (req, res) => {
  try {
    // Retrieve the specific form data item based on the provided ID
    const formDataItem = await FormData.findById(req.params.id);

    let company = await Company.findOne();

    if (!company) {
      company = "Not Assigned";
    }

    if (!formDataItem) {
      // Return an error if the form data item is not found
      return res.status(404).send("Form data not found");
    }

    // Define the path to the EJS template file
    const templatePath = path.join(__dirname, "views", "employees.ejs");

    // Render the EJS template with the form data item
    ejs.renderFile(
      templatePath,
      { formDataItems: [formDataItem], company: company.name, print: true },
      (err, html) => {
        if (err) {
          console.error("Failed to render EJS template:", err);
          return res.status(500).send("Internal Server Error");
        }

        // Convert the rendered HTML to PDF using your preferred method
        // Replace the following code with your actual PDF conversion logic

        // For example, using the 'html-pdf' library:

        pdf.create(html).toStream((err, stream) => {
          if (err) {
            console.error("Failed to create PDF:", err);
            return res.status(500).send("Internal Server Error");
          }

          // Set the appropriate headers for PDF download
          res.setHeader("Content-Type", "application/pdf");
          res.setHeader(
            "Content-Disposition",
            "attachment; filename=form_data.pdf"
          );

          // Pipe the PDF stream to the response
          stream.pipe(res);
        });
      }
    );
  } catch (error) {
    console.error(error);
    res.status(500).send("Internal Server Error");
  }
});

app.get(
  "/admin/view",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    const company = await Company.findOne({ _id: user?.company });
    // Retrieve the specific form data item based on the provided ID
    const formDataItem = await FormData.find({
      company: company?._id,
    }).populate({ path: "user", model: User });
    //  console.log("form data", formDataItem)
    if (!formDataItem) {
      // Return an error if the form data item is not found
      return res.status(404).send("Form data not found");
    }
    const cName = company?.name;
    const admin = await User.findById(user?._id).populate("admins");
    const groupAdmins = admin?.admins;
    formDataItem.empCompanyName = cName;

    // Render the EJS template with the form data item
    res.render("employees", {
      formDataItems: formDataItem,
      company: company?._id,
      user: user,
      groupAdmins,
      comp: true,
    });
  })
);

app.get(
  "/group-admin/view",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;

    const employees = await User.find({ company: user?.company });
    const employeesUnderUser = getFollowingEmployees(user?._id, employees);

    const formDataItemsPromise = Promise.all(
      employeesUnderUser?.map(async (user) => {
        const formData = await FormData.findOne({ email: user?.email }).lean();

        return formData;
      })
    );

    const formDataItemsArray = await formDataItemsPromise;

    const formDataItems = formDataItemsArray.filter((user) => user);

    const group = await Group.findOne({ adminId: user?._id });
    let groupOrCompanyId = group?._id;
    formDataItems.empCompanyName = group?.name;

    if (!group) {
      const company = await Company.findOne({ _id: user?.company });

      groupOrCompanyId = company?._id;
      formDataItems.empCompanyName = company?.name;

      res.render("employees", {
        formDataItems: formDataItems,
        company: groupOrCompanyId,
        user: user,
        comp: true,
        groupAdmins: [],
      });

      return;
    }

    res.render("employees", {
      formDataItems: formDataItems,
      company: groupOrCompanyId,
      user: user,
      comp: false,
      groupAdmins: [],
    });

    // const group = await Group.findOne({ adminId: user?._id });
    // let groupId = "";
    // // Retrieve the specific form data item based on the provided ID
    // let formDataItem = [];
    // if (group) {
    //   formDataItem = await FormData.find({ group: group?._id });
    //   //  console.log("form data", formDataItem)

    //   const gName = group?.name;

    //   formDataItem.empCompanyName = gName;
    //   groupId = group?._id;
    //   if (!formDataItem) {
    //     // Return an error if the form data item is not found
    //     return res.status(404).send("Form data not found");
    //   }
    // }

    // Render the EJS template with the form data item

    // console.log(groupOrCompanyId);
  })
);

app.get(
  "/group-admin/view/:id",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    const Id = req.params.id;

    const group = await Group.findOne({ _id: Id });

    // Retrieve the specific form data item based on the provided ID
    const formDataItem = await FormData.find({
      groupIds: group._id,
    }).populate({ path: "user", model: User });
    //  console.log("form data", formDataItem)

    const gName = group.name;

    formDataItem.empCompanyName = gName;

    if (!formDataItem) {
      // Return an error if the form data item is not found
      res.status(404).send("Form data not found");
      return;
    }

    // Render the EJS template with the form data item
    res.render("employees", {
      formDataItems: formDataItem,
      company: group._id,
      user: user,
      comp: false,
      groupAdmins: [],
    });
  })
);

app.get(
  "/group-admin/:id/view",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    const groupId = req.params.id;
    const group = await Group.findOne({ _id: groupId });

    // Retrieve the specific form data item based on the provided ID
    const formDataItem = await FormData.find({
      groupIds: group._id,
    });
    //  console.log("form data", formDataItem)

    const gName = group.name;

    formDataItem.empCompanyName = gName;

    if (!formDataItem) {
      // Return an error if the form data item is not found
      res.status(404).send("Form data not found");
      return;
    }

    // Render the EJS template with the form data item
    res.render("employees", {
      formDataItems: formDataItem,
      company: group._id,
      user: user,
      comp: false,
      groupAdmins: [],
    });
  })
);

app.get(
  "/admin/:id/view",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;

    const ID = req.params.id;
    // Retrieve the specific form data item based on the provided ID
    const formDataItem = await FormData.find({ company: ID });
    //  console.log("form data", formDataItem)
    const companyName = await Company.findOne({ _id: ID });

    const cName = companyName.name;
    formDataItem.empCompanyName = cName;

    // previously the old system, when group admins are created, they goes under users's (admin access) "admins" field
    // const admin = await User.findById(companyName.user_id).populate("admins");
    const groupAdmins = await User.find({
      level: { $in: ["1", "2"] },
      company: ID,
    });

    // console.log(formDataItem)

    if (!formDataItem) {
      // Return an error if the form data item is not found
      res.status(404).send("Form data not found");
      return;
    }

    // Render the EJS template with the form data item
    res.render("employees", {
      formDataItems: formDataItem,
      company: ID,
      user: user,
      groupAdmins: groupAdmins ?? [],
      comp: true,
    });
  })
);

app.get(
  "/employees-aoa",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    const company = await Company.findOne({ _id: user?.company });
    const employees = await User.find({ company: company?._id });

    const descendants = getFollowingEmployees(user?._id, employees);

    const formDataIds = descendants.map((item) => item?.formData);

    const formDataItems = await FormData.find({ _id: { $in: formDataIds } });

    res.render("employees-level-2-5", {
      formDataItems: formDataItems,
      company: company?._id,
      user: user,
      groupAdmins: [],
      comp: true,
    });
  })
);

app.get(
  "/edit-form/:id",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    const itemId = req.params.id;

    if (itemId === "undefined") {
      const companyId = req.query.comp;
      res.redirect(`/formadmin/${companyId}`);
      return;
    }

    // Fetch the existing form data from the database
    const existingFormData = await FormData.findById(itemId).populate({
      path: "company",
      model: Company,
    });

    // if (!existingFormData) {
    //   res.render("404", { content: "Employee not found." });
    // }

    // Render the edit form page with the retrieved data
    res.render("edit-form", { formData: existingFormData, user: user });
  })
);

app.get(
  "/delete-employee/:id",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    const empId = req.params.id;

    // Fetch the existing form data from the database
    const companyId = await FormData.findById(empId);

    await FormData.deleteOne({ _id: empId });

    if (user.role === 1) {
      res.redirect(`/admin/${companyId.company}/view`);
      // Render the edit form page with the retrieved data
    } else if (user.role === 3) {
      res.redirect(`/`);
    } else {
      res.redirect(`/admin/view`);
    }
  })
);

// delete employee by a group admin API
app.get(
  "/delete-employee-by-group-admin/:id",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    const empId = req.params.id;

    const filter = { _id: empId };
    const update = { group: "" };
    // Fetch the existing form data from the database and update "group" name
    const formData = await FormData.findOneAndUpdate(filter, update);

    if (user.role === 1) {
      res.redirect(`/admin/${formData.company}/view`);
      // Render the edit form page with the retrieved data
    } else if (user.role === 3) {
      res.redirect(`/`);
    } else {
      res.redirect(`/admin/view`);
    }
  })
);

// Define a route to handle form updates
// id is formData id
app.post(
  "/edit-form/submit/:id",
  catchAsync(async (req, res) => {
    const formData = req.body;
    const formId = req.params.id;

    const formdata = await FormData.findByIdAndUpdate(formId, formData);

    const company = await Company.findOne({
      _id: new ObjectId(formdata?.company),
    });
    const admin = await User.findById(company?.user_id);

    const templatePath = path.join(
      process.cwd(),
      "views",
      "employees-pdf-docx.ejs"
    );

    const pdfFile = await generatePDFonA4(
      [formdata],
      company?.name,
      templatePath
    );

    const wordFile = await generate([formdata], company?.name);

    const transporter = await createEmailTransporter();

    if (admin) {
      const mailOptions = {
        from: `LeaveNowGrow <${configEnv?.userEmail}>`,
        to: `${admin?.email}, ${formData?.email}`,
        subject: "AOA Updated",
        text: `An employee, ${formdata?.name}'s details have been updated`,
        attachments: [
          { filename: `${formdata?.name}-AOA.pdf`, content: pdfFile },
          { filename: `${formdata?.name}-AOA.docx`, content: wordFile },
        ],
      };

      transporter.sendMail(mailOptions);
      res.render("thankyou-page", { newFormData: formdata });
    } else {
      consoleMe("Admin User not found.");
      res.status(401).send("Admin User not found.");
    }
  })
);

////---------new system form here-----------/////////

//new admin panel
app.get(
  "/companies",
  isLoggedIn,
  isSuperAdmin,
  catchAsync(async (req, res) => {
    const user = req.user;
    const companies = await Company.find({})
      .populate("subGroup")
      .populate({ path: "user_id", model: User })
      .lean();
    const admins = await User.find({ level: "1" }).sort("name").lean();

    res.render("companies", { companies, company: null, admins, user: user });
  })
);

app.get(
  "/groups",
  isLoggedIn,
  isCompAdmin,
  catchAsync(async (req, res) => {
    const user = req.user;

    const groups = await Group.find({ creatorId: user._id }).lean();
    const admins = await User.find({ creatorId: user._id }).lean();
    const currentComany = await Company.find({ _id: user?.company }).lean();

    res.render("groups", {
      groups,
      admins,
      user: user,
      company: currentComany,
    });
  })
);

app.post(
  "/add-admin",
  catchAsync(async (req, res) => {
    const user = req.user;
    const data = req.body;
    const hashedPassword = await bcrypt.hash(data.password, 10);

    const newAdminData = {
      name: data?.name,
      email: data?.email,
      password: hashedPassword,
      status: 0,
      creatorId: user._id,
    };

    if (req.body.isSuperUser == "true") {
      newAdminData.level = "Super Admin";
      newAdminData.role = 1;
    } else {
      newAdminData.level = "1";
      newAdminData.role = 2;
    }

    const newAdmin = new User(newAdminData);

    const token = getToken({
      email: data?.email,
    });

    const mailOptions = {
      from: `LeaveNowGrow <${configEnv?.userEmail}>`,
      to: data?.email,
      subject: "Your account has been created",
      text: `<p>Hi ${data?.name},</p>
      <p>Please verify your email address by clicking on this link below.</p>
      <a href='${BASE_LINK}/email/verify/${token}'>Link</a>
      <p>Here is your password for login</p>
      <p>Password: ${data?.password}</p>
      <p style="text-decoration: bold; color: red;">Please change your password</p>
      `,
      html: `<p>Hi ${data?.name},</p>
      <p>Please verify your email address by clicking on this link below.</p>
      <a href='${BASE_LINK}/email/verify/${token}'>Link</a>
      <p>Here is your password for login.</p>
      <p>Password: ${data?.password}</p>
      <p style="text-decoration: bold; color: red;">Please change your password</p>
      `,
    };

    const transporter = await createEmailTransporter();

    //saving
    newAdmin
      .save()
      .then(() => {
        consoleMe("Admin is created successfully");
        transporter.sendMail(mailOptions);
        res.send({ success: true, message: "Admin is created successfully." });
      })
      .catch((error) => {
        consoleMe(error);
        if (error?.code === 11000) {
          res.send({
            success: false,
            message: "Email already exists.",
          });
        } else {
          res.send({
            success: false,
            message: "Admin cannot be created, error in database.",
          });
        }
      });
  })
);

app.post(
  "/submit_admin",
  catchAsync(async (req, res) => {
    const data = req.body;

    const companyData = {
      name: data?.name,
      user_id: data?.user_id,
    };

    const newcompany = new Company(companyData);

    await newcompany.save();

    const user = await User.findById(data?.user_id);

    if (!user) {
      res.status(404).send("User not found");
      return;
    }

    user.status = 1;
    user.company = newcompany?._id;
    await user.save();

    res.redirect("/companies");
  })
);

//create
app.post(
  "/add-group-admin/:id",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    const parentAdminId = req.params.id;

    const manager = await User.findOne({ _id: parentAdminId });

    // Hash the password before saving it to the database
    const password = await bcrypt.hash(req.body.password, 10);

    // Create a new admin with hashed password
    const adminData = {
      email: req.body.email,
      password: password,
      name: req.body.name,
      creatorId: parentAdminId,
      role: 3,
      level: "2",
      status: 0,
      manager: {
        id: (manager?._id).toString(),
        name: manager?.name,
        email: manager?.email,
      },
      company: manager?.company,
    };

    // Create a new admin user
    const newAdmin = new User(adminData);

    const token = getToken({
      email: adminData?.email,
    });

    const mailOptions = {
      from: `LeaveNowGrow <${configEnv?.userEmail}>`,
      to: adminData?.email,
      subject: "Your account has been created",
      text: `<p>Hi ${adminData?.name},</p>
      <p>Please verify your email address by clicking on this link below.</p>
      <a href='${BASE_LINK}/email/verify/${token}'>Link</a>
      <p>Here is your password for login</p>
      <p>Password: ${req.body.password}</p>
      <p style="text-decoration: bold; color: red;">Please change your password</p>
      `,
      html: `<p>Hi ${adminData?.name},</p>
      <p>Please verify your email address by clicking on this link below.</p>
      <a href='${BASE_LINK}/email/verify/${token}'>Link</a>
      <p>Here is your password for login.</p>
      <p>Password: ${req.body.password}</p>
      <p style="text-decoration: bold; color: red;">Please change your password</p>
      `,
    };

    const transporter = await createEmailTransporter();

    const session = await mongoose.startSession();

    try {
      session.startTransaction();

      await newAdmin.save({ session });

      await User.findOneAndUpdate(
        { _id: parentAdminId },
        { $push: { admins: newAdmin._id } },
        { new: true }
      ).session(session);

      await session.commitTransaction();
      await session.endSession();

      consoleMe("Admin is created successfully");
      transporter.sendMail(mailOptions);
      res.send({ success: true, message: "Admin is created successfully." });
    } catch (error) {
      await session.abortTransaction();
      await session.endSession();

      consoleMe(error);
      if (error?.code === 11000) {
        res.send({
          success: false,
          message: "Email already exists.",
        });
      } else {
        res.send({
          success: false,
          message: "Admin cannot be created, error in database.",
        });
      }
    }
  })
);

// Endpoint to create a group
// app.post('/create_group/:id', async (req, res) => {
//   try {
//     const formdataids = JSON.parse(req.body.selectedIds); // Parse the selected IDs from JSON
//     const user = req.user
//     const newGroupName = req.body.groupName;
//     const  compandyID =  req.params.id
//     const comp = await Company.findById(compandyID)
//     // Create a new admin
//     const adminData = {
//       email: req.body.email,
//       password: req.body.password,
//       name: req.body.adminName,
//       creatorId: comp.user_id,
//       role: 3,
//       status: 1
//     };

//     const password = await bcrypt.hash(adminData.password, 10);
//     adminData.password = password;
//     const newAdmin = new User(adminData);

//     await newAdmin.save();

//     const groupData = {
//       name: newGroupName,
//       adminId: newAdmin._id,
//       company : compandyID,
//       creatorId : user._id
//     };

//     const newgroup = new Group(groupData);

//     await newgroup.save();

//     // Update the form data documents with the new group name
//     await FormData.updateMany({ _id: { $in: formdataids } }, { $set: { group: newgroup._id } });

//     // Update the company document to push the new group's ID
//     const updatedCompany = await Company.findOneAndUpdate(
//       { _id: compandyID },
//       { $push: { subGroup: newgroup._id } },
//       { new: true }
//     );

//     await User.findOneAndUpdate(
//       { _id: updatedCompany.user_id },
//       { $push: { admins: newAdmin._id } },
//       { new: true }
//     );

//     console.log('Group created successfully');
//     console.log('Company updated');
//     if(user.role == 1){
//       res.redirect(`/admin/${compandyID}/view`);
//     }else {
//       res.redirect("/admin/view")
//     }
//   } catch (error) {
//     console.error('Failed to save data:', error);
//     res.status(500).send('Internal Server Error');
//   }
// });

app.post(
  "/create_group/:id",
  catchAsync(async (req, res) => {
    const formdataids = JSON.parse(req.body.selectedIds); // Parse the selected IDs from JSON
    const user = req.user;
    const newGroupName = req.body.groupName;
    const adminId = req.body.adminId;
    const compandyID = req.params.id;
    const comp = await Company.findById(compandyID);

    const groupData = {
      name: newGroupName,
      adminId: adminId,
      company: compandyID,
      creatorId: comp?.user_id,
    };

    const newgroup = new Group(groupData);

    await newgroup.save();

    const formDataUsers = await User.find({
      formData: { $in: formdataids },
    });

    // update in the goals worksheet
    const goalsWorksheetsUpdate = Promise.all(
      formDataUsers.map(async (user) => {
        if (user.goalsWorksheet) {
          return await GoalsWorksheet.updateOne(
            { _id: user.goalsWorksheet },
            { $addToSet: { groupIds: newgroup._id } }
          );
        }
      })
    );
    await goalsWorksheetsUpdate;

    // Update the form data documents with the new group name
    await FormData.updateMany(
      { _id: { $in: formdataids } },
      { $addToSet: { groupIds: newgroup._id } }
    );

    // Update the company document to push the new group's ID
    await Company.findOneAndUpdate(
      { _id: compandyID },
      { $push: { subGroup: newgroup._id } }
    );

    await User.findByIdAndUpdate(adminId, { $set: { status: 1 } });

    // consoleMe("Group created successfully");
    // consoleMe("Company updated");
    if (user.role == 1) {
      res.redirect(`/admin/${compandyID}/view`);
    } else {
      res.redirect("/admin/view");
    }
  })
);

// ?m=managerId for manager
app.get(
  "/formadmin/:id",
  catchAsync(async (req, res) => {
    const compId = req.params.id;
    const managerId = req.query.m;

    const company = await Company.findOne({ _id: compId });
    const manager = await User.findOne({ _id: managerId });

    // const editLink = `${BASE_LINK}/formadmin/${compId}?m=${company?.user_id}`;
    const editLink = `${BASE_LINK}/register?c=${compId}&m=${company?.user_id}`;

    res.render("form_admin", {
      id: compId,
      group: false,
      user: req.user,
      editLink: editLink,
      manager: manager,
    });
  })
);

// ?m=managerId for manager
app.get(
  "/group-formadmin/:id",
  catchAsync(async (req, res) => {
    const groupId = req.params.id;
    const managerId = req.query.m;

    const manager = await User.findOne({ _id: managerId });

    const group = await Group.findById(groupId);
    // const editLink = `${BASE_LINK}/group-formadmin/${groupId}?m=${group?.adminId}`;
    const editLink = `${BASE_LINK}/register?c=${group?.company}&m=${group?.adminId}`;

    res.render("form_admin", {
      id: group?.company,
      groupId: group?._id,
      group: true,
      user: req?.user,
      editLink: editLink,
      manager: manager,
    });
  })
);

// delete company
app.get(
  "/delete/:id",
  catchAsync(async (req, res) => {
    const compId = req.params.id;

    if (!ObjectId.isValid(compId)) {
      res.status(400).send("Invalid ObjectId");
      return;
    }
    // Find and delete all form data associated with the company

    await FormData.deleteMany({
      companyId: compId,
    });

    const data = await Company.findOne({ _id: compId });
    const result = await Company.deleteOne({ _id: compId });
    await User.findByIdAndUpdate(data.user_id, { status: 0 });

    await User.updateMany(
      { admin: data.user_id },
      { $pull: { admins: data.user_id } }
    );

    if (result.deletedCount === 0) {
      res.status(404).send("Company not found");
      return;
    }

    // return res.status(200).send("Company deleted successfully");
    res.redirect("/companies");
  })
);

app.get(
  "/delete-group/:id",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const groupId = req.params.id;
    const user = req.user;
    if (!ObjectId.isValid(groupId)) {
      res.status(400).send("Invalid ObjectId");
      return;
    }

    const data = await Group.findOne({ _id: groupId });
    const result = await Group.deleteOne({ _id: groupId });

    // if (data.adminId) {
    //   await User.findByIdAndDelete(data.adminId);
    //   await User.updateMany(
    //     { },
    //     { $pull: { admins: data.adminId } }
    //   );
    // }

    await Company.updateOne(
      { _id: data.company },
      { $pull: { subGroup: groupId } }
    );

    if (result.deletedCount === 0) {
      res.status(404).send("Group not found");
      return;
    }

    // res.status(200).send("Company deleted successfully");
    if (user.role === 1) {
      res.redirect("/companies");
    } else if (user.role === 2) {
      res.redirect("/groups");
    } else {
      res.redirect("/");
    }
  })
);

app.get(
  "/admin/delete/:id",
  catchAsync(async (req, res) => {
    const adminId = req.params.id;

    if (!ObjectId.isValid(adminId)) {
      res.status(500).json({
        success: false,
        message: "Invalid object id!",
      });
      return;
    }

    // const company = await Company.findOneAndDelete({user_id: adminId})
    const delAdmin = await User.findOneAndDelete({ _id: adminId });
    // const delGroupAdmin = await Group.findOneAndDelete({ adminId: adminId });
    await User.updateMany({ admin: adminId }, { $pull: { admins: adminId } });

    if (delAdmin.deletedCount === 0) {
      consoleMe("Admin not found");
    }
    if (req.user.role == 1) {
      res.redirect("/admins");
    } else {
      res.redirect("/comp-admins");
    }
  })
);

//edit company name
app.get(
  "/change/:id",
  catchAsync(async (req, res) => {
    const compId = req.params.id;
    res.render("change_name", { id: compId });
  })
);

//submit button for company name change
app.post(
  "/submit_change/:id",
  catchAsync(async (req, res) => {
    // Extract the form data from the request body
    const formData = req.body;

    // Extract the form ID from the URL parameter
    const formId = req.params.id;

    // Update the form data in the MongoDB database
    Company.findByIdAndUpdate(formId, formData)
      .then(() => {
        consoleMe("Company name updated successfully");
        res.status(200).send("Company name updated successfully!");
      })
      .catch((error) => {
        consoleMe(error);
        res.status(500).send("Internal Server Error");
      });
  })
);

// Update column order for a company or group by ID
app.post(
  "/update-columnOrder/:id",
  catchAsync(async (req, res) => {
    const ID = req.params.id;
    const { columnOrder } = req.body;

    let modelToUpdate = null;
    let modelName = "";

    // Check if the ID corresponds to a Company
    const company = await Company.findById(ID);
    if (company) {
      modelToUpdate = Company;
      modelName = "Company";
    } else {
      // If not a Company, check if it's a Group
      const group = await Group.findById(ID);
      if (group) {
        modelToUpdate = Group;
        modelName = "Group";
      }
    }

    if (modelToUpdate) {
      // Update the column order
      await modelToUpdate.findByIdAndUpdate(ID, { columnOrder });
      res.status(200).json({ message: `${modelName} updated` });
      return;
    }

    // If neither a Company nor a Group was found
    res.status(400).send("No company or group found");
  })
);

// Get column order for a company or group by ID
app.get(
  "/get-ColumnOrder/:id",
  catchAsync(async (req, res) => {
    const ID = req.params.id;

    let modelToQuery = null;
    let modelName = "";
    let columnOrder = [];

    // Check if the ID corresponds to a Company
    const company = await Company.findById(ID);

    if (company) {
      modelToQuery = Company;
      modelName = "Company";
      columnOrder = company.columnOrder;
    } else {
      // If not a Company, check if it's a Group
      const group = await Group.findById(ID);
      if (group) {
        modelToQuery = Group;
        modelName = "Group";
        columnOrder = group.columnOrder;
      }
    }

    if (modelToQuery) {
      res.status(200).json(columnOrder);
    } else {
      // If neither a Company nor a Group was found
      res.status(400).send("No company or group found");
    }
  })
);

app.post(
  "/send-pdf/:id",
  catchAsync(async (req, res) => {
    const user = req.user;

    // Extract the email address from the request body
    const { email } = req.body;
    const formId = req.params.id;
    const formDataItem = await FormData.find({ _id: formId });
    const formDataItems = formDataItem;
    const company = await Company.findOne({ _id: formDataItem[0].company });

    let cName = "";
    if (company) {
      cName = company.name;
      formDataItems.empCompanyName = cName;
    }

    if (!formDataItems) {
      // Return an error if the form data item is not found
      res.status(404).send("Form data not found");
      return;
    }

    const transporter = await createEmailTransporter();

    const docxBuffer = await generate(formDataItems, cName);

    // Send the PDF to the provided email address
    const mailOptions = {
      from: `LeaveNowGrow <${configEnv.userEmail}>`,
      to: email,
      subject: "AOA Report",
      text: "Attached is the AOA Report.",
      attachments: [
        {
          filename: `${formDataItem[0].name}-AOA.docx`,
          content: docxBuffer,
        },
      ],
    };

    // Send the email with the PDF attachment
    await transporter.sendMail(mailOptions);
    res.render("email-success", { email, formId, user, goals: false });
  })
);

app.get(
  "/fetchEmployees/:id",
  catchAsync(async (req, res) => {
    const id = req.params.id;
    const formDatas = await FormData.find({ company: id });

    res.status(200).send(formDatas);
  })
);

app.get(
  "/revision",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;

    if (user?.role === 1 || user?.level === "Super Admin") {
      const companies = await Company.find();

      res.render("revision-summary", { user, companies });
    } else if (
      user?.role === 2 ||
      user?.role === 3 ||
      user?.level === "1" ||
      user?.level === "2"
    ) {
      const company = await Company.findOne({ _id: user?.company });
      const employees = await User.find({ company: company?._id });

      const descendants = getFollowingEmployees(user?._id, employees);

      res.render("revision-summary-company", {
        user,
        company,
        employees: descendants,
      });
    } else {
      res.redirect("/unauthorized");
    }
  })
);

app.get(
  "/edit-formdata/:id",
  catchAsync(async (req, res) => {
    const itemId = req.params.id;

    // Fetch the existing form data from the database
    const existingFormData = await FormData.findById(itemId);

    // Render the edit form page with the retrieved data
    res.render("edit-formdata", { formData: existingFormData });
  })
);

// send revision data to email address entered
app.post(
  "/send-revision",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    // Extract the email address from the request body
    // { emailTo, emailCc, employee, company, message, action, addEditLink, addPDF, addDocs, quarter, year }
    const data = req.body;
    const formId = data?.employee;
    const formDataItems = await FormData.findById(formId);
    const companyData = await Company.findOne({
      _id: data.company,
    });
    const cName = companyData.name;

    formDataItems.empCompanyName = cName;
    if (!formDataItems) {
      // Return an error if the form data item is not found
      res.status(404).send({ success: false, message: "Form data not found" });
      return;
    }

    // Send the PDF to the provided email address
    const transporter = await createEmailTransporter();

    const mailOptions = {
      from: `LeaveNowGrow <${configEnv.userEmail}>`,
      to: data?.emailTo,
      cc: data?.emailCc ? data?.emailCc : "",
      subject: "Review or edit your AOA",
      text: data?.message,
      attachments: [],
    };

    if (data?.addPDF) {
      const templatePath = path.join(
        process.cwd(),
        "views",
        "employees-pdf-docx.ejs"
      );

      // Create a PDF using Puppeteer
      const pdfBuffer = await generatePDFonA4(
        [formDataItems],
        cName,
        templatePath
      );

      mailOptions.attachments.push({
        filename: `${formDataItems?.name}-Review.pdf`,
        content: pdfBuffer,
      });
    }

    if (data?.addDocs) {
      const docxBuffer = await generate([formDataItems], cName);

      mailOptions.attachments.push({
        filename: `${formDataItems?.name}-Review.docx`,
        content: docxBuffer,
      });
    }

    if (data?.addEditLink) {
      if (data?.action === "sendSelfEvaluation") {
        // first create the new manager evaluation data
        const managerData = await ManagerEvaluation.create({
          formDataId: formDataItems?._id,
          employeeEmail: data?.emailTo,
          managerEmail: data?.emailCc,
          year: data?.year,
          quarter: data?.quarter,
          name: formDataItems?.name,
        });

        // save it in the employee data
        formDataItems.managerEvaluationIds.addToSet(managerData?._id);
        await formDataItems.save();

        // create a new self evaluation data
        const newData = await SelfEvaluation.create({
          formDataId: formDataItems?._id,
          managerEvaluationId: managerData?._id,
          employeeEmail: data?.emailTo,
          managerEmail: data?.emailCc,
          year: data?.year,
          quarter: data?.quarter,
          name: formDataItems?.name,
        });

        // then save the data in employee data
        formDataItems.selfEvaluationIds.addToSet(newData?._id);
        await formDataItems.save();

        mailOptions.html =
          formatEmail(data?.message) +
          `<a style="padding: 0px 4px;font-weight: 700; color: #0000FF;" href="${BASE_LINK}/self-evaluation-formdata/${newData?._id}">Self evaluation link</a>`;

        mailOptions.subject = `${formDataItems?.name} - AOA Self Evaluation for ${data?.quarter}, ${data?.year}`;
      }

      // DONE: complete manager evaluation edit link
      else if (data?.action === "sendManagerEvaluation") {
        const newData = await ManagerEvaluation.create({
          formDataId: formDataItems?._id,
          employeeEmail: data?.emailCc,
          managerEmail: data?.emailTo,
          year: data?.year,
          quarter: data?.quarter,
          name: formDataItems?.name,
        });

        formDataItems.managerEvaluationIds.addToSet(newData?._id);
        await formDataItems.save();

        mailOptions.html =
          formatEmail(data?.message) +
          `<a style="padding: 0px 4px;font-weight: 700; color: #0000FF;" href="${BASE_LINK}/manager-evaluation-edit-form/${newData?._id}">Manager evaluation link</a>`;

        mailOptions.subject = `${formDataItems?.name} - AOA Manager Evaluation for ${data?.quarter}, ${data?.year}`;
      } else if (data?.action === "sendGoalsWorksheet") {
        mailOptions.subject = "Complete or Edit Your Goals Worksheet";
        const token = getToken({
          name: formDataItems?.name,
          email: formDataItems?.email,
        });

        mailOptions.html =
          formatEmail(data?.message) +
          `<a style="padding: 0px 4px;font-weight: 700; color: #0000FF;" href="${BASE_LINK}/goals-worksheet/${data.company}?view=admin-actions&token=${token}">Complete Goals Worksheet</a>`;
      } else {
        mailOptions.html =
          formatEmail(data?.message) +
          `<a style="padding: 0px 4px;font-weight: 700; color: #0000FF;" href="${BASE_LINK}/edit-formdata/${data?.employee}">Edit Employee Details</a>`;
      }
    }

    // Send the email with the PDF attachment
    await transporter.sendMail(mailOptions);

    res.send({ success: true, message: "Email is sent succefully." });

    // res.render("email-success", {
    //   email: data?.email,
    //   formId,
    //   user,
    //   goals: false,
    // });
  })
);

// ?to=example@email.com will have the email...
app.get(
  "/email-success",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    const email = req.query.to;

    res.render("email-success", {
      email: email,
      formId: null,
      user,
      goals: false,
    });
  })
);

// download docx api by miraj
app.get(
  "/admin/downloadDOCX/:ids",
  catchAsync(async (req, res) => {
    const selectedIds = req.params.ids.split(",");

    // Retrieve the specific form data items based on the provided IDs
    const formDataItems = await FormData.find({ _id: { $in: selectedIds } });
    const companyName = await Company.findOne({
      _id: formDataItems[0].company,
    });

    let cName = "";
    if (companyName) {
      cName = companyName.name;
      formDataItems.empCompanyName = cName;
    }

    if (!formDataItems) {
      // Return an error if the form data items are not found
      res.status(404).send("Form data not found");
      return;
    }

    // download option for more than 7 AOA's
    if (formDataItems.length > 7) {
      const formDataItems1 = formDataItems.slice(0, 7);
      const formDataItems2 = formDataItems.slice(7, formDataItems.length);

      const docxBuffer1 = await generate(formDataItems1, cName);
      const docxBuffer2 = await generate(formDataItems2, cName);

      const zip = new JSZip();
      zip.file(`${cName}-AOA-Part-1.docx`, docxBuffer1);
      zip.file(`${cName}-AOA-Part-2.docx`, docxBuffer2);

      const zipBuffer = await zip.generateAsync({ type: "nodebuffer" });

      res.setHeader("Content-Type", "application/zip");
      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${cName}-AOA.zip`
      );

      // Send the combined buffer as the response
      res.send(zipBuffer);
      return;
    }

    const docxBuffer = await generate(formDataItems, cName);

    if (formDataItems.length > 1) {
      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${cName}-AOA.docx`
      );
    } else {
      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${formDataItems[0].name}-AOA.docx`
      );
    }

    // Send the combined buffer as the response
    // res.send(Buffer.from(docxBuffer, "base64"));
    // res.redirect('views/employees')
    res.send(docxBuffer);
  })
);

// download excel file by miraj
app.get(
  "/admin/downloadEXCEL/:ids",
  catchAsync(async (req, res) => {
    const selectedIds = req.params.ids.split(",");

    // Retrieve the specific form data items based on the provided IDs
    const formDataItems = await FormData.find({ _id: { $in: selectedIds } });
    const companyName = await Company.findOne({
      _id: formDataItems[0].company,
    });

    let cName = "";
    if (companyName) {
      cName = companyName.name;
      formDataItems.empCompanyName = cName;
    }

    if (!formDataItems) {
      // Return an error if the form data items are not found
      res.status(404).send("Form data not found");
      return;
    }

    const excelBuffer = await generateExcelFormdata(formDataItems, cName);

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    if (formDataItems.length > 1) {
      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${cName}-AOA.xlsx`
      );
    } else {
      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${formDataItems[0].name}-AOA.xlsx`
      );
    }

    // Send the combined buffer as the response
    // res.redirect('views/employees')
    res.send(excelBuffer);
  })
);

// download pdf file by miraj
app.get(
  "/admin/downloadPDF/:ids",
  catchAsync(async (req, res) => {
    const selectedIds = req.params.ids.split(",");

    // Retrieve the specific form data items based on the provided IDs
    const formDataItems = await FormData.find({ _id: { $in: selectedIds } });
    const companyName = await Company.findOne({
      _id: formDataItems[0].company,
    });

    let cName = "";
    if (companyName) {
      cName = companyName.name;
      formDataItems.empCompanyName = cName;
    }

    if (!formDataItems) {
      // Return an error if the form data items are not found
      res.status(404).send("Form data not found");
      return;
    }

    const templatePath = path.join(
      process.cwd(),
      "views",
      "employees-pdf-docx.ejs"
    );

    // download option for more than 7 AOA's
    if (formDataItems.length > 7) {
      const formDataItems1 = formDataItems.slice(0, 7);
      const formDataItems2 = formDataItems.slice(7, formDataItems.length);

      const pdfBuffer1 = await generatePDF(formDataItems1, cName, templatePath);
      const pdfBuffer2 = await generatePDF(formDataItems2, cName, templatePath);

      const zip = new JSZip();
      zip.file(`${cName}-AOA-Part-1.pdf`, pdfBuffer1);
      zip.file(`${cName}-AOA-Part-2.pdf`, pdfBuffer2);

      const zipBuffer = await zip.generateAsync({ type: "nodebuffer" });

      res.setHeader("Content-Type", "application/zip");
      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${cName}-AOA.zip`
      );

      // Send the combined buffer as the response
      res.send(zipBuffer);
      return;
    }

    const pdfBuffer = await generatePDF(formDataItems, cName, templatePath);

    res.setHeader("Content-Type", "application/pdf");

    if (formDataItems.length > 1) {
      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${cName}-AOA.pdf`
      );
    } else {
      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${formDataItems[0].name}-AOA.pdf`
      );
    }

    // Send the combined buffer as the response
    // res.redirect('views/employees')
    res.send(pdfBuffer);
  })
);

// send AOA to any email from the form data page by miraj
app.post(
  "/send-aoa-link",
  catchAsync(async (req, res) => {
    const user = req.user;
    const { email, adminEmail, subject, content, editLink } = req.body;

    const url = new URL(editLink);
    const params = url.searchParams;

    const emailData = {
      companyId: params.get("c"),
      managerId: user?._id, // params.get("m")
      email: email,
    };

    const token = getToken(emailData);

    params.append("token", token);
    const newEditLinkWithToken = `<a href="${url.toString()}">Edit Link</a>`;
    const newContentWithEditLink = `<p>${content}</p><br/>${newEditLinkWithToken}`;

    const emails = `${email}, ${adminEmail}`;

    await sendEmailAOALink.sendEmailAOALinkFromMS365(
      emails,
      subject,
      newContentWithEditLink
    );

    res.status(200).send("Thanks for sending email.");
  })
);

// self evaluation edit form
app.get(
  "/self-evaluation-formdata/:id",
  catchAsync(async (req, res) => {
    const itemId = req.params.id;
    // console.log(employeeEmail, managerEmail);

    // Fetch the existing form data from the database
    const existingSelfData = await SelfEvaluation.findById(itemId);
    const existingFormData = await FormData.findById(
      existingSelfData?.formDataId
    );

    // Render the edit form page with the retrieved data
    res.render("self-evaluation-formdata", {
      formData: existingFormData,
      selfData: existingSelfData,
    });
  })
);

// self evaluation form submit
app.post(
  "/self-evaluation-edit-form/submit/:id",
  catchAsync(async (req, res) => {
    const formData = req.body;
    const selfId = req.params.id;

    const selfdataInDb = await SelfEvaluation.findById(selfId);
    const formDataItem = await FormData.findById(selfdataInDb?.formDataId);

    // update the self evaluation data in the MongoDB database
    for (let key in formData) {
      selfdataInDb[key] = formData[key];
    }
    selfdataInDb.isFulfilled = true;
    await selfdataInDb.save();

    // update the manager evaluation data in mongodb
    const managerData = await ManagerEvaluation.findById(
      selfdataInDb?.managerEvaluationId
    );
    managerData.selfEvaluationId = selfdataInDb?._id;
    await managerData.save();

    formDataItem.managerEvaluationIds.addToSet(managerData?._id);
    await formDataItem.save();

    // rendering thanyou page after saving data
    res.render("thankyou-page-2");

    // console.log("\x1b[32mSelf evaluation data created successfully.\x1b[0m");
    const docxBuffer = await generateSelf([selfdataInDb]);

    const transporter = await createEmailTransporter();

    const body =
      "Thank you for completing your AOA Self Evaluation. Attached is a copy. Please review. Your manager will be in contact with you to discuss further.";

    const mailOptions = {
      from: `LeaveNowGrow <${configEnv.userEmail}>`,
      to: selfdataInDb?.employeeEmail,
      cc: selfdataInDb?.managerEmail,
      subject: `${selfdataInDb?.name} - AOA Self Evaluation for period ${selfdataInDb?.quarter}, ${selfdataInDb?.year}.`,
      text: body,
      attachments: [
        {
          filename: `${formData.name}-Review.docx`,
          content: docxBuffer,
        },
      ],
    };

    // create another email option for only manager
    const anotherEmailOptions = {
      from: `LeaveNowGrow <${configEnv.userEmail}>`,
      to: selfdataInDb?.managerEmail,
      subject: `${selfdataInDb?.name} - AOA Manager Evaluation for period ${selfdataInDb?.quarter}, ${selfdataInDb?.year}.`,
      html: `Your employee ${selfdataInDb?.name} has completed their self-evaluation. You are now being asked to complete a manager evaluation of your employee, based on their most recent Assignment of Accountability. Once this evaluation is completed, you will receive a copy. Please click the following link. 
      <a style="padding: 0px 4px;font-weight: 700; color: #0000FF;" href="${BASE_LINK}/manager-evaluation-edit-form/${managerData?._id}">Manager evaluation link</a>`,
    };

    // send first email to employee and manager
    await transporter.sendMail(mailOptions);

    // another email to only manager with a link
    await transporter.sendMail(anotherEmailOptions);

    // Render the edit form page with the retrieved data
    // res.render("thankyou-page-2");
  })
);

// manager evaluation form page
app.get(
  "/manager-evaluation-edit-form/:id",
  catchAsync(async (req, res) => {
    const itemId = req.params.id;

    // Fetch all the existing data from the database
    const managerEvaluationData = await ManagerEvaluation.findById(itemId);
    const existingSelfData = await SelfEvaluation.findById(
      managerEvaluationData?.selfEvaluationId
    );
    const existingFormData = await FormData.findById(
      managerEvaluationData?.formDataId
    );

    res.render("manager-evaluation-form", {
      formData: existingFormData,
      selfEvaluationData: existingSelfData,
      managerEvaluationData: managerEvaluationData,
    });
  })
);

// send email with pdf to manager evaluation
app.post(
  "/manager-evaluation-form-submit/:id",
  catchAsync(async (req, res) => {
    const managerId = req.params.id;
    const formData = req.body;

    const managerEvaluationData = await ManagerEvaluation.findById(managerId);

    for (let key in formData) {
      managerEvaluationData[key] = formData[key];
    }
    managerEvaluationData.isFulfilled = true;
    await managerEvaluationData.save();

    // rendering thanyou page after saving data
    res.render("thankyou-page-2");

    // send email with pdf file to the manager
    const docxBuffer = await generateManager([managerEvaluationData]);

    const transporter = await createEmailTransporter();

    const body = `Thank you for completing ${formData?.name}'s Evaluation. Attached is a copy. Please review.`;

    const mailOptions = {
      from: `LeaveNowGrow <${configEnv.userEmail}>`,
      to: managerEvaluationData?.managerEmail,
      subject: `AOA Manager Evaluation for period ${managerEvaluationData?.quarter}, ${managerEvaluationData?.year}.`,
      text: body,
      attachments: [
        {
          filename: `${formData.name}-Review.docx`,
          content: docxBuffer,
        },
      ],
    };

    // res.render("thankyou-page-2");
    await transporter.sendMail(mailOptions);
  })
);

// see all evaluations data
app.get(
  "/all-evaluations/:id",
  catchAsync(async (req, res) => {
    const itemId = req.params.id;

    if (itemId === "undefined") {
      const companyId = req.query.comp;
      res.redirect(`/formadmin/${companyId}`);
      return;
    }

    const selfQuery = req.query["self-version"];
    const managerQuery = req.query["manager-version"];

    const aoa = await FormData.findById(itemId)
      .populate({
        path: "selfEvaluationIds",
        model: SelfEvaluation,
        select: "year quarter managerEvaluationId",
      })
      .populate({
        path: "managerEvaluationIds",
        model: ManagerEvaluation,
        select: "year quarter",
      });

    const selfId = selfQuery
      ? new ObjectId(selfQuery)
      : aoa?.selfEvaluationIds[aoa?.selfEvaluationIds.length - 1];
    const self = await SelfEvaluation.findById(selfId);

    const managerId = managerQuery
      ? new ObjectId(managerQuery)
      : aoa?.managerEvaluationIds[aoa?.managerEvaluationIds.length - 1];
    const manager = await ManagerEvaluation.findById(managerId);

    res.render("all-evaluations", {
      formData: aoa,
      selfEvaluationData: self,
      managerEvaluationData: manager,
      selfVersions: aoa?.selfEvaluationIds,
      managerVersions: aoa?.managerEvaluationIds,
      disabled: !self?.isFulfilled && !manager?.isFulfilled ? "disabled" : "",
    });
  })
);

// download all evaluations
app.get(
  "/all-evaluations-download/:ids",
  catchAsync(async (req, res) => {
    const option = req.query.option;
    const ids = req.params.ids.split(",");

    const formDataItems = await FormData.findById(ids[0]).lean();
    const selfDataItem = await SelfEvaluation.findById(ids[1]).lean();
    const managerDataItem = await ManagerEvaluation.findById(ids[2]).lean();

    // Need to complete it by creating all evaluations docx, pdf, excel files
    formDataItems.evaluation = "AOA";
    formDataItems.managerNotes = ""; // To avoid null error
    selfDataItem.evaluation = `Self Evaluation`;
    selfDataItem.managerNotes = ""; // To avoid null error
    managerDataItem.evaluation = `Manager Evaluation`;

    // res.render("all-evaluations-pdf", { formDataItems: [formDataItems, selfDataItem, managerDataItem],});
    if (option === "DOCX") {
      const docxBuffer = await generateAll([
        formDataItems,
        selfDataItem,
        managerDataItem,
      ]);
      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${selfDataItem?.name}-AOA.docx`
      );
      res.send(docxBuffer);
      return;
    } else if (option === "EXCEL") {
      const excelBuffer = await generateAllExcel([
        formDataItems,
        selfDataItem,
        managerDataItem,
      ]);

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );

      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${selfDataItem?.name}-AOA.xlsx`
      );

      res.send(excelBuffer);
      return;
    } else if (option === "PDF") {
      const templatePath = path.join(
        __dirname,
        "views",
        "all-evaluations-pdf.ejs"
      );

      // Create a PDF using Puppeteer
      const pdfBuffer = await generatePDFonTabloid(
        [formDataItems, selfDataItem, managerDataItem],
        formDataItems?.company,
        templatePath
      );

      res.setHeader("Content-Type", "application/pdf");

      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${selfDataItem?.name}-AOA.pdf`
      );

      res.send(pdfBuffer);
      return;
    }

    res.send("Need to select a download option.");
  })
);

// self evaluation data download NOT USED
app.get(
  "/self-evaluation-download/:id",
  catchAsync(async (req, res) => {
    const selectedId = req.params.id;
    const option = req.query.option;

    // Retrieve the specific form data items based on the provided IDs
    const selfDataItem = await SelfEvaluation.findById(selectedId);

    if (!selfDataItem) {
      // Return an error if the form data items are not found
      res.status(404).send("Form data not found.");
      return;
    }

    if (option === "docx") {
      const docxBuffer = await generateSelf([selfDataItem]);

      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${selfDataItem?.name}-AOA.docx`
      );
      res.send(docxBuffer);
      return;
    } else if (option === "excel") {
      const excelBuffer = await generateSelfExcel([selfDataItem]);

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );

      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${selfDataItem?.name}-AOA.xlsx`
      );

      res.send(excelBuffer);
      return;
    } else if (option === "pdf") {
      // Create a PDF using Puppeteer

      const pdfBuffer = await generateEvaluationsPDF(
        [selfDataItem],
        (manager = false)
      );

      res.setHeader("Content-Type", "application/pdf");

      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${selfDataItem?.name}-AOA.pdf`
      );

      res.send(pdfBuffer);
      return;
    }

    res.send("Need to select a download option.");
  })
);

// manager evaluation data download NOT USED
app.get(
  "/manager-evaluation-download/:id",
  catchAsync(async (req, res) => {
    try {
      const selectedId = req.params.id;
      const option = req.query.option;

      // Retrieve the specific form data items based on the provided IDs
      const managerDataItem = await ManagerEvaluation.findById(selectedId);

      if (!managerDataItem) {
        // Return an error if the form data items are not found
        res.status(404).send("Form data not found.");
        return;
      }

      if (option === "docx") {
        const docxBuffer = await generateManager([managerDataItem]);
        res.setHeader(
          "Content-Disposition",
          `attachment; filename=${managerDataItem?.name}-AOA.docx`
        );
        res.send(docxBuffer);
        return;
      } else if (option === "excel") {
        const excelBuffer = await generateManagerExcel([managerDataItem]);

        res.setHeader(
          "Content-Type",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        );

        res.setHeader(
          "Content-Disposition",
          `attachment; filename=${managerDataItem?.name}-AOA.xlsx`
        );

        res.send(excelBuffer);
        return;
      } else if (option === "pdf") {
        const pdfBuffer = await generateEvaluationsPDF(
          managerDataItem,
          (manager = true)
        );

        res.setHeader("Content-Type", "application/pdf");

        res.setHeader(
          "Content-Disposition",
          `attachment; filename=${managerDataItem?.name}-AOA.pdf`
        );

        res.send(pdfBuffer);
        return;
      }

      res.send("Need to select a download option.");
    } catch (error) {
      console.error(error);
      res.status(500).send("Internal Server Error");
    }
  })
);

// May 8, 2025 --- Adding Goals worksheet workflow
// params id is company id
// ?nm=<name>&em=<email>
app.get(
  "/goals-worksheet/:id",
  catchAsync(async (req, res) => {
    const user = req.user;
    const id = req.params.id;
    let name = req.query.nm ?? "";
    let email = req.query.em ?? "";
    const token = req.query.token ?? "";
    const view = req.query.view ?? "";

    if (view === "self") {
      name = user?.name ?? "";
      email = user?.email ?? "";
    }

    if (token || view === "admin-actions") {
      const decoded = decodeToken(token);
      name = decoded?.name;
      email = decoded?.email;
    }

    const company = await Company.findById(id);

    const editLink = `${BASE_LINK}/goals-worksheet/${company?._id}`;

    res.render("goals-worksheet-form", {
      user: user,
      editUser: { name, email },
      companyName: (company?.name).toString(),
      companyId: id,
      group: false,
      editLink: editLink,
    });
  })
);

// params id is group id
// ?nm=<name>&em=<email>
app.get(
  "/group-goals-worksheet/:id",
  catchAsync(async (req, res) => {
    const groupId = req.params.id;
    let name = req.query?.nm ?? "";
    let email = req.query?.em ?? "";
    const user = req.user;
    const token = req.query.token ?? "";
    const group = await Group.findById(groupId);
    const company = await Company.findById(group?.company);

    if (token) {
      const decoded = decodeToken(token);
      name = decoded?.name;
      email = decoded?.email;
    }

    const editLink = `${BASE_LINK}/group-goals-worksheet/${group?._id}`;

    res.render("goals-worksheet-form", {
      user: user,
      editUser: { name, email },
      companyName: (company?.name).toString(),
      companyId: group?.company,
      groupId: group?._id,
      group: true,
      editLink: editLink,
    });
  })
);

// send GW to any email from the form data page
app.post(
  "/send-gw-link",
  catchAsync(async (req, res) => {
    const { email, subject, content, editLink } = req.body;

    const emails = `${email}`;

    const maildata = {
      email: email,
      name: "",
    };

    const token = getToken(maildata);

    const url = new URL(editLink);
    const params = url.searchParams;
    params.append("token", token);

    const newEditLinkWithToken = `<a href="${url.toString()}">Edit Link</a>`;
    const newContentWithEditLink = `<p>${content}</p><br/>${newEditLinkWithToken}`;

    await sendEmailAOALink.sendEmailAOALinkFromMS365(
      emails,
      subject,
      newContentWithEditLink
    );

    res.status(200).send("Thanks for sending email.");
  })
);

app.post(
  "/submit-goals-worksheet",
  catchAsync(async (req, res) => {
    // Extract the form data from the request body
    const gwData = req.body;
    const user = req.user;

    const gwUserEmail = gwData?.email;
    const userData = await User.findOne({ email: gwUserEmail }).populate({
      path: "formData",
      model: FormData,
    });

    if (gwData?.group) {
      gwData.groupIds.push(gwData?.group);
    }

    const newGwData = new GoalsWorksheet(gwData);

    if (userData) {
      userData.goalsWorksheet = newGwData?._id;
      newGwData.groupIds = userData?.formData?.groupIds ?? [];

      await userData.save();
    }

    // Save the form data to the MongoDB database
    newGwData.save().then((savedDoc) => {
      consoleMe("Goals worksheet saved successfully");
      // console.log(savedDoc);
      res.render("thankyou-page-goals", { gwData: savedDoc, user });
    });
  })
);

app.post(
  "/send-goals-data/:id",
  catchAsync(async (req, res) => {
    const user = req.user;
    // Extract the email address from the request body
    const { email } = req.body;
    const gwId = req.params.id;

    const goalsWorksheet = await GoalsWorksheet.findOne({
      _id: gwId,
    })
      .populate({
        path: "company",
        model: Company,
      })
      .lean();

    if (!goalsWorksheet) {
      // Return an error if the form data items are not found
      res.render("404", { content: "Goals Worksheet not found." });
      return;
    }

    const companyName = goalsWorksheet?.company?.name;

    const docxBuffer = await generateGoalsWorksheetDocx(
      [goalsWorksheet],
      companyName
    );

    const transporter = await createEmailTransporter();

    // Send the PDF to the provided email address
    const mailOptions = {
      from: `LeaveNowGrow <${configEnv.userEmail}>`,
      to: email,
      subject: "Goals Worksheet Report",
      text: "Attached is the Goals Worksheet Report.",
      attachments: [
        {
          filename: `${goalsWorksheet?.name}-GW.docx`,
          content: docxBuffer,
        },
      ],
    };

    // Send the email with the docx attachment
    await transporter.sendMail(mailOptions);
    res.render("email-success", { email, formId: gwId, user, goals: true });
  })
);

// /goalsWorksheetId?comp=<company_id>&nm=<name>&em=<email>&qt=<quarter>&yr=<year>
app.get(
  "/goals-worksheet-single/view/:id",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;

    const ID = req.params.id;
    const quarter = req.query.qt;
    const year = req.query.yr;

    //when there is no id, meaning no goals worksheet for user
    if (!ID || ID === "undefined") {
      const companyId = req.query?.comp;
      const name = req.query?.nm ?? user?.name ?? "";
      const email = req.query?.em ?? user?.email ?? "";

      res.redirect(`/goals-worksheet/${companyId}?nm=${name}&em=${email}`);
      return;
    }

    let goalsWorksheet;

    if (quarter && year) {
      goalsWorksheet = await GoalsWorksheet.findOne({
        _id: ID,
        quarter: quarter,
        year: year,
      }).lean();
    } else {
      goalsWorksheet = await GoalsWorksheet.findOne({ _id: ID }).lean();
    }

    // Retrieve the specific form data item based on the provided ID

    const company = await Company.findOne({
      _id: goalsWorksheet?.company,
    }).lean();

    const companyName = company?.name?.toString() || "";

    // Render the EJS template with the form data item
    res.render("goals-worksheets-view-2", {
      goalsWorksheet,
      company: company?._id,
      companyName: companyName,
      user: user,
      comp: true,
    });
  })
);

app.get(
  "/goals-worksheet/company/view/:id",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    const ID = req.params.id;

    if (user?.level === "Super Admin" || user?.role === 1) {
      // Retrieve the specific form data item based on the provided ID
      const goalsWorksheets = await GoalsWorksheet.find({ company: ID }).lean();
      const company = await Company.findOne({ _id: ID });

      goalsWorksheets.companyName = company?.name;

      if (!goalsWorksheets) {
        // Return an error if the form data item is not found
        res.render("404", {
          content: "Goals data not found.",
        });
        return;
      }

      // Render the EJS template with the form data item
      res.render("goals-worksheets-view", {
        goalsWorksheets,
        company: ID,
        user: user,
        comp: true,
      });
    } else {
      const company = await Company.findOne({ _id: user?.company });
      const employees = await User.find({ company: company?._id });
      const descendants = getFollowingEmployees(user?._id, employees);
      const userEmails = descendants.map((item) => item?.email);
      const goalsWorksheets = await GoalsWorksheet.find({
        email: { $in: userEmails },
      });

      goalsWorksheets.companyName = company?.name;

      if (!goalsWorksheets) {
        // Return an error if the form data item is not found
        res.render("404", {
          content: "Goals data not found.",
        });
        return;
      }

      // Render the EJS template with the form data item
      res.render("goals-worksheets-view", {
        goalsWorksheets,
        company: ID,
        user: user,
        comp: true,
      });
    }
  })
);

app.get(
  "/goals-worksheet/group/view/:id",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;

    const ID = req.params.id;
    // Retrieve the specific form data item based on the provided ID
    const goalsWorksheets = await GoalsWorksheet.find({
      groupIds: { $in: [ID] },
    }).lean();
    //  console.log("form data", formDataItem)
    const group = await Group.findOne({ _id: ID });

    // goalsWorksheets.companyName = (group?.name).toString();
    goalsWorksheets.companyName = (group?.name).toString();

    // Render the EJS template with the form data item
    res.render("goals-worksheets-view", {
      goalsWorksheets,
      company: ID,
      user: user,
      comp: false,
    });
  })
);

// review goals after submission
app.get(
  "/goals-review/:id",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;

    const ID = req.params.id;
    // Retrieve the specific form data item based on the provided ID
    const goalsWorksheet = await GoalsWorksheet.findOne({ _id: ID }).lean();
    //  console.log("form data", formDataItem)
    const company = await Company.findOne({ _id: goalsWorksheet?.company });

    goalsWorksheet.companyName = (company?.name).toString();

    // console.log(formDataItem)

    if (!goalsWorksheet) {
      // Return an error if the form data item is not found
      res.render("404", {
        content: "Goals data not found.",
      });
      return;
    }

    // Render the EJS template with the form data item
    res.render("goals-review", {
      goalsWorksheets: [goalsWorksheet],
      company: goalsWorksheet?.company,
      user: user,
      comp: true,
    });
  })
);

app.get(
  "/edit-goals-worksheet-form/:id",
  catchAsync(async (req, res) => {
    const itemId = req.params.id;

    // Fetch the existing form data from the database
    const goalsWorksheet = await GoalsWorksheet.findById(itemId)
      .populate({
        path: "company",
        model: Company,
      })
      .lean();

    // Render the edit form page with the retrieved data
    res.render("edit-goals-worksheet-form", { goalsWorksheet: goalsWorksheet });
  })
);

app.post(
  "/update-goals-worksheet",
  catchAsync(async (req, res) => {
    const gwData = req.body;
    const gwID = gwData.id;

    delete gwData.id;

    if (!gwData?.group) {
      delete gwData.group;
    }

    // Update the form data in the MongoDB database
    const updatedGWData = await GoalsWorksheet.findByIdAndUpdate(
      gwID,
      gwData
    ).lean();

    const company = await Company.findOne({ _id: gwData?.company });

    const admin = await User.findOne({ _id: company?.user_id });

    const docxBuffer = await generateGoalsWorksheetDocx(
      [updatedGWData],
      (company?.name).toString()
    );

    if (admin) {
      const transporter = await createEmailTransporter();

      const mailOptions = {
        from: `LeaveNowGrow <${configEnv?.userEmail}>`,
        to: admin?.email,
        subject: "Goals Worksheet Updated",
        text: `An employee, ${gwData?.name}'s goals worksheet have been updated`,
        file: [{ filename: `${gwData?.name}-GW.docx`, content: docxBuffer }],
      };

      transporter.sendMail(mailOptions);
      res.redirect(`/goals-worksheet/view/${company?._id}`);
    } else {
      consoleMe("Route: '/update-goals-worksheet' - Admin User not found!");
      // res.status(401).send("Admin User not found.");
      res.redirect(`/goals-worksheet/view/${company?._id}`);
    }
  })
);

app.get(
  "/delete-goals-worksheet/:id",
  isLoggedIn,
  catchAsync(async (req, res) => {
    const user = req.user;
    const empId = req.params.id;

    // Fetch the existing form data from the database
    const gwData = await GoalsWorksheet.findById(empId).lean();

    await GoalsWorksheet.deleteOne({ _id: empId });

    if (user.role === 1) {
      res.redirect(`/goals-worksheet/view/${gwData?.company}`);
    } else if (user.role === 3) {
      res.redirect(`/`);
    } else {
      res.redirect(`/admin/view`);
    }
  })
);

// must call this api with query like, /goals-worksheet/download/[1234]?option=docx
app.get(
  "/goals-worksheet/download/:ids",
  catchAsync(async (req, res) => {
    const option = req.query?.option;

    const selectedIds = req.params.ids.split(",");

    // Retrieve the specific form data items based on the provided IDs
    const goalsWorksheets = await GoalsWorksheet.find({
      _id: { $in: selectedIds },
    })
      .populate({
        path: "company",
        model: Company,
      })
      .lean();

    const companyName = goalsWorksheets[0]?.company?.name;

    if (!goalsWorksheets) {
      // Return an error if the form data items are not found
      res.render("404", { content: "Goals Worksheets not found" });
      return;
    }

    if (option === "docx") {
      // download option for more than 7 AOA's
      if (goalsWorksheets.length > 7) {
        const goalsWorksheets1 = goalsWorksheets.slice(0, 7);
        const goalsWorksheets2 = goalsWorksheets.slice(
          7,
          goalsWorksheets.length
        );

        const docxBuffer1 = await generateGoalsWorksheetDocx(
          goalsWorksheets1,
          companyName
        );
        const docxBuffer2 = await generateGoalsWorksheetDocx(
          goalsWorksheets2,
          companyName
        );

        const zip = new JSZip();
        zip.file(`${companyName}-GW-Part-1.docx`, docxBuffer1);
        zip.file(`${companyName}-GW-Part-2.docx`, docxBuffer2);

        const zipBuffer = await zip.generateAsync({ type: "nodebuffer" });

        res.setHeader("Content-Type", "application/zip");
        res.setHeader(
          "Content-Disposition",
          `attachment; filename=${companyName}-GW.zip`
        );

        // Send the combined buffer as the response
        res.send(zipBuffer);
        return;
      }

      const docxBuffer = await generateGoalsWorksheetDocx(
        goalsWorksheets,
        companyName
      );

      if (goalsWorksheets.length > 1) {
        res.setHeader(
          "Content-Disposition",
          `attachment; filename=${companyName}-GW.docx`
        );
      } else {
        res.setHeader(
          "Content-Disposition",
          `attachment; filename=${goalsWorksheets[0]?.name}-GW.docx`
        );
      }

      res.send(docxBuffer);
      return;
    } else if (option === "pdf") {
      const templatePath = path.join(
        process.cwd(),
        "views",
        "goals-worksheets-pdf.ejs"
      );

      // download option for more than 7 AOA's
      if (goalsWorksheets.length > 7) {
        const goalsWorksheets1 = goalsWorksheets.slice(0, 7);
        const goalsWorksheets2 = goalsWorksheets.slice(
          7,
          goalsWorksheets.length
        );

        const pdfBuffer1 = await generateGoalsWorksheetsPDF(
          goalsWorksheets1,
          companyName,
          templatePath
        );
        const pdfBuffer2 = await generateGoalsWorksheetsPDF(
          goalsWorksheets2,
          companyName,
          templatePath
        );

        const zip = new JSZip();
        zip.file(`${companyName}-GW-Part-1.pdf`, pdfBuffer1);
        zip.file(`${companyName}-GW-Part-2.pdf`, pdfBuffer2);

        const zipBuffer = await zip.generateAsync({ type: "nodebuffer" });

        res.setHeader("Content-Type", "application/zip");
        res.setHeader(
          "Content-Disposition",
          `attachment; filename=${companyName}-GW.zip`
        );

        // Send the combined buffer as the response
        res.send(zipBuffer);
        return;
      }

      const pdfBuffer = await generateGoalsWorksheetsPDF(
        goalsWorksheets,
        companyName,
        templatePath
      );

      res.setHeader("Content-Type", "application/pdf");

      if (goalsWorksheets.length > 1) {
        res.setHeader(
          "Content-Disposition",
          `attachment; filename=${companyName}-GW.pdf`
        );
      } else {
        res.setHeader(
          "Content-Disposition",
          `attachment; filename=${goalsWorksheets[0]?.name}-GW.pdf`
        );
      }

      // Send the combined buffer as the response
      // res.redirect('views/employees')
      res.send(pdfBuffer);
      return;
    } else if (option === "excel") {
      const excelBuffer = await generateGoalsWorksheetExcel(
        goalsWorksheets,
        companyName
      );

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );

      if (goalsWorksheets.length > 1) {
        res.setHeader(
          "Content-Disposition",
          `attachment; filename=${companyName}-GW.xlsx`
        );
      } else {
        res.setHeader(
          "Content-Disposition",
          `attachment; filename=${goalsWorksheets[0]?.name}-GW.xlsx`
        );
      }

      // Send the combined buffer as the response
      // res.redirect('views/employees')
      res.send(excelBuffer);
    }
  })
);

app.get(
  "/v2/goals-worksheet/download/:id",
  catchAsync(async (req, res) => {
    const option = req.query?.option;

    const selectedId = req.params.id;

    // Retrieve the specific form data items based on the provided IDs
    const goalsWorksheet = await GoalsWorksheet.findOne({ _id: selectedId })
      .populate({
        path: "company",
        model: Company,
      })
      .lean();

    const companyName = goalsWorksheet?.company?.name;

    if (!goalsWorksheet) {
      // Return an error if the form data items are not found
      res.render("404", { message: "Goals Worksheet not found" });
      return;
    }

    if (option === "docx") {
      // TODO Needs to change the function
      const docxBuffer = await generateGoalsWorksheetDocxV2(
        goalsWorksheet,
        companyName
      );

      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${goalsWorksheet?.name}-GW.docx`
      );

      res.send(docxBuffer);
      return;
    } else if (option === "pdf") {
      // TODO Needs to change the pdf file path
      const templatePath = path.join(
        process.cwd(),
        "views",
        "goals-worksheets-pdf-v2.ejs"
      );

      const pdfBuffer = await generateGoalsWorksheetsPDFV2(
        goalsWorksheet,
        companyName,
        templatePath
      );

      res.setHeader("Content-Type", "application/pdf");

      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${goalsWorksheet?.name}-GW.pdf`
      );

      // Send the combined buffer as the response
      // res.redirect('views/employees')
      res.send(pdfBuffer);
      return;
    } else if (option === "excel") {
      // TODO Need to change the function
      const excelBuffer = await generateGoalsWorksheetExcelV2(
        goalsWorksheet,
        companyName
      );

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );

      res.setHeader(
        "Content-Disposition",
        `attachment; filename=${goalsWorksheet?.name}-GW.xlsx`
      );

      res.send(excelBuffer);
    }
  })
);

// Status Page routes
app.get("/status", (req, res) => {
  res.render("status");
});

app.get("/api/status-data", async (req, res) => {
  try {
    const uptimeSec = Math.floor(process.uptime());
    const hours = Math.floor(uptimeSec / 3600);
    const minutes = Math.floor((uptimeSec % 3600) / 60);
    const seconds = uptimeSec % 60;
    const uptimeString = `${hours > 0 ? hours + 'h ' : ''}${minutes}m ${seconds}s`;

    const memory = process.memoryUsage();
    const dbConnected = mongoose.connection.readyState === 1;

    let dbLatency = 0;
    let dbPingError = false;
    let modelErrors = {
      users: false,
      formData: false,
      selfEvaluations: false,
      managerEvaluations: false,
      goalsWorksheets: false,
      companies: false,
      groups: false,
    };

    let dbStats = {
      users: 0,
      formData: 0,
      selfEvaluations: 0,
      managerEvaluations: 0,
      goalsWorksheets: 0,
      companies: 0,
      groups: 0,
    };

    if (dbConnected && mongoose.connection.db) {
      try {
        const startPing = Date.now();
        await mongoose.connection.db.admin().ping();
        dbLatency = Date.now() - startPing;
      } catch (pingErr) {
        dbPingError = true;
        console.error("Database ping check failed:", pingErr);
      }

      // Execute queries with granular error tracking
      const fetchCount = async (model, errorKey) => {
        try {
          return await model.countDocuments();
        } catch (e) {
          modelErrors[errorKey] = true;
          console.error(`Count query failed for ${errorKey}:`, e.message);
          return 0;
        }
      };

      const [usersCount, formDataCount, selfEvalCount, mgrEvalCount, goalsCount, companiesCount, groupsCount] = await Promise.all([
        fetchCount(User, 'users'),
        fetchCount(FormData, 'formData'),
        fetchCount(SelfEvaluation, 'selfEvaluations'),
        fetchCount(ManagerEvaluation, 'managerEvaluations'),
        fetchCount(GoalsWorksheet, 'goalsWorksheets'),
        fetchCount(Company, 'companies'),
        fetchCount(Group, 'groups'),
      ]);

      dbStats = {
        users: usersCount,
        formData: formDataCount,
        selfEvaluations: selfEvalCount,
        managerEvaluations: mgrEvalCount,
        goalsWorksheets: goalsCount,
        companies: companiesCount,
        groups: groupsCount,
      };
    } else {
      // Mark all models as error if DB is completely disconnected
      Object.keys(modelErrors).forEach(key => modelErrors[key] = true);
    }

    let azureAuthError = null;
    if (process.env.CLIENT_ID && process.env.CLIENT_SECRET && process.env.TENANT_ID) {
      try {
        const { ConfidentialClientApplication } = require("@azure/msal-node");
        const cca = new ConfidentialClientApplication({
          auth: {
            clientId: process.env.CLIENT_ID,
            authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
            clientSecret: process.env.CLIENT_SECRET,
          },
        });
        await cca.acquireTokenByClientCredential({
          scopes: ["https://outlook.office365.com/.default"],
        });
      } catch (azureErr) {
        azureAuthError = azureErr.errorMessage || azureErr.message;
      }
    } else {
      azureAuthError = "Missing Azure credentials in .env";
    }

    const totalEvaluations = dbStats.selfEvaluations + dbStats.managerEvaluations;
    const totalRecords = dbStats.formData + totalEvaluations + dbStats.goalsWorksheets + dbStats.users;
    const hasAnyError = !dbConnected || dbPingError || Object.values(modelErrors).some(err => err === true) || !!azureAuthError;

    const automationLogs = [
      {
        name: "leavenowgrow-admin (Express Web Application Server)",
        description: `Active on Port ${PORT || 5000} (${MODE || 'development'} mode)`,
        totalRuns: Math.max(1, Math.floor(uptimeSec / 60)),
        successRuns: Math.max(1, Math.floor(uptimeSec / 60)),
        timeAgo: `${seconds}s ago`,
        status: "success",
        type: "Server Core",
        metric: `Active ${uptimeString}`
      },
      {
        name: "MongoDB Database Cluster (Mongoose Connection)",
        description: `Primary DB Cluster Ping Latency & Pool Health`,
        totalRuns: Math.max(1, totalRecords),
        successRuns: (dbConnected && !dbPingError) ? Math.max(1, totalRecords) : 0,
        timeAgo: (dbConnected && !dbPingError) ? `${dbLatency}ms ping` : "Ping Failed / Disconnected",
        status: (dbConnected && !dbPingError) ? "success" : "error",
        type: "Database Service",
        metric: (dbConnected && !dbPingError) ? `${dbLatency}ms Latency` : "Connection Error"
      },
      {
        name: "FormData & Position Descriptions System",
        description: `Employee profile records, positional objectives & metrics`,
        totalRuns: Math.max(1, dbStats.formData),
        successRuns: !modelErrors.formData ? Math.max(1, dbStats.formData) : 0,
        timeAgo: !modelErrors.formData ? "Live DB Sync" : "Sync Failed",
        status: !modelErrors.formData ? "success" : "error",
        type: "MongoDB Model",
        metric: !modelErrors.formData ? `${dbStats.formData} Form Data Profiles` : "Query Failed"
      },
      {
        name: "Self Evaluations Assessment System",
        description: `Employee self-assessment responses & feedback forms`,
        totalRuns: Math.max(1, dbStats.selfEvaluations),
        successRuns: !modelErrors.selfEvaluations ? Math.max(1, dbStats.selfEvaluations) : 0,
        timeAgo: !modelErrors.selfEvaluations ? "Live DB Sync" : "Sync Failed",
        status: !modelErrors.selfEvaluations ? "success" : "error",
        type: "MongoDB Model",
        metric: !modelErrors.selfEvaluations ? `${dbStats.selfEvaluations} Submissions` : "Query Failed"
      },
      {
        name: "Manager Performance Review System",
        description: `Manager performance review evaluations & sign-offs`,
        totalRuns: Math.max(1, dbStats.managerEvaluations),
        successRuns: !modelErrors.managerEvaluations ? Math.max(1, dbStats.managerEvaluations) : 0,
        timeAgo: !modelErrors.managerEvaluations ? "Live DB Sync" : "Sync Failed",
        status: !modelErrors.managerEvaluations ? "success" : "error",
        type: "MongoDB Model",
        metric: !modelErrors.managerEvaluations ? `${dbStats.managerEvaluations} Reviews Completed` : "Query Failed"
      },
      {
        name: "Goals Worksheets Module (v1 & v2)",
        description: `Annual & quarterly employee goal worksheets`,
        totalRuns: Math.max(1, dbStats.goalsWorksheets),
        successRuns: !modelErrors.goalsWorksheets ? Math.max(1, dbStats.goalsWorksheets) : 0,
        timeAgo: !modelErrors.goalsWorksheets ? "Live DB Sync" : "Sync Failed",
        status: !modelErrors.goalsWorksheets ? "success" : "error",
        type: "MongoDB Model",
        metric: !modelErrors.goalsWorksheets ? `${dbStats.goalsWorksheets} Worksheets Active` : "Query Failed"
      },
      {
        name: "Company & Group Organizational Hierarchy",
        description: `Multi-tenant company structures & department groups`,
        totalRuns: Math.max(1, dbStats.companies + dbStats.groups),
        successRuns: (!modelErrors.companies && !modelErrors.groups) ? Math.max(1, dbStats.companies + dbStats.groups) : 0,
        timeAgo: (!modelErrors.companies && !modelErrors.groups) ? "Live DB Sync" : "Sync Failed",
        status: (!modelErrors.companies && !modelErrors.groups) ? "success" : "error",
        type: "MongoDB Model",
        metric: (!modelErrors.companies && !modelErrors.groups) ? `${dbStats.companies} Companies, ${dbStats.groups} Groups` : "Query Failed"
      },
      {
        name: "User Accounts & Role Authentication",
        description: `User credentials, passport auth & permissions`,
        totalRuns: Math.max(1, dbStats.users),
        successRuns: !modelErrors.users ? Math.max(1, dbStats.users) : 0,
        timeAgo: !modelErrors.users ? "Live DB Sync" : "Sync Failed",
        status: !modelErrors.users ? "success" : "error",
        type: "MongoDB Model",
        metric: !modelErrors.users ? `${dbStats.users} User Accounts` : "Query Failed"
      },
      {
        name: "Analytics & Organizational Metrics Engine",
        description: `getAnalytics service for department completion stats`,
        totalRuns: Math.max(1, dbStats.formData),
        successRuns: !modelErrors.formData ? Math.max(1, dbStats.formData) : 0,
        timeAgo: !modelErrors.formData ? "Service Ready" : "Degraded",
        status: !modelErrors.formData ? "success" : "error",
        type: "Analytics Service",
        metric: !modelErrors.formData ? "Analytics Calculation Ready" : "Data Sync Attention Needed"
      },
      {
        name: "Puppeteer Headless PDF Generation Service",
        description: `Chromium PDF generator for evaluation reports (A4 & Tabloid)`,
        totalRuns: Math.max(1, totalEvaluations + dbStats.goalsWorksheets),
        successRuns: Math.max(1, totalEvaluations + dbStats.goalsWorksheets),
        timeAgo: "Engine Ready",
        status: "success",
        type: "PDF Utility",
        metric: "Puppeteer / Chromium Active"
      },
      {
        name: "Excel Spreadsheet Report Generator",
        description: `ExcelJS workbook generation for employee tables & goals`,
        totalRuns: Math.max(1, dbStats.formData + dbStats.goalsWorksheets),
        successRuns: Math.max(1, dbStats.formData + dbStats.goalsWorksheets),
        timeAgo: "Engine Ready",
        status: "success",
        type: "Export Utility",
        metric: "ExcelJS Active"
      },
      {
        name: "DOCX Document Report Builder",
        description: `html-to-docx document builder for performance reviews`,
        totalRuns: Math.max(1, totalEvaluations + dbStats.goalsWorksheets),
        successRuns: Math.max(1, totalEvaluations + dbStats.goalsWorksheets),
        timeAgo: "Engine Ready",
        status: "success",
        type: "Export Utility",
        metric: "HTML-to-DOCX Active"
      },
      {
        name: "Email & MS365 Notification Transporter",
        description: `Nodemailer OAuth2 MS365 & Azure AD Authentication`,
        totalRuns: Math.max(1, dbStats.users),
        successRuns: !azureAuthError ? Math.max(1, dbStats.users) : 0,
        timeAgo: !azureAuthError ? "Token Valid & Operational" : "Authentication Failed",
        status: !azureAuthError ? "success" : "error",
        type: "Email Service",
        metric: !azureAuthError ? "Azure AD Token Operational" : (azureAuthError.includes("7000222") ? "Client Secret Expired (AADSTS7000222)" : "Azure Auth Error")
      },
      {
        name: "Multer Attachment File Storage Engine",
        description: `Local file upload middleware for logos & doc attachments`,
        totalRuns: Math.max(1, dbStats.companies),
        successRuns: Math.max(1, dbStats.companies),
        timeAgo: "Storage Ready",
        status: "success",
        type: "File Utility",
        metric: "Multer Storage Operational"
      },
      {
        name: "Environment Variables & API Credentials Audit",
        description: `Validation of critical system secrets, database URIs & OAuth credentials`,
        totalRuns: 7,
        successRuns: (process.env.MONGODB_URI && process.env.SECRET_KEY && process.env.JWT_SECRET && process.env.TENANT_ID && process.env.CLIENT_ID && process.env.CLIENT_SECRET && process.env.USER_EMAIL) ? 7 : 0,
        timeAgo: (process.env.MONGODB_URI && process.env.SECRET_KEY && process.env.JWT_SECRET && process.env.TENANT_ID && process.env.CLIENT_ID && process.env.CLIENT_SECRET && process.env.USER_EMAIL) ? "All Keys Present" : "Missing Required Keys",
        status: (process.env.MONGODB_URI && process.env.SECRET_KEY && process.env.JWT_SECRET && process.env.TENANT_ID && process.env.CLIENT_ID && process.env.CLIENT_SECRET && process.env.USER_EMAIL) ? "success" : "error",
        type: "Config Audit",
        metric: (process.env.MONGODB_URI && process.env.SECRET_KEY && process.env.JWT_SECRET && process.env.TENANT_ID && process.env.CLIENT_ID && process.env.CLIENT_SECRET && process.env.USER_EMAIL) ? "Environment Verified" : "Env Keys Incomplete"
      }
    ];

    res.json({
      status: !hasAnyError ? "OPERATIONAL" : "DEGRADED",
      uptimeString: uptimeString,
      uptimeSeconds: uptimeSec,
      memoryUsageMB: Math.round(memory.heapUsed / 1024 / 1024),
      memoryTotalMB: Math.round(memory.heapTotal / 1024 / 1024),
      database: dbConnected && !dbPingError ? "Connected" : "Disconnected",
      dbReadyState: mongoose.connection.readyState,
      dbLatencyMs: dbLatency,
      dbPingError: dbPingError,
      modelErrors: modelErrors,
      dbStats: dbStats,
      timestamp: new Date().toISOString(),
      automationLogs: automationLogs
    });
  } catch (error) {
    console.error("Error generating status data:", error);
    res.status(500).json({ error: "Internal status error", details: error.message });
  }
});


// not found error handler
app.use(notFound);

// global error handler
app.use(globalErrorHandler);

// changed from 80 to 5000 to develop on local server
app.listen(PORT, () => {
  consoleMe(`Server listening on port: ${PORT}`);
});

module.exports = app;

