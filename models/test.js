const mongoose = require("mongoose");

const testSchema = new mongoose.Schema({
  company: { type: String },
  firstName: { type: String },
  lastName: { type: String },
  fullName: { type: String },
  emailAddress: { type: String },
  password: { type: String },
  currentAdminYN: { type: String },
  permissionLevel: { type: String },
  manager: { name: { type: String }, email: { type: String } },
  group: { type: String },
  group1: { type: String },
  group2: { type: String },
  unnamed9: { type: String },
});

const Test = mongoose.model("Test", testSchema);

const test2Schema = new mongoose.Schema({
  company: { type: String },
  fullName: { type: String },
  emailAddress: { type: String },
  password: { type: String },
  currentAdminYN: { type: String },
  permissionLevel: { type: String },
  manager: { name: { type: String }, email: { type: String } },
  group1: { type: String },
  group2: { type: String },
});

const Test2 = mongoose.model("Test2", test2Schema);

module.exports = { Test, Test2 };
