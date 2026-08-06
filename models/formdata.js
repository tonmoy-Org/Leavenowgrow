const mongoose = require("mongoose");

// Create a Mongoose schema for the form data
const formDataSchema = new mongoose.Schema({
  company: { type: String, ref: "Company" },
  group: { type: String, ref: "Group" },
  groupIds: [{ type: String, ref: "Group" }],
  user: { type: mongoose.Schema.Types.ObjectId, ref: "User" },
  name: { type: String },
  email: { type: String },
  title: { type: String },
  jobDescription: { type: String },
  managerSupervisor: { type: String },
  directReports: { type: String },
  accountable: { type: String },
  participate: { type: String },
  compensation: { type: String },
  metrics: { type: String },
  positionalObjectives: { type: String },
  personalObjectives: { type: String },
  authorityLevels: { type: String },
  delegationOfAuthority: { type: String },
  selfEvaluationIds: [
    {
      type: mongoose.Schema.Types.ObjectId,
      ref: "SelfEvaluation",
    },
  ],
  managerEvaluationIds: [
    {
      type: mongoose.Schema.Types.ObjectId,
      ref: "ManagerEvaluation",
    },
  ],
});

formDataSchema.index({ company: 1 });
formDataSchema.index({ group: 1 });
formDataSchema.index({ user: 1 });
formDataSchema.index({ email: 1 });

// Create a Mongoose model based on the schema
const FormData = mongoose.model("FormData", formDataSchema);

module.exports = { FormData };
