const mongoose = require("mongoose");

// manager evaluation schema and model defination
const managerEvaluationSchema = new mongoose.Schema({
  formDataId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "FormData",
  },
  selfEvaluationId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "SelfEvaluation",
  },
  employeeEmail: { type: String },
  managerEmail: { type: String },
  year: { type: String },
  quarter: { type: String },
  name: { type: String },
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
  managerNotes: { type: String },
  managerReview: { type: String },
  isFulfilled: { type: Boolean, default: false },
});

managerEvaluationSchema.index({ formDataId: 1 });
managerEvaluationSchema.index({ selfEvaluationId: 1 });
managerEvaluationSchema.index({ employeeEmail: 1 });
managerEvaluationSchema.index({ managerEmail: 1 });

// Create a Mongoose model based on the schema
const ManagerEvaluation = mongoose.model(
  "ManagerEvaluation",
  managerEvaluationSchema
);

module.exports = { ManagerEvaluation };
