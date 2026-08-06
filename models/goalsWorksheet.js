/**
 * Mongoose schema for the GoalsWorksheet collection.
 * 
 * This schema represents a worksheet for tracking a user's goals, benefits, metrics, and actions
 * for a specific quarter and year. Each worksheet entry contains details about the goal, its benefit,
 * how it is measured, and the action plan.
 * 
 * Fields:
 * - company: Reference to the associated Comapny document.
 * - group: Reference to the associated Group document.
 * - groupIds: Reference to the associated Group document, will replace the group field.
 * - name: Name of the user completing the worksheet.
 * - email: Email address of the user.
 * - quarter: The quarter for which the worksheet is relevant (e.g., "Q1", "Q2").
 * - year: The year for which the worksheet is relevant.
 * - worksheets: Array of worksheet objects, each containing goal, benefit, metric, and action.
 */
const mongoose = require("mongoose");

const worksheetSchema = new mongoose.Schema(
  {
    _id: false,
    goal: { type: String },
    benefit: { type: String },
    metric: { type: String },
    action: { type: String },
  },
  { _id: false }
);

// self evaluation schema and model defination
const goalsWorksheetSchema = new mongoose.Schema({
  company: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "Company",
  },
  group: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "Group",
  },
  groupIds: [{
    type: mongoose.Schema.Types.ObjectId,
    ref: "Group",
  }],
  name: { type: String },
  email: { type: String },
  quarter: { type: String },
  year: { type: String },
  worksheets: [worksheetSchema],
});

goalsWorksheetSchema.index({ company: 1 });
goalsWorksheetSchema.index({ group: 1 });
goalsWorksheetSchema.index({ email: 1 });

// Create a Mongoose model based on the schema
const GoalsWorksheet = mongoose.model("GoalsWorksheet", goalsWorksheetSchema);

module.exports = { GoalsWorksheet };
