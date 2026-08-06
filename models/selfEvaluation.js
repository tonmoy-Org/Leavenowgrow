const mongoose = require("mongoose");

// self evaluation schema and model defination

const selfEvaluationSchema = new mongoose.Schema({
  formDataId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "FormData",
  },
  managerEvaluationId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "ManagerEvaluation",
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
  managerReview: { type: String },
  isFulfilled: { type: Boolean, default: false },
});

selfEvaluationSchema.index({ formDataId: 1 });
selfEvaluationSchema.index({ employeeEmail: 1 });
selfEvaluationSchema.index({ managerEmail: 1 });

// Create a Mongoose model based on the schema
const SelfEvaluation = mongoose.model("SelfEvaluation", selfEvaluationSchema);

module.exports = { SelfEvaluation };

/**
 * Mongoose schema for the SelfEvaluation model.
 *
 * This schema represents a self-evaluation form completed by an employee as part of a performance review process.
 * It includes references to related form data and manager evaluation, as well as various fields capturing
 * employee and manager details, evaluation period, job information, objectives, and review notes.
 *
 * Fields:
 * - formDataId: Reference to the associated FormData document.
 * - managerEvaluationId: Reference to the associated ManagerEvaluation document.
 * - employeeEmail: Email address of the employee completing the self-evaluation.
 * - managerEmail: Email address of the employee's manager.
 * - year: The year of the evaluation period.
 * - quarter: The quarter of the evaluation period.
 * - name: Name of the employee.
 * - title: Job title of the employee.
 * - jobDescription: Description of the employee's job role.
 * - managerSupervisor: Name or identifier of the manager or supervisor.
 * - directReports: Information about direct reports, if any.
 * - accountable: Details regarding the employee's accountabilities.
 * - participate: Information about participation in relevant activities.
 * - compensation: Compensation details.
 * - metrics: Performance metrics relevant to the evaluation.
 * - positionalObjectives: Objectives related to the employee's position.
 * - personalObjectives: Personal objectives set by the employee.
 * - authorityLevels: Levels of authority assigned to the employee.
 * - delegationOfAuthority: Details about delegation of authority.
 * - managerReview: Review notes or comments from the manager.
 * - isFulfilled: Boolean indicating if the self-evaluation process is fulfilled (default: false).
 */