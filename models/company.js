const mongoose = require("mongoose");

const companySchema = new mongoose.Schema({
  name: { type: String },
  user_id: { type: mongoose.Schema.Types.ObjectId, ref: "User" },
  subGroup: [
    {
      type: mongoose.Schema.Types.ObjectId,
      ref: "Group",
    },
  ],
  columnOrder: [{ type: String }],
});

companySchema.index({ user_id: 1 });

const Company = mongoose.model("Company", companySchema);

module.exports = { Company };
