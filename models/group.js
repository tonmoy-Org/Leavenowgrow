const mongoose = require("mongoose");

const groupSchema = new mongoose.Schema({
  name: { type: String },
  adminId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "User",
  },
  columnOrder: [{ type: String }],
  company: { type: mongoose.Schema.Types.ObjectId, ref: "Company" },
  creatorId: { type: mongoose.Schema.Types.ObjectId, ref: "User" },
});

groupSchema.index({ company: 1 });
groupSchema.index({ adminId: 1 });
groupSchema.index({ creatorId: 1 });

const Group = mongoose.model("Group", groupSchema);

module.exports = { Group };
