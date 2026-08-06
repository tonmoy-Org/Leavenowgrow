const mongoose = require("mongoose");

const userSchema = new mongoose.Schema(
  {
    name: { type: String },
    email: {
      type: String,
      unique: true,
      required: true,
      trim: true,
      match: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
    },
    password: { type: String, required: true, minlength: 8 },
    profilePictureUrl: {
      type: String,
    },
    company: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "Company",
    },
    manager: {
      id: { type: String },
      name: { type: String },
      email: { type: String },
    },
    role: { type: Number },
    level: { type: String },
    status: { type: Number },
    formData: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "FormData",
    },
    goalsWorksheet: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "GoalsWorksheet",
    },
    admins: [
      {
        type: mongoose.Schema.Types.ObjectId,
        ref: "User",
      },
    ],
    creatorId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "User",
    },
    isVerified: {
      type: Boolean,
      default: false,
    },
  },
  {
    timestamps: true,
  }
);

// High-performance schema indexes for rapid query resolution
userSchema.index({ company: 1 });
userSchema.index({ role: 1 });
userSchema.index({ level: 1 });
userSchema.index({ formData: 1 });
userSchema.index({ goalsWorksheet: 1 });
userSchema.index({ creatorId: 1 });
userSchema.index({ isVerified: 1 });
userSchema.index({ "manager.id": 1 });

const User = mongoose.model("User", userSchema);

module.exports = { User };
