const { User } = require("../models/user");

function validateEmailAccessibility(email) {
  return User.findOne({ email: email }).then(function (result) {
    return result !== null;
  });
}

module.exports = { validateEmailAccessibility };
