const { Company } = require("../models/company");
const { FormData } = require("../models/formdata");
const { Group } = require("../models/group");
const { User } = require("../models/user");

const getAnalytics = async () => {
  const [
    totalGroups,
    totalCompanies,
    totalFormData,
    totalAdmin,
    formDataCountsByCompany,
    companyEmployeeCounts,
  ] = await Promise.all([
    Group.countDocuments(),
    Company.countDocuments(),
    FormData.countDocuments(),
    User.countDocuments(),
    FormData.aggregate([
      {
        $group: {
          _id: "$company",
          formDataCount: { $sum: 1 },
        },
      },
    ]),
    Company.find({}, "name").lean(),
  ]);

  const countMap = new Map();
  for (const formdata of formDataCountsByCompany) {
    if (formdata._id) {
      countMap.set(formdata._id.toString(), formdata.formDataCount);
    }
  }

  const companies = companyEmployeeCounts.map((company) => {
    const count = countMap.get(company._id.toString()) || 0;
    return {
      name: company.name,
      formDataCount: count,
    };
  });

  const formDataCounts = companies.map((item) => item.formDataCount);
  const companyNames = companies.map((item) => item.name);

  const analytics = {
    totalGroups,
    totalCompanies,
    totalFormData,
    totalAdmin,
    companyNames,
    formDataCounts,
    companyEmployeeCounts,
  };

  return analytics;
};

module.exports = { getAnalytics };
