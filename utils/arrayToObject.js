function arrayToObject(arr) {
  if (!Array.isArray(arr)) {
    return {};
  }

  const result = {};
  for (let i = 0; i < arr.length; i++) {
    const obj = arr[i];
    if (typeof obj === "object" && obj !== null && !Array.isArray(obj)) {
      for (const key in obj) {
        if (Object.prototype.hasOwnProperty.call(obj, key)) {
          result[`${key}${i}`] = obj[key];
        }
      }
    }
  }
  return result;
}

module.exports = { arrayToObject };
