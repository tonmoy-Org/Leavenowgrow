const consoleMe = (text, data) => {
  if (data) {
    console.log(`${new Date()}: `, text, data);
    return;
  }
  console.log(
    `${new Date().toLocaleString()}: `,
    text
  );
};

module.exports = { consoleMe };
