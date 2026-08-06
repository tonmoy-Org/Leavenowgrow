const { default: puppeteer } = require("puppeteer");

const launchPuppeteer = async () => await puppeteer.launch({ headless: "new" });

module.exports = launchPuppeteer;