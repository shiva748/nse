const puppeteer = require("puppeteer");
const fs = require("fs").promises;
const path = require("path");
const cheerio = require("cheerio");
const ExcelJS = require("exceljs");

const url = "https://www.nseindia.com/get-quotes/equity?symbol=TATAMOTORS";

async function createFolderIfNotExists(folderName) {
  try {
    await fs.mkdir(folderName);
  } catch (error) {
    if (error.code !== "EEXIST") {
      throw error;
    }
  }
}

async function saveDataToJsonFile(paths, data) {
  const filePath = path.join(paths, "extracted_data.json");

  try {
    console.log("Trying to save it in JSON format");

    let existingData = [];
    try {
      const existingContent = await fs.readFile(filePath, "utf-8");
      existingData = JSON.parse(existingContent);

      if (!Array.isArray(existingData)) {
        console.error("Existing data is not an array. Creating a new array.");
        existingData = [];
      }
    } catch (readError) {
      // File doesn't exist or is not in JSON format, which is fine
    }

    existingData.push(data);

    const jsonData = JSON.stringify(existingData, null, 2);
    await fs.writeFile(filePath, jsonData);

    console.log("Data saved");
  } catch (error) {
    console.error("Error saving data:", error);
  }
  (async () => {
    let shouldRestart = true;
    while (shouldRestart) {
      shouldRestart = await runScript();
    }
  })();
}

async function saveDataToExcel(paths, data) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Data");

  // Add headers
  const headers = Object.keys(data);
  worksheet.addRow(headers);

  // Add data rows
  const values = Object.values(data);
  worksheet.addRow(values);

  const excelFilePath = path.join(paths, "extracted_data.xlsx");

  try {
    await workbook.xlsx.writeFile(excelFilePath);
    console.log("Data saved to Excel");
  } catch (error) {
    console.error("Error saving data to Excel:", error);
  }
}

// async function saveDataToJsonFile(paths, data) {
//   try {
//     console.log("trying to save it in json format");
//     const jsonData = JSON.stringify(data, null, 2);
//     await fs.writeFile(path.join(paths, "extracted_data.json"), jsonData);
//     console.log("Data saved");
//   } catch (error) {
//     console.error("Error saving data:", error);
//   }
// }

function extractUsefulData(text) {
  console.log("Extracting usefull content from raw data");
  const $ = cheerio.load(text);
  const table = $("#priceInfoTable");
  const rows = table.find("tbody tr");

  // Initialize an object to store the extracted data
  let extractedData = {};

  // Iterate through each row and extract the data
  function formatFieldName(fieldName) {
    return fieldName
      .replace(/\n/g, "")
      .trim()
      .replace(/[^a-zA-Z0-9]/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_/, "")
      .replace(/_$/, "")
      .toLowerCase();
  }
  rows.each((index, row) => {
    $(row)
      .find("td")
      .each((cellIndex, cell) => {
        const cellHeader = formatFieldName(
          table.find("thead th").eq(cellIndex).text()
        );
        const cellValue = $(cell).text().trim();
        extractedData[cellHeader] = cellValue;
      });
  });
  extractedData = {
    ...extractedData,
    Tradevalue: $("#orderBookTradeVal").text(),
    Tradevolume: $("#orderBookTradeVol").text(),
    Impactcost: $("#orderBookTradeIC").text(),
  };
  const table2 = $("#marketDepthTable");
  const rows2 = table2.find("tbody tr");

  // Initialize an array to store the order book data
  const orderBook = [];

  // Iterate through each row and extract the data
  rows2.each((index, row) => {
    const cells = $(row).find("td");
    if (cells.length === 4) {
      const buyQty = $(cells[0]).text().trim();
      const sellQty = $(cells[3]).text().trim();

      const rowData = {
        buyquantity: buyQty,
        bid: $(cells[1]).text().trim(),
        sellquantity: sellQty,
        ask: $(cells[2]).text().trim(),
      };

      orderBook.push(rowData);
    }
  });
  extractedData = {
    ...extractedData,
    orderBook,
    total_orders: {
      buy: $("#orderBuyTq").text(),
      sell: $("#orderSellTq").text(),
    },
    timestamp: Date.now(),
  };
  return extractedData;
}
async function runScript() {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();

  try {
    await page.setUserAgent(
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36"
    );

    await page.setViewport({ width: 1366, height: 768 });

    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 60000 });

    const dynamicDataSelector = "#orderBookTradeVal";

    let contentFound = false;
    let iterationCount = 0;

    while (!contentFound && iterationCount < 3) {
      await page.screenshot({ path: "current_state.png", fullPage: true });
      const element = await page.$(dynamicDataSelector);
      if (element) {
        const content = await page.evaluate(
          (el) => el.textContent.trim(),
          element
        );
        if (content == "-") {
          console.log("Looking for values ");
        } else {
          console.log("Iteration " + iterationCount + ": " + content);
        }
        if (content !== "-") {
          contentFound = true;
        }
      }

      iterationCount++;

      await page.waitForTimeout(1000);
    }

    if (!contentFound) {
      console.log("Data not found within iterations. Reloading page");
      await browser.close();
      return true;
    }

    const companyName = url.split("=")[1];
    await createFolderIfNotExists(companyName);
    console.log("Final screenshot captured after dynamic data is loaded");
    console.log("extracting data from the file");
    const htmlContent = await page.content();
    // const htmlFilePath = path.join(companyName, "page.html");
    // await fs.writeFile(htmlFilePath, htmlContent);
    let data = extractUsefulData(htmlContent);
    saveDataToJsonFile(companyName, { companyName, ...data });
    saveDataToExcel(companyName, { companyName, ...data });
    const screenshotFilePath = path.join(companyName, "screenshot_final.png");
    await page.screenshot({ path: screenshotFilePath, fullPage: true });
    return false;
  } catch (error) {
    console.error("An error occurred:", error);
  } finally {
    await browser.close();
  }
}

(async () => {
  let shouldRestart = true;
  while (shouldRestart) {
    shouldRestart = await runScript();
  }
})();
