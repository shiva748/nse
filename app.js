const puppeteer = require("puppeteer");
const fs = require("fs").promises;
const path = require("path");
const cheerio = require("cheerio");
const ExcelJS = require("exceljs");

const baseUrl = "https://www.nseindia.com/get-quotes/equity?symbol=";
const companies = [
  "TCS",
  "TATACONSUM",
  "TATAMOTORS",
  "ITC",
  "NTPC",
];
const loopInterval = 45000;
const userAgents = [
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/95.0",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Edge/98.0.1108.43 Safari/537.36 Edg/98.0.1108.43",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.0 Safari/605.1.15",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Firefox/96.0",
  "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36",
  "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/95.0",
  "Mozilla/5.0 (Linux; Android 11; Pixel 5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Mobile Safari/537.36",
  "Mozilla/5.0 (iPhone; CPU iPhone OS 15_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.0 Mobile/15E148 Safari/604.1",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:100.0) Gecko/20100101 Firefox/100.0",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/999.0 (KHTML, like Gecko) Chrome/100.0.1000.0 Safari/999.0",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/999.0 (KHTML, like Gecko) Safari/999.0",
  "Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)",
  "Mozilla/5.0 (compatible; Bingbot/2.0; +http://www.bing.com/bingbot.htm)",
  "Mozilla/5.0 (compatible; YandexBot/3.0; +http://yandex.com/bots)",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/91.0",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:90.0) Gecko/20100101 Firefox/90.0",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Edge/92.0.902.67",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36 Edg/91.0.864.59",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36 Edg/92.0.902.67",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/91.0",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Edge/92.0.902.67",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/93.0",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Edge/96.0.1064.0 Safari/537.36 Edg/96.0.1064.0",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:110.0) Gecko/20100101 Firefox/110.0",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/91.0",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Edge/92.0.902.67",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/92.0",
  "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36",
  "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/84.0",
  "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Edge/88.0.705.74 Safari/537.36 Edg/88.0.705.74",
  "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:85.0) Gecko/20100101 Firefox/85.0",
  "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/84.0",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Edge/88.0.705.74 Safari/537.36 Edg/88.0.705.74",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36",
  "Mozilla/5.0 (Linux; Android 11; Pixel 4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.101 Mobile Safari/537.36",
  "Mozilla/5.0 (Linux; Android 11; Pixel 4) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/84.0",
  "Mozilla/5.0 (Linux; Android 11; Pixel 4) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/87.0.4280.101 Safari/537.36",
  "Mozilla/5.0 (Linux; Android 11; Pixel 4) AppleWebKit/537.36 (KHTML, like Gecko) CriOS/87.0.4280.101 Mobile Safari/537.36",
  "Mozilla/5.0 (Linux; Android 11; Pixel 4) AppleWebKit/537.36 (KHTML, like Gecko) Edg/88.0.705.68 Mobile Safari/537.36",
  "Mozilla/5.0 (iPhone; CPU iPhone OS 14_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0.2 Mobile/15E148 Safari/604.1",
  "Mozilla/5.0 (iPhone; CPU iPhone OS 14_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) CriOS/87.0.4280.101 Mobile/15E148 Safari/604.1",
  "Mozilla/5.0 (iPhone; CPU iPhone OS 14_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Edg/88.0.705.68 Mobile/15E148 Safari/604.1",
  "Mozilla/5.0 (iPhone; CPU iPhone OS 14_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) FxiOS/32.0 Mobile/15E148 Safari/605.1.15",
  "Mozilla/5.0 (iPad; CPU OS 14_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0.2 Mobile/15E148 Safari/604.1",
  "Mozilla/5.0 (iPad; CPU OS 14_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) CriOS/87.0.4280.101 Mobile/15E148 Safari/604.1",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:100.0) Gecko/20100101 Firefox/100.0",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/999.0 (KHTML, like Gecko) Chrome/100.0.1000.0 Safari/999.0",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/999.0 (KHTML, like Gecko) Edge/100.0.1000.0 Safari/999.0 Edg/100.0.1000.0",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/999.0 (KHTML, like Gecko) Version/15.0 Safari/999.0",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/999.0 (KHTML, like Gecko) Firefox/100.0",
  "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/999.0 (KHTML, like Gecko) Chrome/100.0.1000.0 Safari/999.0",
  "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/999.0 (KHTML, like Gecko) Firefox/100.0",
  "Mozilla/5.0 (Linux; Android 12; Pixel 6) AppleWebKit/999.0 (KHTML, like Gecko) Chrome/100.0.1000.0 Mobile Safari/999.0",
  "Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) AppleWebKit/999.0 (KHTML, like Gecko) Version/16.0 Mobile/15E148 Safari/999.0",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:110.0) Gecko/20100101 Firefox/110.0",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/999.0 (KHTML, like Gecko) Chrome/110.0.1000.0 Safari/999.0",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/999.0 (KHTML, like Gecko) Safari/110.0",
  "Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)",
  "Mozilla/5.0 (compatible; Bingbot/2.0; +http://www.bing.com/bingbot.htm)",
  "Mozilla/5.0 (compatible; YandexBot/3.0; +http://yandex.com/bots)",
  "MyCustomBot/1.0 (+http://www.example.com/bot)",
  "MySpecialUserAgent/1.0",
  "SuperBot/2.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36",
];
let browser;

async function createFolderIfNotExists(folderName) {
  try {
    await fs.mkdir(folderName);
  } catch (error) {
    if (error.code !== "EEXIST") {
      throw error;
    }
  }
}
async function saveDataToExcel(paths, data) {
  const excelFileName = paths + "/extracted_data.xlsx";

  let workbook;
  try {
    await fs.access(excelFileName);
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFileName);
  } catch (error) {
    workbook = new ExcelJS.Workbook();
  }

  const worksheet =
    workbook.getWorksheet(1) || workbook.addWorksheet("Sheet 1");

  if (!worksheet.getRow(1).getCell(1).value) {
    worksheet.addRow(["Order Book Data", data.timestamp]);
    worksheet.addRow([]);
    worksheet.addRow(["Ask Side", "", "", "", "Bid Side", "", "", ""]);
  }

  data.orderBook.forEach((element, i) => {
    worksheet.addRow([
      `Price L${i + 1}`,
      element.ask,
      `Q${i + 1}`,
      element.sellquantity,
      `Price L${i + 1}`,
      element.bid,
      `Q${i + 1}`,
      element.buyquantity,
    ]);
  });

  worksheet.addRow([]);
  worksheet.addRow(["Trade Book", "", "", "", "", "", "", "", ""]);
  worksheet.addRow([
    "Open",
    "Low",
    "High",
    "Close",
    "Adjusted Price",
    "Traded Volume",
    "Traded Value",
    "Impact Cost",
  ]);

  worksheet.addRow([
    data.open,
    data.low,
    data.high,
    data.close,
    data.adjusted_price,
    data.Tradevolume,
    data.Tradevalue,
    data.Impactcost,
  ]);
  worksheet.addRow([""]);

  try {
    await workbook.xlsx.writeFile(excelFileName);
    console.log(`${excelFileName} saved.`);
  } catch (err) {
    console.error("Error updating Excel file:", err);
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
}

function extractUsefulData(text) {
  console.log("Extracting usefull content from raw data");
  const $ = cheerio.load(text);
  const table = $("#priceInfoTable");
  const rows = table.find("tbody tr");
  let extractedData = {};
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
  const orderBook = [];
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
    timestamp: `${Date.now()}`,
  };
  return extractedData;
}

async function runScriptForCompany(company) {
  const url = `${baseUrl}${company}`;
  const dynamicDataSelector = "#orderBookTradeVal";
  const page = await browser.newPage();

  try {
    const randomUserAgent =
      userAgents[Math.floor(Math.random() * userAgents.length)];
    await page.setUserAgent(randomUserAgent);
    await page.setViewport({ width: 1366, height: 768 });
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 60000 });

    let contentFound = false;
    let iterationCount = 0;

    while (!contentFound && iterationCount < 5) {
      const element = await page.$(dynamicDataSelector);
      if (element) {
        const content = await page.evaluate(
          (el) => el.textContent.trim(),
          element
        );
        if (content !== "-") {
          contentFound = true;
        }
      }

      iterationCount++;

      await page.waitForTimeout(1000);
    }

    if (!contentFound) {
      console.log("Data not found within iterations for", company);
      await page.close();
      return true;
    }

    const companyName = url.split("=")[1];
    await createFolderIfNotExists(companyName);
    const htmlContent = await page.content();
    let data = extractUsefulData(htmlContent);
    saveDataToJsonFile(companyName, { companyName, ...data });
    saveDataToExcel(companyName, { companyName, ...data });

    await page.close();
    return false;
  } catch (error) {
    console.error("An error occurred for", company, ":", error);
    await page.close();
  }
}

let companiesToRetry = [...companies];

async function runScript() {
  const promises = companiesToRetry.map(runScriptForCompany);
  const results = await Promise.all(promises);
  companiesToRetry = companiesToRetry.filter((_, index) => results[index]);
  return companiesToRetry.length > 0;
}

async function mainLoop() {
  while (true) {
    console.log("Running the script...");
    const shouldRestart = await runScript();
    if (!shouldRestart) {
      console.log("Waiting for the next iteration...");
      await new Promise((resolve) => setTimeout(resolve, loopInterval));
      companiesToRetry = [...companies]
    }
  }
}

(async () => {
  browser = await puppeteer.launch();
  await mainLoop();
})();
