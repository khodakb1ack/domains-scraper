
const puppeteer = require("puppeteer");
const cheerio = require("cheerio")
var Excel = require('exceljs');
const xlsx = require('xlsx');
const fs = require('fs');
// const axios = require("axios")
// const { resolve } = require("path");
// const readline = require('readline-sync');
// const path = require('path');
// const https = require("https");

const delay = (ms) => new Promise((r) => setTimeout(r, ms));
// const domains = [
//   // 'https://www.agima.ru/',
//   // 'https://grphn.ru/',
//   'https://procontext.ru/',
//   // 'https://www.artics.ru/',
//   // 'https://itb-company.com/' // cloudflare
//   // 'https://smmheadshot.ru',
//   // 'https://serenity.agency/',
// ]

async function scrapeDomain(url) {
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();
  try {
    console.log(`===================${url}===================`)
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 15000 }); //domcontentloaded load
    const content = await page.content();
    const $ = cheerio.load(content);
    let data = [];
    let contacts = []
    let contactsUrl
    $('a').each(function (index) {
      if (/контакты|contacts/i.test($(this).text())) contacts.push($(this).attr('href'));
    })
    // await delay(1000)
    if (contacts[0] && /^\//.test(contacts[0])) contactsUrl = contacts[0] ? (/http/.test(contacts[0])) ? (contacts[0]) : (url.replace(/\/\?.*/, '') + contacts[0]) : null;
    else contactsUrl = contacts[0] ? (/http/.test(contacts[0])) ? (contacts[0]) : (url.replace(/\?.*/, '') + contacts[0]) : null;
    $('a').each(function (index) {
      let href = $(this).attr('href');
      if (/@/.test(href)) {
        href = href?.replace(/mailto:/, '')
        if (href?.length < 50 && !data.find(item => item == href)) {
          data.push(href)
        }
      }
    });
    // console.log(`data.length = ${data.length}`)
    console.log(`contactsUrl = ${contactsUrl}`)
    if (data.length == 0 && contactsUrl) {
      await delay(1000)
      await page.goto(contactsUrl, { waitUntil: 'domcontentloaded', timeout: 15000 });
      // await delay(10000)
      const content = await page.content();
      const $ = cheerio.load(content);

      $('a').each(function (index) {
        let href = $(this).attr('href')
        // console.log(href)
        if (/@/.test(href)) {
          href = href?.replace(/mailto:/, '')
          if (href?.length < 50 && !data.find(item => item == href)) {
            data.push(href)
          }
        }
      });
    }
    // console.log(script)
    console.log(JSON.stringify(data));

    await browser.close();
    return data;
  } catch (error) {
    await browser.close();
    console.log(error)
  }
  await page.setViewport({
    width: 1600,
    height: 900,
  });
}

function convertExcelFileToJsonUsingXlsx() {
  const file = xlsx.readFile('./input/input.xlsx');
  const sheetNames = file.SheetNames;
  const totalSheets = sheetNames.length;
  let parsedData = [];

  for (let i = 0; i < totalSheets; i++) {
    const tempData = xlsx.utils.sheet_to_json(file.Sheets[sheetNames[i]]);
    tempData.shift(); // Skip header row (column names)
    parsedData.push(...tempData);
  }
  return parsedData
  // generateJSONFile(parsedData);
}

function generateJSONFile(data) {
  try {
    fs.writeFileSync('./input/data.json', JSON.stringify(data));
  } catch (err) {
    console.error(err);
  }
}

function generateExelFile(data) {
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet('SYUDA VATA');
  worksheet.columns = [
    { header: 'name', key: 'name', width: 20 },
    { header: 'site', key: 'site', width: 20 },
    { header: 'email', key: 'email', width: 50 },
  ];
  data.forEach(item => {
    worksheet.addRow(item)
  });
  console.log('===================cteate xlsx file===================');
  let abc = "abcdefghijklmnopqrstuvwxyz";
  let id = "";
  while (id.length < 6) {
    id += abc[Math.floor(Math.random() * abc.length)];
  }
  return workbook.xlsx.writeFile(`./results/result-${id}.xlsx`);
}

const result = []
let domains = convertExcelFileToJsonUsingXlsx();
// domains.length = 10
console.log(`Domains: ${domains.length}`);

(async () => {

  for (let i = 0; i < domains.length; i++) {
    let tryIndex = 0;
    let emails;
    while (!emails && tryIndex < 3) {
      try {
        emails = await scrapeDomain(domains[i].site);
        result.push({
          ...{ email: (emails || []).join(' ') },
          ...domains[i]
        })
        await delay(3000);
      } catch (e) {
        console.error('Load Error  : ' + e);
      } finally {
        tryIndex++;
        await delay(3000);
      }
    }
  }

  return generateExelFile(result)

})()

