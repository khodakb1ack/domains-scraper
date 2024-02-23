const https = require("https");
const fs = require("fs");
const puppeteer = require("puppeteer");
const { resolve } = require("path");
const readline = require('readline-sync');
const cheerio = require("cheerio")
const path = require('path');
var Excel = require('exceljs');
// const axios = require("axios")

const delay = (ms) => new Promise((r) => setTimeout(r, ms));

const result = []

const categories = JSON.parse(`[
  {
    "URL": "https:\/\/ratingruneta.ru\/outstaffing\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/creative-agencies\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/web-support\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/web+seo\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/e-commerce\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/web\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/apps-creative\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/goverment\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/major\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/digital-agencies\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/seo\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/seo+context\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/context\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/performance\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/marketplaces\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/target\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/crm\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/corporate\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/pr\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/smm\/"
  },
  {
    "URL": "https:\/\/ratingruneta.ru\/branding\/"
  }
]`)
  // {
  //   "URL": "https:\/\/ratingruneta.ru\/apps\/"
  // },
  // const categories = JSON.parse(`[
  //   {
  //     "URL": "https:\/\/ratingruneta.ru\/e-commerce\/premium\/"
  //   },
  //   {
  //     "URL": "https:\/\/ratingruneta.ru\/e-commerce\/high\/"
  //   },
  //   {
  //     "URL": "https:\/\/ratingruneta.ru\/e-commerce\/middle\/"
  //   },
  //   {
  //     "URL": "https:\/\/ratingruneta.ru\/e-commerce\/middle\/"
  //   }
  // ]`)

const workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet('SYUDA VATA');
worksheet.columns = [
  { header: 'name', key: 'name', width: 20 },
  { header: 'site', key: 'site', width: 20 },
];

// let url = readline.question("Enter category url without paging:");
// console.log(`url: ${url}`); // https://ratingruneta.ru/apps
// url = url || 'https://ratingruneta.ru/apps';


async function category (url) {
  // url = url.replace(/\/$/, '') + '/1-200/';
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();
  await page.goto(url, { waitUntil: 'load' });
  await page.setViewport({
    width: 1600,
    height: 900,
  });

  const content = await page.content();
  const $ = cheerio.load(content);
  console.log($('h1').text());

  let script = $('script:contains(__INITIAL_STATE__ )').html();
  script = script ? '[{' + script.replace(/\s/g, '').match(/\"rows\":\[\{(.*)]/)?.[1]?.replace(/,\"next\".*/, '') : '';
  // console.log(script)
  let data
  try {
    data = JSON.parse(script)
  } catch (error) {
    console.log(error)
  }
  data?.forEach(item => {
    item.name && result.push({
      name: item.name,
      site: item.link
    })
  });
  // console.log(JSON.stringify(result))

  result.forEach(item => {
    worksheet.addRow(item)
  });
  await browser.close();

}

(async () => {
  // await category('https://ratingruneta.ru/apps')

  for(let i = 0; i < categories.length; i++) {
    await category(categories[i].URL)
    await delay (5000)
  }
  
  console.log('cteate xlsx file');
  let abc = "abcdefghijklmnopqrstuvwxyz";
  let id = "";
  while (id.length < 6) {
    id += abc[Math.floor(Math.random() * abc.length)];
  }
  return workbook.xlsx.writeFile(`./results/result-ratingruneta-${id}.xlsx`);
})()

