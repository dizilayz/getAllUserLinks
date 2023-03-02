const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const fs = require('fs');
const { resolve } = require('path');


async function getUserLinks() {
    const browser = await puppeteer.launch({
        headless: false,
        userDataDir: 'C:\\Users\\yarovayaev\\AppData\\Local\\Google\\Chrome\\User',
    })
    const page = await browser.newPage();
    // const login = await page.$('.spui-Input__input');
    // await login.type('****');
    // const password = await page.$('input[type=password].spui-Input__input');
    // await password.type('******');
    // const submit = await page.$('.spui-Button');
    // await submit.click();
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const pages = [];
    const userLinks = [];
    for (let i = 2; i <= 16200; i++) {
        if (sheet[`A${i}`]) {
            pages.push(sheet[`A${i}`].v);
        }
    }
    await new Promise(resolve => setTimeout(resolve, 5000));
    for (const url of pages) {
        try {
            await page.goto(url)
            await new Promise(resolve => setTimeout(resolve, 5000));
            const linkInnerText = await page.$eval('span.text-danger', el => el.textContent);
            userLinks.push({ url: linkInnerText });
            console.log(linkInnerText);
        } catch (e) {
            console.log(e.name);
        }
    }
    if (fs.existsSync('userLinks.xlsx')) {
        const workbook = XLSX.readFile('userLinks.xlsx')
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        XLSX.utils.sheet_add_json(sheet, userLinks, { origin: -1, skipHeader: true })
        XLSX.writeFile(workbook, 'userLinks.xlsx')
    } else {
        const newWorkbook = XLSX.utils.book_new()
        const newSheet = XLSX.utils.json_to_sheet(userLinks)
        XLSX.utils.book_append_sheet(newWorkbook, newSheet)
        XLSX.writeFile(newWorkbook, 'userLinks.xlsx')
    }
    try {
        browser.close();
    } catch (e) {
        console.log(e)
    }
}
const filePath = "C:\\Users\\yarovayaev\\Desktop\\allLandings.xlsx"
getUserLinks();