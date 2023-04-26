const { Builder, By, until } = require('selenium-webdriver');
const XLSX = require('xlsx');
const fs = require('fs');
const { promisify } = require('util');
const sleep = promisify(setTimeout);

let today = new Date();
let month = today.getMonth() + 1;
let day = today.getDate();
let year = today.getFullYear();

// Adding Zero Before Month or Day When Month and Day Is Less Than 10
if (month < 10) {
  month = `0${month}`;
}

if (day < 10) {
  day = `0${day}`;
}

async function performSteps() {
  let driver = await new Builder().forBrowser('chrome').build();
  try {
    // Step 1: Read URL and input value from an Excel sheet
    const workbook = XLSX.readFile('C:/Users/anike/Desktop/2324/New folder/challan.xlsx');
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const url = sheet['A1'].v;
    const inputValue = sheet['B1'].v;

    // Step 2: Read data from the second Excel sheet and convert to JSON
    const dataWorkbook = XLSX.readFile('C:/Users/anike/Desktop/2324/New folder/Manual Challan Creation.xlsx');
    const dataSheet = dataWorkbook.Sheets[dataWorkbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(dataSheet);

    // Open Website
    await driver.manage().window().maximize();
    await driver.get(url);

    // Array to Hold Skipped GR Numbers
    const skippedValues = [];
    fs.truncateSync('C:/Users/anike/Desktop/2324/New folder/Skipped Values Transit.txt', 0);

    // Loop JSON file until all GR Numbers have been processed
    for (let i = 0; i < jsonData.length; i++) {
      const data = jsonData[i];
      const gr_number = data.gr_number;

      // Enter input value from Excel
      let inputBox = await driver.findElement(By.id('d-1329935-fn_c_grNumber'));
      await inputBox.clear();
      await inputBox.sendKeys(gr_number);

      // Click Show button
      let show_Button = await driver.findElement(By.xpath("//input[@value='Show']"));
      await show_Button.click();

      // Wait for page to load
      await driver.wait(until.titleContains('Challan Creation'), 10000);

      try {
        // Click Edit button
        let Edit_Button = await driver.findElement(By.xpath("//a[contains(@href, 'mode=edit')]"));
        await Edit_Button.click();


let grNumberValue = await driver.findElement(By.id('grNumber')).getAttribute('value');
if (grNumberValue !== gr_number.toString()) {
  await driver.executeScript(`alert('${gr_number} Is Not Matched With This Gr Number ${grNumberValue}')`);
  await driver.sleep(10 * 60 * 1000); // Wait for 10 minute
  continue;
}

        // Enter Ch Number
        let ChBox = await driver.findElement(By.id('outwardChNumber'));
        await ChBox.clear();
        await ChBox.sendKeys(inputValue);

        // Enter outward date from Excel
        let dateBox = await driver.findElement(By.name('outwardDate'));
        await dateBox.clear();
        await dateBox.sendKeys(`${month}/${day}/${year}`);

        // Click Submit button
        let submitButton = await driver.findElement(By.name('submit'));
        await submitButton.click();

        // // Click on Exit Button
        // let exitButton = await driver.findElement(By.name('exit'));
        // await exitButton.click();

        // Wait for page to load
        await driver.wait(until.titleContains('Challan Creation'), 10000);
      } catch (e) {
        // Add skipped value to the array and the file
        skippedValues.push(gr_number);
        fs.appendFileSync('C:/Users/anike/Desktop/2324/New folder/Skipped Values Transit.txt', `${gr_number}\n`);
        continue;
      }
    }

fs.writeFileSync('C:/Users/anike/Desktop/2324/New folder/Skipped Values Transit.txt', skippedValues.join('\n'))

} finally {
// Quit the WebDriver instance
await driver.quit();
}
}

// Call the performSteps function
performSteps();
