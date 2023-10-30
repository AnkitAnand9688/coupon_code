
const puppeteer = require("puppeteer");
const XLSX = require("xlsx");

const checkCouponsCode = async () => {
  const workbook = XLSX.readFile("couponcode.xlsx");
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(sheet);
  console.log(jsonData);

  const excelData = jsonData.map((obj) => Object.values(obj));

  console.log('excelData');
  console.log(excelData);

  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();

  await page.goto("https://www.woohoo.in/balenq");
  await page.waitForTimeout(1000)
  console.log('timeout');
   
  const cancel = await page.$("#wzrk-cancel")
  cancel.click()
  const carDetails = []
  console.log('excelData')
  console.log(excelData)
  for (const [code, pin] of excelData) {
    console.log(`code:${code}, pin:${pin}`);
    try {
      const coupounInput = await page.$("#cardNumber");
      const pinInput = await page.$("#cardPin");
      const checkBalanceButton = await page.waitForSelector(
        "#app > div > div.mainContainer > div > div > div.col-12.col-sm-12.col-md-10 > div > div.col-12.col-sm-12.col-md-8.col-lg-8.col-xl-8 > div > div.col-12.col-sm-10.col-md-10.col-lg-10.col-xl-6 > form > div.text-center.text-sm-right.position-relative.form-group > button.mt-2.mt-sm-0.btn.btn-primary"
      );

      await coupounInput.type(code.toString());
      await pinInput.type(pin.toString());


    
          await checkBalanceButton.click();
        
      
    
        await page.waitForSelector(
          "#app > div > div.mainContainer > div > div > div.col-12.col-sm-12.col-md-10 > div > div.col-12.col-sm-12.col-md-8.col-lg-8.col-xl-8 > div > div.mt-3.col-12.col-sm-10.col-md-10.col-lg-10.col-xl-7 > dl",
          { timeout: 5000 }
        );


        const resultElement = await page.$(
          "#app > div > div.mainContainer > div > div > div.col-12.col-sm-12.col-md-10 > div > div.col-12.col-sm-12.col-md-8.col-lg-8.col-xl-8 > div > div.mt-3.col-12.col-sm-10.col-md-10.col-lg-10.col-xl-7 > dl"
        );

        const resultText = await page.evaluate(
          (element) => element.textContent,
          resultElement
        );

        // console.log(resultText);
        const newVal = resultText.split(":")
        // console.log(newVal)
   const regex = /([A-Za-z])\w+/g
   const myval  = newVal[1].replace(regex,'')
   const myvals  = newVal[2].replace(regex,'')
  
     carDetails.push([myval,myvals,newVal[3]]);
     console.log(carDetails)
    } catch (error) {
      console.error(`An error occurred for Coupon Code ${code}: ${error}`);
    }

    await page.reload();
  
  }
// data store in excel
const workbooks = XLSX.readFile('details.xlsx');
const worksheet = workbooks.Sheets[workbooks.SheetNames[0]];


const columnNames = ['Card No', 'Balance', 'Expiry Date'];


const columnNameRow = 1; // Assuming the column names should be in the first row

// Add the column names to the worksheet
columnNames.forEach((columnName, columnIndex) => {
  const cellAddress = XLSX.utils.encode_cell({ r: columnNameRow, c: columnIndex });
  worksheet[cellAddress] = { t: 's', v: columnName };
});

// Calculate the next available row to add the data
const nextRow = XLSX.utils.sheet_add_json(worksheet, carDetails, { skipHeader: true, origin: -1 });

XLSX.writeFile(workbooks, 'details.xlsx');



  await browser.close();
};

checkCouponsCode();

