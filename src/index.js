const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('node:path');
const XLSX = require('xlsx');
const writeXlsxFile = require('write-excel-file/node');
const { isContext } = require('node:vm');
const fs = require('fs');
if (require('electron-squirrel-startup')) {
  app.quit();
}


const createWindow = () => {
  const mainWindow = new BrowserWindow({
    title: 'Nimar Motors Khargone',
    // width: 1290,
    // height: 1080,
    icon: path.join(__dirname, './assets/NimarMotor.png'),
    // autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: true,
      preload: path.join(__dirname, "preload.js"),
    },
  });
  ipcMain.on('reset-application', () => {
    mainWindow.reload();
  });
  mainWindow.once('ready-to-show', () => {
    mainWindow.maximize()
  })

  ipcMain.on('reset-app', () => {
    if (mainWindow) {
      KeyMissing = false;
      mainWindow.reload();
    }
  });

  mainWindow.loadFile(path.join(__dirname, "index.html"));
  // mainWindow.webContents.openDevTools();
};

//Separate Calculation Functions
const MGAfunc = require('./functions/MGACalculation');
const CDIfunc = require('./functions/CDICalculation');
const EWfunc = require('./functions/EWCalculation');
const CCPfunc = require('./functions/CCPCalculation');
const MSSFfunc = require('./functions/MSSFCalculation');
const DiscountFunc = require('./functions/DiscountCalculation')
const ExchangeFunc = require('./functions/ExchangeStatusCalculation')
const ComplaintFunc = require('./functions/ComplaintCalculation');
const PerModelCarFunc = require('./functions/PerModelCalculation');
const SpecialCarFunc = require('./functions/SpecialCarCalculation');
const PerCarFunc = require('./functions/PerCarCalculation');
const MSRFunc = require('./functions/MSRCalculation');
const SuperCarFunc = require('./functions/SuperCarCalculation');
const NewDSEincentiveCalculation = require('./functions/NewDSEincentiveCalculation');

// Global Variables
let MGAdata = [];
let CDIdata = [];
let salesExcelDataSheet = [];
let employeeStatusDataSheet = [];
let qualifiedRM = [];
let nonQualifiedRM = [];
let newRm = [];
let newDSEIncentiveDataSheet = [];
let KeyMissing = false;

//////////////////////////////////////////////////////
function checkKeys(array, keys) {
  const firstObject = array[0];
  const missingKeys = [];
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (!firstObject.hasOwnProperty(key)) {
      missingKeys.push(key);
    }
  }
  return missingKeys.length > 0 ? missingKeys : null;
}

function transformKeys(array) {
  return array.map(obj => {
    let newObj = {};
    for (let key in obj) {
      if (obj.hasOwnProperty(key)) {
        let newKey = key.trim();
        newObj[newKey] = obj[key];
      }
    }
    return newObj;
  });
}



function trimValuesArray(arr) {
  return arr.map(obj => {
    const trimmedObj = {};
    for (const [key, value] of Object.entries(obj)) {
      if (typeof value === 'string') {
        trimmedObj[key] = value.trim();
      } else {
        trimmedObj[key] = value;
      }
    }
    return trimmedObj;
  });
}

//////////////////////////////////////////////////
const checkQualifingCondition = (formData, employeeStatusDataSheet) => {
  // console.log("checkQualifingCondition");
  salesExcelDataSheet.forEach((item) => {
    let numberCheck = 0;
    let Discount = 0;
    let ComplaintCheck = 0;
    let EWCheck = 0;
    let EWPCheck = 0;
    let ExchangeStatusCheck = 0;
    let TotalNumberCheck = 0;
    let CCPcheck = 0;
    let DiscountCount = 0;
    let MSSFcheck = 0;
    let autoCardCheck = 0;
    let obj = {};
    let MSRcheck = 0;

    let carObj = {
      "ALTO": 0,
      "ALTO K-10": 0,
      "S-Presso": 0,
      "CELERIO": 0,
      "WagonR": 0,
      "BREZZA": 0,
      "DZIRE": 0,
      "EECO": 0,
      "Ertiga": 0,
      "SWIFT": 0
    }

    const DSE_NoOfSoldCarExcelDataArr = Object.values(item)[0];
    // check OLD / NEW DSE
    let empStatus = true;
    // console.log(employeeStatusDataSheet)
    employeeStatusDataSheet.forEach(employee => {
      if (employee["DSE ID"] == DSE_NoOfSoldCarExcelDataArr[0]['DSE ID']) {
        if (employee["STATUS"] === "NEW")
          empStatus = false;
      }
    });

    obj = {
      "DSE ID": DSE_NoOfSoldCarExcelDataArr[0]['DSE ID'],
      "DSE Name": DSE_NoOfSoldCarExcelDataArr[0]['DSE Name'],
      "BM AND TL NAME": DSE_NoOfSoldCarExcelDataArr[0]['BM AND TL NAME'],
      "Status": "OLD",
      "Focus Model Qualification": "No",
      ...carObj,
      "Grand Total": 0,
      "Vehicle Incentive ": 0,
      "Special Car Incentive": 0,
      "Total Vehicle Incentive": 0,
      "Super Car Incentive Qualification": 0,
      "Super Car Incentive": 0,
      "CDI Score": 0,
      "CDI Incentive": 0,
      "Extended Warranty Penetration": 0,
      "Extended Warranty Incentive": 0,
      "CCP Score": 0,
      "CCP Incentive": 0,
      "MSSF Score": 0,
      "MSSF Incentive": 0,
      "MSR Score": 0,
      "MSR Incentive": 0,
      "Total Discount": 0,
      "Discount Incentive": 0,
      "Exchange Incentive": 0,
      "Complaint Deduction": 0,
      "MGA/Vehicle": 0,
      "MGA Incentive": 0,
      "Final Incentive": 0,
    }
    if (empStatus) {
      DSE_NoOfSoldCarExcelDataArr.forEach((sold) => {

        Discount = Discount + parseInt(sold["FINAL DISCOUNT"]);
        carObj[sold["Model Name"]]++;
        if (DSE_NoOfSoldCarExcelDataArr[0]['DSE ID'] === "BAD018") {
          // console.log(`carObj[sold["Model Name"]]::::`)
          // console.log(sold["Model Name"], '::::', carObj[sold["Model Name"]])
        }
        if (parseInt(sold["FINAL DISCOUNT"]) > 0) {
          DiscountCount++;
        }
        if (parseInt(sold["CCP PLUS"]) > 0) {
          CCPcheck++;
        }
        if (sold["Financer REMARK"] == "MSSF") {
          MSSFcheck++;
        }
        if (parseInt(sold["Extended Warranty"]) > 0) {
          EWPCheck++;
        }
        if (sold["Exchange Status"] == 'YES' || sold["Exchange Status"] == 'yes') {
          ExchangeStatusCheck++;
        }
        if (sold["Complaint Status"] == 'YES' || sold["Complaint Status"] == 'yes') {
          ComplaintCheck++;
        }
        if (sold["Autocard"] == 'YES' || sold["Autocard"] == 'yes') {
          MSRcheck++;
        }
        TotalNumberCheck++;

        if (formData.QC.focusModel.includes(sold["Model Name"])) {
          numberCheck++;
        }
        if (formData.QC.autoCard == "yes") {
          if (sold["Autocard"] == "YES") {
            autoCardCheck++;
          }
        }
        if (formData.QC.EW == "yes") {
          if (sold["Extended Warranty"] > 0) {
            EWCheck++;
          }
        }
      })

      //for EW and auto card check
      if (numberCheck >= formData.QC.numOfCars) {
        let EWFlag = true;
        let autoCardFlag = true;

        //checking autocard from the excel [form ] 
        if (formData.QC.autoCard === "yes" && (EWCheck >= DSE_NoOfSoldCarExcelDataArr.length))
          autoCardFlag = true;
        else {
          if (formData.QC.autoCard === "yes")
            autoCardFlag = false;
        }
        if (formData.QC.EW === "yes" && (EWCheck >= DSE_NoOfSoldCarExcelDataArr.length))
          EWFlag = true;
        else {
          if (formData.QC.EW === "yes")
            EWFlag = false;
        }
        if (EWFlag && autoCardFlag) {
          // console.log("sdfghgfcvhjkjhv  :  ", obj);
          obj = {
            ...obj,
            ...carObj,
            "Status": "OLD",
            "Focus Model Qualification": "YES",
            "Discount": Discount,
            "AVG. Discount": Discount / DiscountCount,
            "Exchange Status": ExchangeStatusCheck,
            "Complaints": ComplaintCheck,
            "EW Penetration": (EWPCheck / TotalNumberCheck) * 100,
            "MSR": (MSRcheck / TotalNumberCheck) * 100,
            "CCP": (CCPcheck / TotalNumberCheck) * 100,
            "MSSF": (MSSFcheck / TotalNumberCheck) * 100,
            "MSSFCount": MSSFcheck,
            "EWPCount": EWPCheck,
            "MSRCount": MSRcheck,
            "CCPCount": CCPcheck,
            "Grand Total": TotalNumberCheck
          }
          qualifiedRM.push(obj)
        } else {
          obj = {
            ...obj,
            ...carObj,
            "Status": "OLD",
            "Focus Model Qualification": "No",
            "Discount": Discount,
            "AVG. Discount": Discount / DiscountCount,
            "Exchange Status": ExchangeStatusCheck,
            "Complaints": ComplaintCheck,
            "EW Penetration": (EWPCheck / TotalNumberCheck) * 100,
            "MSR": (MSRcheck / TotalNumberCheck) * 100,
            "CCP": (CCPcheck / TotalNumberCheck) * 100,
            "MSSF": (MSSFcheck / TotalNumberCheck) * 100,
            "Grand Total": TotalNumberCheck
          }
          nonQualifiedRM.push(obj)
        }
      }
    } else {
      DSE_NoOfSoldCarExcelDataArr.forEach((sold) => {
        carObj[sold["Model Name"]]++;
        TotalNumberCheck++;

        obj = {
          ...obj,
          ...carObj,
          "Status": "NEW",
          "Focus Model Qualification": "NO",
          "Grand Total": TotalNumberCheck
        }

      })
      newRm.push(obj)

    }

  })
  // console.log("qualifiedRM : ", qualifiedRM)
  // console.log("nonQualifiedRM : ", nonQualifiedRM)


}


function getIncentiveValue(item, key) {
  return typeof (item[key]) === 'number' ? item[key] : 0;
}

ipcMain.on('form-submit', (event, formData) => {

  // console.log("Form Data Input", formData);
  if (!KeyMissing) {
    checkQualifingCondition(formData, employeeStatusDataSheet);
    newDSEIncentiveDataSheet = NewDSEincentiveCalculation(newRm, formData)
    qualifiedRM = PerCarFunc(qualifiedRM, formData);
    qualifiedRM = SpecialCarFunc(qualifiedRM, formData);
    qualifiedRM = PerModelCarFunc(qualifiedRM, formData);//TODO
    qualifiedRM = CDIfunc(qualifiedRM, CDIdata, formData);//TODO
    qualifiedRM = EWfunc(qualifiedRM, formData);
    qualifiedRM = CCPfunc(qualifiedRM, formData);
    qualifiedRM = MSSFfunc(qualifiedRM, formData);
    qualifiedRM = MSRFunc(qualifiedRM, formData);
    qualifiedRM = DiscountFunc(qualifiedRM, formData);
    qualifiedRM = ExchangeFunc(qualifiedRM, formData);
    qualifiedRM = ComplaintFunc(qualifiedRM, formData);
    qualifiedRM = MGAfunc(qualifiedRM, MGAdata, formData);
    qualifiedRM = SuperCarFunc(qualifiedRM, MGAdata, salesExcelDataSheet, formData)
    // console.log("final qualifiedRM ::");
    // console.log(qualifiedRM);
    let finalExcelobjOldDSE = [];

    qualifiedRM.forEach((item) => {

      // if (item["Super Car Incentive"] === 'NaN') {
      //   item["Super Car Incentive"] = 0
      // }
      const grandTotal =
        getIncentiveValue(item, "Total Vehicle Incentive Amt. Slabwise") +
        getIncentiveValue(item, "SpecialCar Incentive") +
        getIncentiveValue(item, "PerModel Incentive") +
        getIncentiveValue(item, "CDI Incentive") +
        getIncentiveValue(item, "EW Incentive") +
        getIncentiveValue(item, "CCP Incentive") +
        getIncentiveValue(item, "MSSF Incentive") +
        getIncentiveValue(item, "MSR Incentive") +
        getIncentiveValue(item, "Discount Incentive") +
        getIncentiveValue(item, "Exchange Incentive") +
        getIncentiveValue(item, "Complaint Deduction") +
        getIncentiveValue(item, "Super Car Incentive") +
        getIncentiveValue(item, "MGA Incentive");

      obj = {
        "DSE ID": item['DSE ID'],
        "DSE Name": item['DSE Name'],
        "BM AND TL NAME": item['BM AND TL NAME'],
        "Status": item["Status"],
        "Focus Model Qualification": item['Focus Model Qualification'],
        "ALTO": item['ALTO'],
        "ALTO K-10": item['ALTO K-10'],
        "S-Presso": item['S-Presso'],
        "CELERIO": item['CELERIO'],
        "WagonR": item['WagonR'],
        "BREZZA": item['BREZZA'],
        "DZIRE": item['DZIRE'],
        "EECO": item['EECO'],
        "Ertiga": item['Ertiga'],
        "SWIFT": item['SWIFT'],
        "Grand Total": item["Grand Total"],
        "Vehicle Incentive ": item["Total PerCar Incentive"],
        "Special Car Incentive": item['SpecialCar Incentive'],
        "Total Vehicle Incentive": item["Total PerCar Incentive"] + item['SpecialCar Incentive'],
        "Super Car Incentive Qualification": getIncentiveValue(item, "Super Car Incentive") ? "YES" : "NO",
        "Super Car Incentive": getIncentiveValue(item, "Super Car Incentive"),
        "CDI Score": getIncentiveValue(item, "CDI Score"),//TODO Handle NAN values
        "CDI Incentive": item["CDI Incentive"],
        "Total MGA": (item['TOTAL MGA']) ? item['TOTAL MGA'] : 0,
        "MGA/Vehicle": Math.round(item["MGA"]),
        "MGA Incentive": Math.round(item["MGA Incentive"]),
        "Exchnage Count": item["Exchange Status"],
        "Exchange Incentive": item["Exchange Incentive"],

        //TODO
        "Extended Warranty Penetration": Math.round(item["EW Penetration"]),
        "Extended Warranty Count": item["EWPCount"],
        "Extended Warranty Incentive": item["EW Incentive"],

        "CCP Score": Math.round(item["CCP"]),
        "CCP Incentive": item["CCP Incentive"],

        //TODO
        "Total Discount": item["Discount"],//TODO Handle value result is not calculating
        "AVG. Discount": item["AVG. Discount"]?item["AVG. Discount"]:0,
        "Vehicle Incentive % Slabwise": item["Vehicle Incentive % Slabwise"],
        "Total Vehicle Incentive Amt. Slabwise": item["Total Vehicle Incentive Amt. Slabwise"],


        "MSSF Score": Math.round(item["MSSF"]),
        "MSSF Incentive": item["MSSF Incentive"],
        "MSR Score": Math.round(item["MSR"]),
        "MSR Incentive": item["MSR Incentive"],
        "Complaint Deduction": item["Complaint Deduction"],//TODO
        "Final Incentive": Math.round(grandTotal),

      }
      finalExcelobjOldDSE.push(obj);
    })

    // finalExcelobjOldDSE = {...nonQualifiedRM,...newDSEIncentiveDataSheet,}


    finalExcelobjOldDSE = [...finalExcelobjOldDSE, ...nonQualifiedRM, ...newRm];
    // console.log("finalExcelobjOldDSE", finalExcelobjOldDSE);
    // console.log("nonQualifiedRM", nonQualifiedRM);


    event.reply("dataForExcel", finalExcelobjOldDSE);
    // event.reply("newDSEIncentiveDataSheet", newDSEIncentiveDataSheet);
    const oldDSE = "oldDSE";
    // const newDSE = "newDSE";
    creatExcel(finalExcelobjOldDSE, oldDSE);
    // creatExcel(newDSEIncentiveDataSheet, newDSE);

    MGAdata = [];
    CDIdata = [];
    salesExcelDataSheet = [];
    employeeStatusDataSheet = [];
    newDSEIncentiveDataSheet = []
    qualifiedRM = [];
    nonQualifiedRM = [];
    newRm = [];
    finalExcelobjOldDSE = []
  }
});

const creatExcel = (dataForExcelObj, text) => {
  // console.log("text :: ", text);
  const nowDate = new Date();
  const month = nowDate.getMonth() + 1;
  const date = nowDate.getDate();
  const year = nowDate.getFullYear();
  const time = nowDate.toLocaleTimeString().replace(/:/g, '-');

  const newWorkbook = XLSX.utils.book_new();
  const newSheet = XLSX.utils.json_to_sheet(dataForExcelObj);
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Sheet1");

  const fileName = `calculatedIncentive_${text}_${date}-${month}-${year}_${time}.xlsx`;
  const folderPath = "./DataSheets";
  if (!fs.existsSync(folderPath)) {
    fs.mkdirSync(folderPath);
    // console.log(`Directory ${folderPath} created.`);
  } else {
    // console.log(`Directory ${folderPath} already exists.`);
  }
  XLSX.writeFile(newWorkbook, `./DataSheets/${fileName}`);

}

ipcMain.on('file-selected-salesExcel', (event, path) => {

  //sales datasheet
  const workbook = XLSX.readFile(path);
  const salesSheetName = workbook.SheetNames[0];
  const salesSheet = workbook.Sheets[salesSheetName];
  let salesSheetData = XLSX.utils.sheet_to_json(salesSheet);
  salesSheetData = transformKeys(salesSheetData);
  salesSheetData = trimValuesArray(salesSheetData);

  const keysToCheckInsalesexcel = ["Model Name", "DSE ID", "DSE Name", "BM AND TL NAME", "Insurance", "Extended Warranty", "Autocard", "CCP PLUS", "FINAL DISCOUNT"
  ];

  const missingKeyForSalesExcel = checkKeys(salesSheetData, keysToCheckInsalesexcel);
  // console.log("missingKeyForSalesExcel")
  // console.log(missingKeyForSalesExcel)
  if (missingKeyForSalesExcel) {
    KeyMissing = true;
    event.reply("formateAlertSalesExcel", missingKeyForSalesExcel);
  }

  //salesExcel
  salesSheetData.shift();
  let groupedData = {};
  salesSheetData.forEach(row => {
    const dseId = row['DSE ID'];
    if (!groupedData[dseId]) {
      groupedData[dseId] = [];
    }
    groupedData[dseId].push(row);
  });
  for (const key in groupedData) {
    if (groupedData.hasOwnProperty(key)) {
      const obj = {};
      obj[key] = groupedData[key];
      salesExcelDataSheet.push(obj);
    }
  }



  //MGA Datasheet
  const MGAsheetName = workbook.SheetNames[2];
  const MGAsheet = workbook.Sheets[MGAsheetName];
  const options = {
    range: 3
  };
  let MGAsheetData = XLSX.utils.sheet_to_json(MGAsheet, options);
  MGAsheetData = transformKeys(MGAsheetData);
  MGAsheetData = trimValuesArray(MGAsheetData);


  const keysToCheckInMGAexcel = ["DSE NAME", "ID", "MGA/VEH", "TOTAL MGA SALE DDL", "MGA SALE FOR ARGRIMENT"];

  const missingKeyForMGAExcel = checkKeys(MGAsheetData, keysToCheckInMGAexcel);
  // console.log("missingKeyForMGAExcel")
  // console.log(missingKeyForMGAExcel)
  if (missingKeyForMGAExcel) {
    KeyMissing = true;
    event.reply("formateAlertMGAExcel", missingKeyForMGAExcel);
  }



  MGAsheetData.forEach((MGArow) => {

    if (MGArow.hasOwnProperty("ID")) {
      MGAdata.push(MGArow);
    }
  })




  //employe Status Sheet
  const employeeStatusSheetName = workbook.SheetNames[3];
  const employeeStatusSheet = workbook.Sheets[employeeStatusSheetName];
  employeeStatusDataSheet = XLSX.utils.sheet_to_json(employeeStatusSheet);

  employeeStatusDataSheet = transformKeys(employeeStatusDataSheet);
  employeeStatusDataSheet = trimValuesArray(employeeStatusDataSheet);


  const keysToCheckInStatusexcel = ["DSE", "STATUS", "DSE ID"];

  const missingKeyForStatusExcel = checkKeys(employeeStatusDataSheet, keysToCheckInStatusexcel);
  // console.log("missingKeyForStatusExcel")
  // console.log(missingKeyForStatusExcel)
  if (missingKeyForStatusExcel) {
    KeyMissing = true;
    event.reply("formateAlertStatusExcel", missingKeyForStatusExcel);
  }


  // console.log("Object inside array employeeStatus", JSON.stringify(employeeStatusDataSheet));


});

ipcMain.on('file-selected-CDIScore', (event, path) => {

  const workbook = XLSX.readFile(path);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const CDIsheetData = XLSX.utils.sheet_to_json(sheet);
  CDIdata = CDIsheetData;
  CDIdata = transformKeys(CDIdata);
  CDIdata = trimValuesArray(CDIdata);
  const keysToCheckInCDIexcel = ["DSE ID", "DSE", "CDI"];

  const missingKeyForCDIExcel = checkKeys(CDIdata, keysToCheckInCDIexcel);
  // console.log("missingKeyForCDIExcel")
  // console.log(missingKeyForCDIExcel)
  if (missingKeyForCDIExcel) {
    KeyMissing = true;
    event.reply("formateAlertCDIExcel", missingKeyForCDIExcel);
  }



  // console.log("Object inside array CDI Score", CDIdata);
});



app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});