var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

function processInvoicesAndCreateBills() {
  var data = sheet.getDataRange().getValues(); // Fetch all rows in the sheet
  
  // Get current Indian date
  var currentIndianDate = getIndianDate();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var timestamp = row[0];    // Column A
    var email = row[1];        // Column B
    var name = row[2];         // Column C
    var idProof = row[3];      // Column D
    var idDetails = row[4];    // Column E
    var state = row[5];        // Column F
    var pincode = row[6];      // Column G
    var phone = row[7];        // Column H
    var totalAmount = row[8];  // Column I
    var isCardPay = row[9];
    var billSent = row[12];    // Column L
    var qty0f1000 = row[13];
    var qty0f500 = row[15];
    var qty0f100 = row[17];
    var qty0fRest = row[19]; 
    var rateOfRest = row[20];
    var pattern = row[21];
    var invoiceLink = row[23]
    
    const breakdown = {
      qty1000: qty0f1000, rate1000: 1000,
      qty500: qty0f500, rate500: 500,
      qty100: qty0f100, rate100: 100,
      qtyRest: qty0fRest, rateRest: rateOfRest
    }
    // Skip if bill already sent
    if (billSent === "Yes") continue;
  
    // Process all cases with valid data
    if (timestamp && name && idProof && idDetails && state && phone) {
      const currentIndianDate = getIndianDate();
      const billData = prepareBillData(row, totalAmount , breakdown, row[6], timestamp);
      const invoiceLinkResult = sendInvoiceToAPI(billData, currentIndianDate, billData.invoiceDate, i + 1);
      updateSheetWithProcessedDataInvoiceApi(sheet, i + 1, invoiceLinkResult);
    } else {  
      if (billSent !== "Yes") {
        sheet.getRange(i + 1, 13).setValue("No");
      }
    }
  }
  Logger.log("Finished processing rows and creating invoices.");
  SpreadsheetApp.flush();
}

function updateSheetWithProcessedDataInvoiceApi(sheet, rowIndex,invoiceLink) {
  sheet.getRange(rowIndex, 23).setValue(invoiceLink);        // Invoice link  
  // sheet.getRange(rowIndex, 13).setValue("Yes");              // Processed status
}


function prepareBillData(row, totalAmount, breakdown, pincode, invoiceDate) {
  // Helper function to shuffle array
  function shuffleArray(array) {
    for (let i = array.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [array[i], array[j]] = [array[j], array[i]];
    }
    return array;
  }

  // Create base items with randomized order
  const baseItems = shuffleArray([
    { name: "Imitation Jewellery", rate: 1000 },
    { name: "Imitation Jewellery", rate: 500 },
    { name: "Imitation Jewellery", rate: 100 },
    { name: "Imitation Jewellery", rate: breakdown.rateRest }
  ]);

  // Map the shuffled base items to actual items with quantities
  const items = baseItems.map(item => {
    switch (item.rate) {
      case 1000:
        return breakdown.qty1000 !== 0 ? {
          name: item.name,
          rate: 1000,
          quantity: breakdown.qty1000,
          gstRate: 3
        } : null;
      case 500:
        return breakdown.qty500 !== 0 ? {
          name: item.name,
          rate: 500,
          quantity: breakdown.qty500,
          gstRate: 3
        } : null;
      case 100:
        return breakdown.qty100 !== 0 ? {
          name: item.name,
          rate: 100,
          quantity: breakdown.qty100,
          gstRate: 3
        } : null;
      default:
        return breakdown.qtyRest !== 0 ? {
          name: item.name,
          rate: breakdown.rateRest,
          quantity: breakdown.qtyRest,
          gstRate: 3
        } : null;
    } 
  }).filter(Boolean);

  console.log(items);

  return {
    invoiceTitle: "VM Jewellers",
    invoiceSubTitle: "VM Jwellers",
    invoiceNumber : row[23] ? row[23] : "",
    contact: {
      phone: "+919739432668",
      email: "contact@example.com",
    },
    invoiceDate: formatDate(invoiceDate),
    dueDate: formatDate(new Date(invoiceDate.getTime() + 7 * 24 * 60 * 60 * 1000)),
    invoiceType: "INVOICE",
    currency: "INR",
    billedTo: {
      name: row[2],
      pincode: pincode,
      // gstState: "07",
      state: row[5],
      country: "IN",
      panNumber: row[3] === "PAN card" ? row[4] : "",
      gstin: row[3] === "GSTIN" ? row[4] : "",
      phone: row[7],
      // phoneShowInInvoice: true, 
      customFields : row[3] === "Aadhar Card" ? [
        {
        dataType: "TEXT",
        hideFlagEmoji : true,
        isArchived : false,
        isDefault : true,
        isHidden : false,
        isRequired : false,
        key : "slws155mu3m",
        label : "Aadhar",
        params : {showInInvoice : true},
        value : row[4],
        }
      ] : []
    },
    billedBy: {
      name: "VM jewellers",
      street: "456 Market Street",  
      pincode: "110002",
      gstState: "07",
      state: "Delhi",
      country: "IN",
      panNumber: "XYZDE5678G",
      gstin: "22XYZDE5678G1Z9",
      phone: "+919123456789", 
      email: "vendor@example.com",
    },
    items: items,
    email: {
      to: {
        name: "Invoice Recipient",
        email: "example@gmail.com",
      },
      cc: [
        {
          name: "CC Recipient 1",
          email: "cc1@example.com",
        },
        {
          name: "CC Recipient 2",
          email: "cc2@example.com",
        },
      ],
    },
  };
}
function sendInvoiceToAPI(billData, currentIndianDate, billDate, rowNo) {
  // Skip future dates
  if (billDate > currentIndianDate) {
    Logger.log(`Skipping future date ${billDate} for split ${rowNo}`);
    return;
  }

  const baseUrl = 'https://api.refrens.com/businesses/url-key/invoices';

  // function getRefrensToken() {
  //   const url = "https://api.refrens.com/authentication";
  //   const payload = {
  //     strategy: "app-secret",
  //     appId: "",
  //     appSecret: ""
  //   };
    
  //   const options = {
  //     method: "post",
  //     contentType: "application/json",
  //     payload: JSON.stringify(payload),
  //   };

  //   const response = UrlFetchApp.fetch(url, options);
  //   const data = JSON.parse(response.getContentText());
  //   // Logger.log(data);
  //   return data.accessToken;
  // }

  // const baseUrl = "https://webhook-test.com/2c7a076c3459b5020b8a06d0add26945";
  // const token2 = getRefrensToken();

  try {
    const options = {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}` },
      contentType: 'application/json',
      payload: JSON.stringify(billData),
    };

    const response = UrlFetchApp.fetch(baseUrl, options);
    const jsonResponse =  JSON.parse(response.getContentText());
    if(jsonResponse.invoiceId && jsonResponse.invoiceNumber){
      sheet.getRange(rowNo, 13).setValue("Yes");
      sheet.getRange(rowNo, 24).setValue(jsonResponse.invoiceNumber);
      sheet.getRange(rowNo, 25).setValue(jsonResponse.invoiceId);
      Logger.log(`Row number : ${rowNo}`);
    }else{
      sheet.getRange(rowNo, 13).setValue("No");
    }
    // return jsonResponse.share.link || "Error";
    return `https://www.refrens.com/app/url-key/invoices/${jsonResponse.invoiceId}` || "Error";
  } catch (error) {
    sheet.getRange(rowNo, 13).setValue("No");
    console.error('Error sending invoice:', error.message);
    return "Error";
  }
}
