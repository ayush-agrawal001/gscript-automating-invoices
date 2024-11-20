function processInvoicesAndCreateBills() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues(); // Fetch all rows in the sheet

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var timestamp = row[0];    // Column A
    var email = row[1];        // Column B
    var name = row[2];         // Column C
    var idProof = row[3];      // Column D
    var idDetails = row[4];    // Column E (GSTIN if idProof is GSTIN)
    var state = row[5];        // Column F
    var pincode = row[6];      // Column G
    var phone = row[7];        // Column H
    var totalAmount = row[8];  // Column I
    var billSent = row[11];    // Column L

    // Skip if bill already sent
    if (billSent === "Yes") continue;

    // Validate ID Proof and details
    if (idProof === "PAN card" && !isValidPAN(idDetails)) {
      Logger.log(`Invalid PAN for row ${i + 1}: ${idDetails}`);
      sheet.getRange(i + 1, 12).setValue("Invalid PAN");
      continue;
    } else if (idProof === "GSTIN" && !isValidGSTIN(idDetails)) {
      Logger.log(`Invalid GSTIN for row ${i + 1}: ${idDetails}`);
      sheet.getRange(i + 1, 12).setValue("Invalid GSTIN");
      continue;
    }

    // Process amounts over 50,000 by splitting
    if (totalAmount > 50000) {
      var splits = distributeRandomly(totalAmount, 50000);
      for (var j = 0; j < splits.length; j++) {
        var splitAmount = Math.round(splits[j]); // Ensure split amount is an integer
        var breakdown = calculateBreakdown(splitAmount);
        var pattern = calculatePattern(splitAmount);

        // Prepare bill data for API
        const billData = prepareBillData(row, splitAmount, breakdown, pincode);

        // Send bill and get invoice link
        const invoiceLink = sendInvoiceToAPI(billData);

        // Update sheet for this split
        if (j === 0) {
          // Update the original row with the first split
          updateSheetWithProcessedData(sheet, i + 1, splitAmount, breakdown, pattern, invoiceLink);
        } else {
          // Insert new rows for additional splits
          sheet.insertRowAfter(i + j);
          var newRow = generateNewRowData(row, splitAmount, breakdown, pattern);
          sheet.getRange(i + j + 1, 1, 1, newRow.length).setValues([newRow]);
          sheet.getRange(i + j + 1, 12).setValue("Yes");
          sheet.getRange(i + j + 1, 22).setValue(invoiceLink);
        }
      }
      continue;
    }

    // Process amounts less than or equal to 50,000
    if (timestamp && email && name && idProof && idDetails && state && phone && totalAmount <= 50000 && billSent !== "Yes") {
      totalAmount = Math.round(totalAmount);
      var breakdown = calculateBreakdown(totalAmount);
      var pattern = calculatePattern(totalAmount);

      // Prepare bill data for API
      const billData = prepareBillData(row, totalAmount, breakdown, pincode);

      // Send bill and get invoice link
      const invoiceLink = sendInvoiceToAPI(billData);

      // Update sheet with processed data
      updateSheetWithProcessedData(sheet, i + 1, totalAmount, breakdown, pattern, invoiceLink);
    } else {
      if (billSent !== "Yes") {
        sheet.getRange(i + 1, 12).setValue("No");
      }
    }
  }
  Logger.log("Finished processing rows and creating invoices.");
  SpreadsheetApp.flush();
}

// Helper function to prepare bill data for API
function prepareBillData(row, totalAmount, breakdown, pincode) {
  return {
    invoiceTitle: "VM Jewellers",
    invoiceSubTitle: "VM Jwellers",
    contact: {
      phone: "+919739432668",
      email: "contact@example.com",
    },
    invoiceDate: formatDate(new Date(row[0])),
    dueDate: formatDate(new Date(new Date(row[0]).getTime() + 7 * 24 * 60 * 60 * 1000)),
    invoiceType: "INVOICE",
    currency: "INR",
    billedTo: {
      name: row[2],
      pincode: pincode,
      gstState: "07",
      state: row[5],
      country: "IN",
      panNumber: row[3] === "PAN card" ? row[4] : "",
      gstin: row[3] === "GSTIN" ? row[4] : "",
      phone: row[7],
      email: row[1],
    },
    billedBy: {
      name: "Agrawal and Associates",
      street: "456 Market Street",
      pincode: "110002",
      gstState: "03",
      state: "Punjab",
      country: "IN",
      panNumber: "XYZDE5678G",
      gstin: "22XYZDE5678G1Z9",
      phone: "+919123456789",
      email: "vendor@example.com",
    },
    items: [
      breakdown.qty1000 !== 0 && {
        name: "Item 1",
        rate: 1000,
        quantity: breakdown.qty1000,
        gstRate: 3,
      },
      breakdown.qty500 !== 0 && {
        name: "Item 2",
        rate: 500,
        quantity: breakdown.qty500,
        gstRate: 3,
      },
      breakdown.qty100 !== 0 && {
        name: "Item 3",
        rate: 100,
        quantity: breakdown.qty100,
        gstRate: 3,
      },
      breakdown.qtyRest !== 0 && {
        name: "Item 4",
        rate: breakdown.rateRest,
        quantity: breakdown.qtyRest,
        gstRate: 3,
      },
    ].filter(Boolean),
    email: {
      to: {
        name: "Invoice Recipient",
        email: "thebestayush62@gmail.com",
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

// Helper function to send invoice to API
function sendInvoiceToAPI(billData) {
  const baseUrl = 'https://api.refrens.com/businesses/your-username/invoices';
  const token = 'your-jwt-token';

  try {
    const options = {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}` },
      contentType: 'application/json',
      payload: JSON.stringify(billData),
    };

    const response = UrlFetchApp.fetch(baseUrl, options);
    const jsonResponse = JSON.parse(response.getContentText());

    return jsonResponse.share.link || "Error";
  } catch (error) {
    console.error('Error sending invoice:', error.message);
    return "Error";
  }
}

// Helper function to update sheet with processed data
function updateSheetWithProcessedData(sheet, rowIndex, totalAmount, breakdown, pattern, invoiceLink) {
  sheet.getRange(rowIndex, 9).setValue(totalAmount);         // Total Amount
  sheet.getRange(rowIndex, 13).setValue(breakdown.qty1000);  // Qty for 1000
  sheet.getRange(rowIndex, 14).setValue(breakdown.rate1000); // Rate for 1000
  sheet.getRange(rowIndex, 15).setValue(breakdown.qty500);   // Qty for 500
  sheet.getRange(rowIndex, 16).setValue(breakdown.rate500);  // Rate for 500
  sheet.getRange(rowIndex, 17).setValue(breakdown.qty100);   // Qty for 100
  sheet.getRange(rowIndex, 18).setValue(breakdown.rate100);  // Rate for 100
  sheet.getRange(rowIndex, 19).setValue(breakdown.qtyRest);  // Qty for Rest
  sheet.getRange(rowIndex, 20).setValue(breakdown.rateRest); // Rate for Rest
  sheet.getRange(rowIndex, 21).setValue(pattern);            // Pattern
  sheet.getRange(rowIndex, 12).setValue("Yes");              // Processed status
  sheet.getRange(rowIndex, 22).setValue(invoiceLink);        // Invoice link
}

// Helper function to generate new row data for splits
function generateNewRowData(row, splitAmount, breakdown, pattern) {
  return [
    formatDate(new Date(row[0])),
    row[1],
    row[2],
    row[3],
    row[4],
    row[5],
    row[6],
    row[7],
    splitAmount,
    "",
    "",
    "Yes",
    breakdown.qty1000,
    breakdown.rate1000,
    breakdown.qty500,
    breakdown.rate500,
    breakdown.qty100,
    breakdown.rate100,
    breakdown.qtyRest,
    breakdown.rateRest,
    pattern
  ];
}

// Existing helper functions from previous code
function isValidPAN(pan) {
  var panRegex = /^[A-Z]{5}[0-9]{4}[A-Z]{1}$/;
  return panRegex.test(pan);
}

function isValidGSTIN(gstin) {
  var gstinRegex = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}[Z]{1}[A-Z0-9]{1}$/;
  return gstinRegex.test(gstin);
}

function calculateBreakdown(amount) {
  const denominations = [1000, 500, 100];
  let breakdown = {
    qty1000: 0, rate1000: 0,
    qty500: 0, rate500: 0,
    qty100: 0, rate100: 0,
    qtyRest: 0, rateRest: 0
  };

  for (let denom of denominations) {
    if (amount >= denom) {
      let qty = Math.floor(amount / denom);
      breakdown[`qty${denom}`] = qty;
      breakdown[`rate${denom}`] = denom;
      amount %= denom;
    }
  }

  if (amount > 0) {
    breakdown.qtyRest = 1;
    breakdown.rateRest = amount;
  }

  return breakdown;
}

function calculatePattern(amount) {
  const denominations = [1000, 500, 100];
  let result = [];
  for (let denom of denominations) {
    let count = Math.floor(amount / denom);
    if (count > 0) {
      result.push(`${count} * ${denom}`);
    }
    amount %= denom;
  }
  if (amount > 0) {
    result.push(`${amount}`);
  }
  return result.join(" + ");
}

function distributeRandomly(totalAmount, maxAmount) {
  var splits = [];
  while (totalAmount > 0) {
    var split = Math.min(maxAmount, Math.floor(Math.random() * maxAmount) + 1);
    if (split > totalAmount) {
      split = totalAmount;
    }
    splits.push(split);
    totalAmount -= split;
  }
  return splits;
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
}
