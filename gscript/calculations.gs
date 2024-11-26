function calculateAndVerify(sheet) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Dynamically get the latest data each iteration
  let data = sheet.getDataRange().getValues();
  const currentIndianDate = getIndianDate();

  // Use a while loop instead of a for loop to handle dynamic row changes
  let i = 1;
  while (i < data.length) {
    // Refresh data at the start of each iteration to capture newly added rows
    data = sheet.getDataRange().getValues();
    
    let row = data[i];
    let timestamp = row[0];    // Column A
    let email = row[1];        // Column B
    let name = row[2];         // Column C
    let idProof = row[3];      // Column D
    let idDetails = row[4];    // Column E
    let state = row[5];        // Column F
    let pincode = row[6];      // Column G
    let phone = row[7];        // Column H
    let totalAmount = row[8];  // Column I
    let isCardPay = row[9];
    let billSent = row[12];    // Column L

    // Validate Phone Number
    if (phone && !isValidPhoneNumber(phone)) {
      Logger.log(`Invalid Phone Number for row ${i + 1}: ${phone}`);
      sheet.getRange(i + 1, 13).setValue("Invalid Phone Number");
      i++;
      continue;
    }

    // Validate Aadhar Number if ID Proof is Aadhar
    if (idProof === "Aadhar Card" && (!idDetails || !isValidAadhar(idDetails))) {
      Logger.log(`Invalid Aadhar Number for row ${i + 1}: ${idDetails}`);
      sheet.getRange(i + 1, 13).setValue("Invalid Aadhar");
      i++;
      continue;
    }

    // Skip if bill already sent
    if (billSent === "Yes") {
      i++;
      continue;
    }

    // Check for matching phone number in previous rows
    if (phone) {
      for (let j = 1; j < i; j++) {
        const previousRow = data[j];
        const previousPhone = previousRow[7];
        
        if (phone === previousPhone && previousRow[1]) {
          email = previousRow[1];
          name = previousRow[2];
          idProof = previousRow[3];
          idDetails = previousRow[4];
          state = previousRow[5];
          pincode = previousRow[6];
          
          sheet.getRange(i + 1, 2, 1, 6).setValues([[
            email, name, idProof, idDetails, state, pincode
          ]]);
          
          row[1] = email;
          row[2] = name;
          row[3] = idProof;
          row[4] = idDetails;
          row[5] = state;
          row[6] = pincode;
          break;
        }
      }
    }

    // Validate ID Proof and details
    if (idProof === "PAN card" && !isValidPAN(idDetails)) {
      Logger.log(`Invalid PAN for row ${i + 1}: ${idDetails}`);
      sheet.getRange(i + 1, 13).setValue("Invalid PAN");
      i++;
      continue;
    } else if (idProof === "GSTIN") {
      if (!isValidGSTIN(idDetails)) {
        Logger.log(`Invalid GSTIN for row ${i + 1}: ${idDetails}`);
        sheet.getRange(i + 1, 13).setValue("Invalid GSTIN");
        i++;
        continue;
      }
      
      // Verify GSTIN and update details
      const gstinData = verifyGSTIN(idDetails);
      if (gstinData && gstinData.status) {
        name = gstinData.name || name;
        pincode = gstinData.pincode || pincode;
        
        sheet.getRange(i + 1, 3, 1, 4).setValues([[
          name, "GSTIN", idDetails, state
        ]]);
        row[2] = name;
        row[6] = pincode;
      }
    }

    
    if( isCardPay === "Yes" ){
      sheet.getRange(i + 1, 10).setValue("No");
      function reducePaymentAmount(totalAmount) {
        // Card payment reduction factor
        const CARD_PAYMENT_FACTOR = 1.0506;
        
        // Calculate the reduced amount
        const reducedAmount = totalAmount / CARD_PAYMENT_FACTOR;
        
        // Round to two decimal places for monetary precision
        return Math.round(reducedAmount * 100) / 100;
      }
      totalAmount = reducePaymentAmount(totalAmount);
    }
    console.log(totalAmount)

    // Calculate and validate amounts
    if (timestamp && name && idProof && idDetails && state && phone) {
      if (totalAmount > 50000) {
        // Store the original row length before splitting
        const originalRowCount = data.length;
        
        // Perform split bill calculation
        calculateSplitBill(sheet, row, i, currentIndianDate, totalAmount);
        
        // Get the updated data after splitting
        data = sheet.getDataRange().getValues();
        
        // Adjust the index based on how many rows were added
        const rowsAdded = data.length - originalRowCount;
        i += rowsAdded;
      } else {
        calculateRegularBill(sheet, row, i, currentIndianDate, totalAmount);
        i++;
      }
    } else {
      if (billSent !== "Yes") {
        sheet.getRange(i + 1, 13).setValue("No");
      }
      i++;
    }
  }
  
  Logger.log("Finished calculation and verification.");
  SpreadsheetApp.flush();
}
function calculateSplitBill(sheet, row, rowIndex, currentIndianDate, totalAmount) {
  const splits = distributeRandomly(totalAmount, 49000);
  const baseDate = new Date(row[0]); // Get the original date
  const validSplits = []; // To store valid splits

  // Populate the validSplits array with details
  splits.forEach((splitAmount, index) => {
    const splitDate = new Date(baseDate);
    splitDate.setDate(splitDate.getDate() + (index * 2));

    const result = calculatePattern(splitAmount);
    const breakdown = result.breakdown;
    const pattern = result.pattern.join(" + ");

    validSplits.push({
      amount: Math.round(splitAmount),
      date: splitDate,
      breakdown: breakdown,
      pattern: pattern,
    });
  });

  // Process and update the sheet based on valid splits
  validSplits.forEach((split, index) => {
    if (index === 0) {
      // Update the original row
      updateSheetWithProcessedData(sheet, rowIndex + 1, split.amount, split.breakdown, split.pattern);
    } else {
      // Add new rows for additional splits
      sheet.insertRowAfter(rowIndex + index);
      const newRow = generateNewRowData(row, split.amount, split.breakdown, split.pattern, split.date);
      sheet.getRange(rowIndex + index + 1, 1, 1, newRow.length).setValues([newRow]);
    }
  });

  // Log the total number of valid splits
  Logger.log(`Total valid splits: ${validSplits.length}`);
}


function calculatePattern(amount) {
  const denominations = [1000, 500, 100];
  let remaining = amount;
  let breakdownData = {
    pattern: [],
    breakdown: {
      qty1000: 0, rate1000: 1000,
      qty500: 0, rate500: 500,
      qty100: 0, rate100: 100,
      qtyRest: 0, rateRest: 0
    }
  };

  // Helper function for random numbers
  function getRandomInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }

  // Process each denomination except the last one
  for (let i = 0; i < denominations.length - 1; i++) {
    const denom = denominations[i];
    const maxPossible = Math.floor(remaining / denom);
    
    if (maxPossible > 0) {
      // Leave some room for smaller denominations
      const minCount = Math.max(0, maxPossible - 20);
      const count = getRandomInt(minCount, maxPossible);
      
      if (count > 0) {
        breakdownData.pattern.push(`${count} x ${denom}`);
        breakdownData.breakdown[`qty${denom}`] = count;
        breakdownData.breakdown[`rate${denom}`] = denom;
        remaining -= count * denom;
      }
    }
  }

  // Handle the last denomination (100)
  const lastDenom = denominations[denominations.length - 1];
  const lastCount = Math.floor(remaining / lastDenom);
  if (lastCount > 0) {
    breakdownData.pattern.push(`${lastCount} x ${lastDenom}`);
    breakdownData.breakdown[`qty${lastDenom}`] = lastCount;
    breakdownData.breakdown[`rate${lastDenom}`] = lastDenom;
    remaining -= lastCount * lastDenom;
  }

  // Add any remaining amount
  if (remaining > 0) {
    breakdownData.pattern.push(`${remaining}`);
    breakdownData.breakdown.qtyRest = 1;
    breakdownData.breakdown.rateRest = remaining;
  }

  return breakdownData;
}


function distributeRandomly(totalAmount, maxAmount) {
  const MIN_SPLIT = 38000;
  const MAX_SPLIT = 48000;
  var splits = [];

  // If total amount is less than minimum split, return the entire amount
  if (totalAmount <= MIN_SPLIT) {
    return [totalAmount];
  }

  while (totalAmount > 0) {
    // Calculate the split amount
    var split;
    
    // If remaining amount is between MIN_SPLIT and MAX_SPLIT, use the entire amount
    if (totalAmount <= MAX_SPLIT) {
      split = totalAmount;
    } 
    // Otherwise, generate a random split between MIN_SPLIT and MAX_SPLIT
    else {
      split = Math.floor(Math.random() * (MAX_SPLIT - MIN_SPLIT + 1)) + MIN_SPLIT;
    }

    // Ensure we don't exceed the total amount
    if (split > totalAmount) {
      split = totalAmount;
    }

    splits.push(split);
    totalAmount -= split;
  }

  return splits;
}

function updateSheetWithProcessedData(sheet, rowIndex, totalAmount, breakdown, pattern, invoiceLink) {
  sheet.getRange(rowIndex, 9).setValue(totalAmount);         // Total Amount
  sheet.getRange(rowIndex, 14).setValue(breakdown.qty1000);  // Qty for 1000
  sheet.getRange(rowIndex, 15).setValue(breakdown.rate1000); // Rate for 1000
  sheet.getRange(rowIndex, 16).setValue(breakdown.qty500);   // Qty for 500
  sheet.getRange(rowIndex, 17).setValue(breakdown.rate500);  // Rate for 500
  sheet.getRange(rowIndex, 18).setValue(breakdown.qty100);   // Qty for 100
  sheet.getRange(rowIndex, 19).setValue(breakdown.rate100);  // Rate for 100
  sheet.getRange(rowIndex, 20).setValue(breakdown.qtyRest);  // Qty for Rest
  sheet.getRange(rowIndex, 21).setValue(breakdown.rateRest); // Rate for Rest
  sheet.getRange(rowIndex, 22).setValue(pattern);            // Pattern
  // sheet.getRange(rowIndex, 13).setValue("Yes");              // Processed status
}

function calculateRegularBill(sheet, row, rowIndex, currentIndianDate, totalAmount) {
  totalAmount = Math.round(totalAmount);
  const result = calculatePattern(totalAmount);
  const breakdown = result.breakdown;
  const pattern = result.pattern.join(" + ");

  updateSheetWithProcessedData(sheet, rowIndex + 1, totalAmount, breakdown, pattern);
}
