const token = '';

function isValidGSTIN(gstin) {
  var gstinRegex = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}[Z]{1}[A-Z0-9]{1}$/;
  return gstinRegex.test(gstin);
}

function getIndianDate() {
  // Get current time in UTC
  var currentTime = new Date();
  // Convert to Indian Standard Time
  var indianTime = new Date(currentTime.toLocaleString("en-US", { timeZone: "Asia/Kolkata" }));
  // Reset time to start of the day
  indianTime.setHours(0, 0, 0, 0);
  return indianTime;
}

function verifyGSTIN(gstin) {
  const baseUrl = 'https://api.refrens.com/businesses/url-key/irn/';

  try {
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(baseUrl + gstin, options);
    return JSON.parse(response.getContentText());
  } catch (error) {
    Logger.log('Error verifying GSTIN:', error.message);
    return null;
  }
}

function isValidAadhar(aadharNumber) {
  // Regex for Aadhar number validation:
  // - 12 digits only
  // - No specific validation of Verhoeff algorithm (you can add that later if needed)
  const aadharRegex = /^\d{12}$/;
  
  if (!aadharNumber) return false;
  
  if (aadharNumber.length === 0) return false;
  // Remove any spaces or hyphens
  aadharNumber = String(aadharNumber);
  const cleanedAadhar = aadharNumber.replace(/[\s-]/g, '');
  
  // Check regex and ensure no repeated digits
  if (!aadharRegex.test(cleanedAadhar)) return false;
  
  // Optional: Check for obvious invalid patterns
  const invalidPatterns = [
    '000000000000', 
    '111111111111', 
    '222222222222', 
    '333333333333', 
    '444444444444', 
    '555555555555', 
    '666666666666', 
    '777777777777', 
    '888888888888', 
    '999999999999'
  ];
  
  return !invalidPatterns.includes(cleanedAadhar);
}

function isValidPhoneNumber(phoneNumber) {
  // Handle null or undefined
  if (phoneNumber == null || phoneNumber == undefined) {
    return false;
  }

  // Convert to string and remove all non-digit characters
  const phoneStr = String(phoneNumber)
    .trim()
    .replace(/[^\d]/g, '');
  
  // If no digits remain, return false
  if (phoneStr.length === 0) return false;

  // Remove any leading country code
  const cleanedPhone = phoneStr.replace(/^(?:\+91)/, '');
  
  // Validate final cleaned number
  // - Must be exactly 10 digits
  // - Must start with 6, 7, 8, or 9
  const phoneRegex = /^[6-9]\d{9}$/;
  
  // Check length and first digit
  if (cleanedPhone.length !== 10) return false;
  
  // Check against regex
  if (!phoneRegex.test(cleanedPhone)) return false;
  
  // Optional: Check for obviously invalid patterns
  const invalidPatterns = [
    '0000000000', 
    '1111111111', 
    '2222222222', 
    '3333333333', 
    '4444444444', 
    '5555555555', 
    '6666666666', 
    '7777777777', 
    '8888888888', 
    '9999999999'
  ];
  
  // Final check: ensure it's not an invalid pattern
  return !invalidPatterns.includes(cleanedPhone);
}

// Helper function to standardize phone number format
function standardizePhoneNumber(phoneNumber) {
  // Handle null or undefined
  if (phoneNumber == null || phoneNumber == undefined) {
    return '';
  }

  // Convert to string and remove all non-digit characters
  const phoneStr = String(phoneNumber)
    .trim()
    .replace(/[^\d]/g, '');
  
  // If no digits remain, return empty string
  if (phoneStr.length === 0) return '';

  // Remove any leading country code
  const cleanedPhone = phoneStr.replace(/^(?:\+91)/, '');
  
  // Validate and return
  return isValidPhoneNumber(cleanedPhone) ? cleanedPhone : '';
}


function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
}

function generateNewRowData(row, splitAmount, breakdown, pattern, date) {
  return [
    date,  // Use the new date instead of the original timestamp
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
    "",
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