function processPayments() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headersRow = 1; // Assuming headers are on the first row
  const dataRange = sheet.getDataRange(); // Get the entire data range
  const data = dataRange.getValues(); // Get all values in the range
  const urlKey = 'url-key';

  const headers = {
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json',
  };

  // Loop through rows (skip the header row)
  for (let i = headersRow; i < data.length; i++) {
    try {
      const invoiceId = data[i][24]; // Invoice ID is in column X (0-based index 23)
      const paymentDate = data[i][11]; // Payment Date is in column L (0-based index 11)
      if (!invoiceId || !paymentDate) {
        Logger.log(`Skipping row ${i + 1}: Missing data.`);
        continue;
      }

      // Step 1: GET request to retrieve payerBusiness
      const getUrl = `https://api.refrens.com/businesses/${urlKey}/invoices/${invoiceId}`;
      const getResponse = UrlFetchApp.fetch(getUrl, { method: 'get', headers });
      const getData = JSON.parse(getResponse.getContentText());
      const payerBusiness = getData.business?._id;
      const totalAmount = getData.finalTotal?.total;
      Logger.log(payerBusiness);
      Logger.log(totalAmount);
      if (!payerBusiness) {
        Logger.log(`Row ${i + 1}: payerBusiness ID not found.`);
        continue;
      }

      Logger.log(`Row ${i + 1}: Payer Business ID: ${payerBusiness}`);

      // Step 2: POST request to process the payment
      const postUrl = `https://api.refrens.com/businesses/${urlKey}/invoices/${invoiceId}/payments`;
      const payload = {
        tds: 0,
        transactionCharge: 0,
        finalAmount: totalAmount,
        amount: totalAmount,
        paymentDate: new Date(paymentDate).toISOString(),
        paymentMethod: "UPI",
        paymentNotificationEmail: null,
        paymentLedgerId: null,
        mailToClient: "NO",
        paymentAccount: "673fa0926200310027e29a4d",
        payerBusiness: payerBusiness,
      };

      const postResponse = UrlFetchApp.fetch(postUrl, {
        method: 'post',
        headers,
        payload: JSON.stringify(payload),
        contentType: 'application/json',
      });

      Logger.log(`Row ${i + 1}: Payment Response: ${postResponse.getContentText()}`);
    } catch (error) {
      Logger.log(`Row ${i + 1}: Error - ${error.message}`);
    }
  }
}
