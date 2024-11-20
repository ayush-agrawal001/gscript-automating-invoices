# Google Apps Script: Invoice Processing and Bill Creation

This script automates the processing of invoices and the creation of bills using **Google Apps Script** for Google Sheets. It validates input data, handles billing requirements, interacts with an API for invoice generation, and updates the Google Sheet with relevant information.

---

## Script Purpose
- **Validate User Input**: Ensures PAN and GSTIN details are valid.
- **Split Large Amounts**: Breaks down amounts over 50,000 into smaller chunks for processing.
- **Generate and Send Invoices**: Uses an external API to send invoices and retrieves invoice links.
- **Update Google Sheets**: Writes processed data, including invoice links, back to the sheet.

---

## File Type
This is a `.gs` file, **not a `.js` file**. It must be run in Google Apps Script Editor, which is integrated with Google Sheets.

---

## Setup and Configuration

1. **Open Google Apps Script Editor**:
   - In your Google Sheet, navigate to `Extensions > Apps Script`.

2. **Create a New Project**:
   - Copy the script code from `processInvoicesAndCreateBills.gs` into the Apps Script editor.

3. **Authorization**:
   - Apps Script requires permissions to access your Google Sheet and external APIs. Grant these permissions when prompted.

4. **Set API Credentials**:
   - Update the `baseUrl` and `token` in the script with the correct API endpoint and authorization token.

5. **Sheet Structure**:
   - Ensure the Google Sheet contains the following columns in this exact order:
     | Column | Name           | Description                         |
     |--------|----------------|-------------------------------------|
     | A      | Timestamp      | Date of the invoice                |
     | B      | Email          | Customer's email                   |
     | C      | Name           | Customer's name                    |
     | D      | ID Proof       | Type of ID proof (e.g., PAN, GSTIN) |
     | E      | ID Details     | ID details (e.g., GSTIN number)     |
     | F      | State          | Customer's state                   |
     | G      | Pincode        | Customer's pincode                 |
     | H      | Phone          | Customer's phone number            |
     | I      | Total Amount   | Total billing amount               |
     | L      | Bill Sent      | "Yes" if the bill has been sent    |

6. **Run the Script**:
   - Click the `▶` (Run) button in the Apps Script editor to execute the `processInvoicesAndCreateBills` function.

---

## Features

### 1. **Validation**
- Validates PAN using regex: `^[A-Z]{5}[0-9]{4}[A-Z]{1}$`.
- Validates GSTIN using regex: `^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}[Z]{1}[A-Z0-9]{1}$`.

### 2. **Handling Large Amounts**
- Splits amounts > 50,000 into random chunks ≤ 50,000.
- Generates breakdowns for each chunk (e.g., `1000 * qty`, `500 * qty`, etc.).

### 3. **Invoice Generation**
- Sends invoice data to an API and retrieves invoice links.
- Prepares data in a structured format for external API interaction.

### 4. **Sheet Updates**
- Updates the processed rows with:
  - Breakdown of amounts.
  - Pattern of split amounts.
  - Invoice link.
  - Marking the bill as sent.

---

## Important Notes

1. **This Script Runs in Google Sheets**:
   - It is designed for use in the Apps Script environment and **cannot** run in a browser or Node.js.

2. **API Integration**:
   - Ensure the API URL and credentials are correct. If not, the script will fail to generate invoices.

3. **Error Handling**:
   - Logs errors in the Apps Script console (`View > Logs`).
   - Invalid PAN or GSTIN entries are flagged in the Google Sheet.

---

## Example Use Case

1. Add customer data to the Google Sheet.
2. Run the script to validate the data and generate invoices.
3. Check the sheet for updated data, including invoice links and processed statuses.

---

## License
This script is open-source and available for personal and professional use. Modify as needed for your specific requirements.
