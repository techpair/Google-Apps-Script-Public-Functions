/**
 * Sends an email using Gmail service.
 *
 * @param {string} to - The email address of the recipient.
 * @param {string} subject - The subject of the email.
 * @param {string} body - The body of the email.
 */
function sendEmail(to, subject, body) {
  GmailApp.sendEmail(to, subject, body);
}



function extractEmailsToSheet() {
  // Define the search query for the emails you want to extract
  var searchQuery = 'from:example@example.com subject:Enquiry Form'; // Update with your specific search criteria
  
  // Get the active spreadsheet and the active sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Clear any existing content in the sheet
  sheet.clear();
  
  // Set the headers for the sheet
  var headers = ["Date", "From", "Subject", "Body"];
  sheet.appendRow(headers);
  
  // Search for the emails matching the query
  var threads = GmailApp.search(searchQuery);
  var row = 2;
  
  // Loop through each email thread
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    
    // Loop through each message in the thread
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var date = message.getDate();
      var from = message.getFrom();
      var subject = message.getSubject();
      var body = message.getPlainBody();
      
      // Append the message details to the sheet
      sheet.getRange(row, 1).setValue(date);
      sheet.getRange(row, 2).setValue(from);
      sheet.getRange(row, 3).setValue(subject);
      sheet.getRange(row, 4).setValue(body);
      
      row++;
    }
  }
  
  Logger.log("Emails have been successfully extracted to the sheet.");
}

function processEmails() {
  // Get Gmail threads with label "Amazon Order Confirmation" (replace if needed)
  const threads = GmailApp.search("label:Amazon Order Confirmation", GmailApp.SearchOperators.HAS);
  
  // Loop through each thread
  for (const thread of threads) {
    const messages = thread.getMessages();
    const latestMessage = messages[messages.length - 1]; // Assuming latest is confirmation email
    const body = latestMessage.getBody();
    
    // Extract data using regular expressions (adjust patterns as needed)
    const account = extractAccount(body);
    const orderDate = extractDate(body);
    const deliveryDate = extractDeliveryDate(body);
    const productTitle = extractProductTitle(body);
    const orderId = extractOrderId(body);
    const asin = extractAsin(body);
    const itemTotal = extractItemTotal(body);
    const quantity = extractQuantity(body);
    const singlePrice = extractSinglePrice(body);
    
    // Add extracted data to a new row in the sheet
    const sheet = SpreadsheetApp.getActiveSheet();
    const dataRow = sheet.appendRow([account, orderDate, deliveryDate, productTitle, orderId, asin, itemTotal, quantity, singlePrice]);
    
    // Mark thread as read (optional)
    thread.markRead();
  }
}
