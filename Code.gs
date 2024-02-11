/**
 * @fileoverview This file contains the main logic for processing bank emails and extracting transaction information.
 * @module BankEmailProcessor
 */

/**
 * The sheet containing the regex patterns.
 * @constant {GoogleAppsScript.Spreadsheet.Sheet} regexSheet
 * @private
 * @description The sheet containing the regex patterns.
 * This sheet must contain the following columns:
 * - Bank: The name of the bank.
 * - Subject: The subject of the email.
 * - Sender: The email address of the sender.
 * - BodyRegex: The regex pattern to match the email body.
 * - MatchGroups: The names of the match groups in the regex pattern.
 */
const regexSheet = getSheet("Regex");

/**
 * The sheet containing the transaction data. It may be empty or contain existing data.
 * @type {GoogleAppsScript.Spreadsheet.Sheet}
 * @private
 * @description
 * After processing bank emails, the transaction data will be appended to this sheet.
 */
let transactionSheet = getSheet("Transactions");

/**
 * The sheet containing the email data that could not be parsed because the any of the 
 * regex pattern did not match.
 * @type {GoogleAppsScript.Spreadsheet.Sheet}
 * @private
 */
let regexNotMatchedSheet = getSheet("_RegexNotMatched");

/**
 * The sheet containing the email data that could not be parsed because the regex
 * pattern was not found in the regex sheet.
 * @type {GoogleAppsScript.Spreadsheet.Sheet}
 * @private
 */
let regexNotFoundSheet = getSheet("_RegexNotFound");

/**
 * The value to return when the regex pattern is found in the "Regex" sheet
 * but the email body does not match the pattern.
 * @constant {string} REGEX_NOT_MATCHED
 */
const REGEX_NOT_MATCHED = "regex_not_matched";

/**
 * The value to return when the regex pattern is not found in the "Regex" sheet.
 * @constant {string} REGEX_NOT_FOUND
 */
const REGEX_NOT_FOUND = "regex_not_found";

/**
 * The key to use in the regexPatterns map when certain regex patterns are added
 * to the "Regex" sheet without a specific subject or sender.
 * Certain banks may send emails from random email addresses or with subjects
 * containing transaction details.
 */
const UNSPECIFIED_SENDER_OR_SUBJECT = "no_fixed_subject_or_sender";

/**
 * The maximum number of rows to buffer before appending to the sheet.
 * @constant {number} maxBufferSize
 */
const maxBufferSize = 10;

/**
 * The maximum number of email threads to query from gmail in one go.
 * @constant {number} threadLimit
 * @private
 */
const threadLimit = 10;

/**
 * Maximum number of letters to consider from the email subject to create
 * a key for regex lookup. This also helps to load regex for emails that
 * have subject containing dynamic values like merchant name, amount etc.
 * @constant {number} subjectKeyLength
 */
const subjectKeyLength = 42;

/**
 * The map that stores the precompiled regex patterns loaded from the "Regex" sheet.
 * @type {Map}
 */
let regexPatterns = new Map();

/**
 * The map that stores the order of column names in a sheet.
 * The key is the sheet name and the value is an array of column names.
 * @type {Map}
 */
let columnMappings = new Map();

/**
 * Processes bank emails and extracts transaction information.
 * @function ProcessBankEmails
 * @public
 * @todo: Find a way to resume processing from the last processed email.
 */
function ProcessBankEmails() {
    Logger.log("Loading column mappings from sheets");
    loadColumnMappings([transactionSheet, regexNotFoundSheet, regexSheet], columnMappings);
    Logger.log("Loading regex from sheet");
    loadRegexFromSheet(regexSheet);

    // TODO: Find the last processed email date and start from there
    let query = 'label:bank-transaction is:starred AND after:2024/2/1';
    Logger.log(`Processing bank emails: ${query}`);
    let start = 0;
    // Buffer to store transaction sheet rows before appending them to the sheet
    let bufferTransactionRows = [];
    let threads = [];
    do {
        threads.length = 0;
        threads = GmailApp.search(query, start, threadLimit);
        for (let i = 0; i < threads.length; i++) {
            let thread = threads[i];
            for (let j = 0; j < thread.getMessageCount(); j++) {
                let message = thread.getMessages()[j];
                let emailMessage = {
                    messageId: message.getId(),
                    emailDateTime: message.getDate(),
                    mailFrom: message.getFrom(),
                    subject: message.getSubject(),
                    body: message.getPlainBody(),
                    processTime: new Date()
                };
                processEmail(emailMessage, bufferTransactionRows);
            }
        }
        start += threads.length;
    } while (threads.length > 0);
    // Flush remaining rows in buffer to transaction sheet
    if (bufferTransactionRows.length > 0) {
        Logger.log(`Appending buffered ${bufferTransactionRows.length} rows to transaction sheet`);
        // appendMultipleRows(transactionSheet, bufferTransactionRows);
    }
}

/**
 * Extracts transaction information from an email and appends it to the transaction sheet.
 * @param {Object} emailMessage - The email message object.
 * @param {Array[]} bufferTransactionRows - The buffer to store transaction sheet rows before appending them to the sheet.
 * @private
 */
function processEmail(emailMessage, bufferTransactionRows) {
    let transactionInfo = extractTransactionInfo(emailMessage);
    // Logger.log(transactionInfo);
    if (transactionInfo === REGEX_NOT_MATCHED) {
        // Some emails are from known sender with known subject still
        // they do not contain transaction information, so they should be ignored
        appendRow(regexNotMatchedSheet, emailMessage);
    } else if (transactionInfo === REGEX_NOT_FOUND) {
        // TODO: Ask OpenAI to create regex for this email and update the regex in Regex sheet
        appendRow(regexNotFoundSheet, emailMessage);
    } else if (typeof transactionInfo === 'object') {
        appendRow(transactionSheet, transactionInfo, bufferTransactionRows);
    } else {
        Logger.log(`Unknown transaction info: ${transactionInfo}`);
    }
}

/**
 * Appends a row to the sheet using buffer if provided. Flushing buffer is caller's responsibility.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to append rows to.
 * @param {Object} data - The JSON object representing the row data.
 * @param {Array[]} buffer - (Optional) buffer to store rows before appending them to the sheet.
 * @private
 */
function appendRow(sheet, data, buffer) {
    let columns = columnMappings.get(sheet.getName());
    for (let field in data) {
        // If json has new field for which there's no column in sheet, add it
        if (columns.indexOf(field) === -1) {
            columns.push(field);
            columnMappings.set(sheet.getName(), columns);
            // sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
            sheet.getRange(1, columns.length + 1).setValue(field);
        }
    }
    let row = [];
    for (let i = 0; i < columns.length; i++) {
        // Use empty string if the field is not present in the object
        row.push(data[columns[i]] || '');
    }
    if (buffer) {
        buffer.push(row);
        if (buffer.length >= maxBufferSize) {
            appendMultipleRows(sheet, buffer);
            buffer.length = 0;
        }
    } else {
        sheet.appendRow(row);
    }
}

/**
 * Extracts transaction information from an email.
 * @param {Object} emailMessage - The email message object.
 * @returns {(Object|null|string)} - The extracted transaction information object, or string values,
 * "regex_not_matched" or "regex_not_found"
 * @private
 */
function extractTransactionInfo(emailMessage) {
    let senderEmail = getEmailAddress(emailMessage.mailFrom);
    let regexEntries = lookupRegexPatterns(emailMessage.subject, senderEmail);
    for (let i = 0; i < regexEntries.length; i++) {
        let regexEntry = regexEntries[i];
        let match = emailMessage.body.match(regexEntry.BodyRegex);
        if (match) {
            if (regexEntry.MatchGroups.length === 0) {
                return REGEX_NOT_MATCHED;
            }
            let matchGroups = regexEntry.MatchGroups;
            let extractedInfo = {
                'MessageID': emailMessage.messageId,
                'EmailDateTime': emailMessage.emailDateTime,
                'Bank': regexEntry.Bank,
                'ProcessTime': new Date(),
            };
            for (let j = 0; j < matchGroups.length; j++) {
                extractedInfo[matchGroups[j]] = match[j + 1];
            }
            return extractedInfo;
        }
    }
    return REGEX_NOT_FOUND;
}

/**
 * Loads the column names from the specified sheets and stores them in a map.
 * Key is the sheet name and value is an array of column names.
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} sheets - The sheets to load column names from.
 * @param {Map} columnMappings - The map to store the column mappings in.
 * @private
 */
function loadColumnMappings(sheets, columnMappings) {
    sheets.forEach((sheet) => {
        const lastColumn = sheet.getLastColumn();
        let columnNames = [];
        if (lastColumn > 0) {
            columnNames = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
        }
        columnMappings.set(sheet.getName(), columnNames);
    });
}

/**
 * Looks up the relevant regex patterns based on the subject and sender email address.
 * @param {string} subject - The subject of the email.
 * @param {string} sender - The sender's email address.
 * @returns {Object[]} - An array of regex entries.
 * @private
 */
function lookupRegexPatterns(subject, sender) {
    let key = subject.substring(0, subjectKeyLength) + "-" + sender;
    if (regexPatterns.has(key)) {
        return regexPatterns.get(key);
    } else {
        if (regexPatterns.has(UNSPECIFIED_SENDER_OR_SUBJECT)) {
            return regexPatterns.get(UNSPECIFIED_SENDER_OR_SUBJECT);
        } else {
            return [];
        }
    }
}

/**
 * Loads the regex patterns from the "Regex" sheet and stores them in a map.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to load regex patterns from.
 * @private
 * @description Constructs a lookup key using email subject(upto subjectKeyLength characters) and
 * the sender's email address and value contains the regex patterns.
 * Example, "View: Account update for your HDFC Bank A-alerts@hdfcbank.net"
 * This key is used to find the set of regex patterns from the `regexPatterns` Map.
 * This approach avoids iterating through the entire 'Regex' sheet on every email
 * and significantly improves performance.
 */
function loadRegexFromSheet(sheet) {
    let data = sheet.getDataRange().getValues();
    let columns = columnMappings.get(sheet.getName());
    // Skip the first row which contains the column names
    for (let i = 1; i < data.length; i++) {
        let bankName = data[i][columns.indexOf('Bank')];
        let subject = data[i][columns.indexOf('Subject')];
        let sender = data[i][columns.indexOf('Sender')];
        let bodyRegex = new RegExp(data[i][columns.indexOf('BodyRegex')], 's');
        let matchGroups = data[i][columns.indexOf('MatchGroups')].split(',');
        let subjectStartsWithSlash = /^\/.*$/.test(subject);
        let senderStartsWithSlash = /^\/.*$/.test(sender);
        let key = subject.substring(0, subjectKeyLength) + "-" + sender;
        if (subjectStartsWithSlash || senderStartsWithSlash) {
            // If the subject or sender starts with a slash, then the user 
            // provides regex pattern for the subject or sender.
            key = UNSPECIFIED_SENDER_OR_SUBJECT;
        }
        let regexPattern = {
            'Bank': bankName,
            'Subject': subject,
            'Sender': sender,
            'BodyRegex': bodyRegex,
            'MatchGroups': matchGroups
        };

        if (regexPatterns.has(key)) {
            regexPatterns.get(key).push(regexPattern);
        } else {
            regexPatterns.set(key, [regexPattern]);
        }
    }
}

/**
 * Appends multiple rows to a sheet at once.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to append rows to.
 * @param {Array[]} rows - The rows to append.
 * @private
 */
function appendMultipleRows(sheet, rows) {
    let lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * Regular expression to extract only the email address from the "From" field of an email.
 * @constant {string} emailRegex
 */
const emailRegex = new RegExp(/<([^>]+)>/);

/**
 * Extracts only the email address from the "From" field of an email.
 * @param {string} mailFrom - The "From" field of the email.
 * @returns {string} - The extracted email address.
 * @example "HDFC Bank InstaAlerts <alerts@hdfcbank.net>" -> "alerts@hdfcbank.com"
 * @private
 */
function getEmailAddress(mailFrom) {
    let match = emailRegex.exec(mailFrom);
    return match ? match[1] : mailFrom;
}


/**
 * Returns the sheet with the given name or creates a new sheet if not found.
 * @param sheetName Name of the sheet to get or create.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet with the given name.
 */
function getSheet(sheetName) {
    try {
        let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
            Logger.log(`Creating sheet ${sheetName}`);
            sheet = spreadsheet.insertSheet(sheetName);
        }
        return sheet;
    } catch (error) {
        Logger.log(`Error getting sheet ${sheetName}: ${error}`);
        return null;
    }
}
