/**
 * Main function:
 * 1. Reads "InputFolderId" and "OutputFolderName" from Script Properties.
 * 2. Ensures there is a subfolder named "OutputFolderName" inside the input folder.
 * 3. Checks for new files in the input folder, processes them, and stores
 *    artifacts in the subfolder.
 */
function main() {
  var scriptProperties = PropertiesService.getScriptProperties();

  // 1. Read the folder ID and subfolder name from script properties
  var inputFolderId = scriptProperties.getProperty('InputFolderId');
  if (!inputFolderId) {
    Logger.log("No 'InputFolderId' property found or it is empty. Exiting...");
    return;
  }
  var outputFolderName = scriptProperties.getProperty('OutputFolderName');
  if (!outputFolderName) {
    Logger.log("No 'OutputFolderName' property found or it is empty. Exiting...");
    return;
  }

  // 2. Access the input folder and create/find a subfolder within it
  var inputFolder = DriveApp.getFolderById(inputFolderId);
  var outputFolder = findOrCreateSubfolder(inputFolder, outputFolderName);

  // 3. Identify new files in the input folder
  var files = inputFolder.getFiles();
  var knownFiles = scriptProperties.getProperty('knownFileIDs');
  knownFiles = knownFiles ? JSON.parse(knownFiles) : [];

  var runMode = scriptProperties.getProperty('RunMode');
  if (runMode === 'test') {
    Logger.log('Running in test mode');
  }
  var newFiles = [];
  while (files.hasNext()) {
    var file = files.next();
    var fileId = file.getId();
    if (knownFiles.indexOf(fileId) === -1) {
      newFiles.push(file);
    }
  }

  // 4. Update known file IDs
  Logger.log('found ' + knownFiles.length + ' known and ' + newFiles.length + ' new');

  // 5. Process new files, placing artifacts in the subfolder
  var spreadsheet = findOrCreateSheetInFolder(outputFolder);

  newFiles.forEach(function(file) {
    try {
      processNewFile(file, outputFolder, spreadsheet);
      if (runMode !== 'test') {
        knownFiles.push(file.getId());
      }
    } catch (e) {
      Logger.log("Error processing file: " + file.getName() + " (" + file.getId() + "): " + e.toString());
    }
  });
  if (runMode !== 'test') {
    scriptProperties.setProperty('knownFileIDs', JSON.stringify(knownFiles));
  }
}

/**
 * Find or create a subfolder with the given name inside `parentFolder`.
 * Returns the found or newly created folder.
 */
function findOrCreateSubfolder(parentFolder, subfolderName) {
  // Check if it already exists in parentFolder
  var subFolders = parentFolder.getFolders();
  while (subFolders.hasNext()) {
    var folder = subFolders.next();
    if (folder.getName() === subfolderName) {
      Logger.log("Found existing subfolder: " + folder.getName() + " (" + folder.getId() + ")");
      return folder;
    }
  }
  // Otherwise, create a new subfolder
  try {
    var newSubFolder = parentFolder.createFolder(subfolderName);
    Logger.log("Created subfolder: " + newSubFolder.getName() + " (" + newSubFolder.getId() + ")");
  } catch (e) {
    Logger.log("Error creating subfolder: " + e.toString());
  }
  return newSubFolder;
}

/**
 * Process each new file:
 * 1. Extract content (OCR if PDF, direct read if Google Doc, etc.).
 * 2. Extract sum/num/date from text.
 * 3. Create .txt file in the output folder.
 * 4. Append a row to the "Extraction Log" (also in the output folder).
 */
function processNewFile(file, outputFolder, spreadsheet) {
  var fileName = file.getName();
  Logger.log("New file to process: " + fileName);

  // 1) Extract file content
  var fileContent = getFileContent(file);
  if (!fileContent) {
    Logger.log("No file content was extracted. Skipping...");
    return;
  }

  // 2) Extract needed fields via regex
  var patterns = [
    { tag: "sum", pattern: /סה"?כ (?:(?:בשח:\s*)|(?:לתשלום:?))\s+(?<value>\d+(?:.\d+)?)/gm },
    { tag: "num", pattern: /(?:חשבונית מס\/קבלה(?:\D|\s)*(?<value>\d+))/gm },
    { tag: "date", pattern: /(?<value>(0[1-9]|1\d|2[0-8]|29(?=(\/|\.)\d\d(\/|\.)(?!1[01345789]00|2[1235679]00)\d\d(?:[02468][048]|[13579][26]))|30(?!(\/|\.)02)|31(?=(\/|\.)0[13578]|(\/|\.)1[02]))(\/|\.)(?:0[1-9]|1[0-2])(\/|\.)(?:([12]\d{3})|\d{2}))/gm}
  ];

  var extracted = {};
  patterns.forEach(function(item) {
    v = matchit(item.tag, item.pattern, fileContent);
    extracted[item.tag] = v? v : "N/A";
  });

  // 3) Create a .txt file in the output folder
  var textFileLink = createTextFile(fileName, fileContent, outputFolder);

  // 4) Append a row to "Extraction Log" in the output folder
  appendExtractionLog(file, fileContent, extracted, textFileLink, outputFolder, spreadsheet);
}

/**
 * Extract text from a file:
 * - PDF → Google Doc OCR (Hebrew).
 * - Google Doc → read with DocumentApp.
 * - Other (e.g. txt) → read blob text.
 * Returns the string content or '' if something fails.
 */
function getFileContent(file) {
  var fileName = file.getName();
  var fileId = file.getId();
  var mimeType = file.getMimeType();
  var fileContent = '';

  // 1) PDF with Hebrew OCR
  if (mimeType === 'application/pdf' || fileName.toLowerCase().endsWith('.pdf')) {
    try {
      var resource = {
        mimeType: 'application/vnd.google-apps.document'
      };
      var newDoc = Drive.Files.copy(
        resource,
        fileId,
        { convert: true, ocr: true, ocrLanguage: 'he' }
      );
      var docId = newDoc.id;
      var doc = DocumentApp.openById(docId);
      fileContent = doc.getBody().getText();

      // Remove the temporary Google Doc
      try {
        Drive.Files.remove(docId);
        Logger.log("Removed temporary Google Doc: " + docId);
      } catch (remErr) {
        Logger.log("Error removing temporary Google Doc: " + remErr);
      }
    } catch (err) {
      Logger.log("Error converting PDF to Google Doc: " + err);
      throw err;
    }
  }

  // 2) Google Docs (vnd.google-apps.document)
  else if (mimeType.startsWith('application/vnd.google-apps')) {
    if (mimeType === 'application/vnd.google-apps.document') {
      try {
        var googleDoc = DocumentApp.openById(fileId);
        fileContent = googleDoc.getBody().getText();

        // Remove the original Google Doc once read
        try {
          Drive.Files.remove(fileId);
          Logger.log("Removed original Google Doc: " + fileId);
        } catch (remErr) {
          Logger.log("Error removing Google Doc: " + remErr);
        }
      } catch (docErr) {
        Logger.log("Error reading Google Doc: " + docErr);
        return '';
      }
    } else {
      Logger.log("Skipping " + fileName + " (unhandled G Suite type: " + mimeType + ")");
      return '';
    }
  }

  // 3) All others (e.g., .txt, .csv)
  else {
    try {
      var fileBlob = file.getBlob();
      fileContent = fileBlob.getDataAsString();
    } catch (blobErr) {
      Logger.log("Error reading blob from " + fileName + ": " + blobErr);
      return '';
    }
  }

  Logger.log("Extracted content from " + fileName + ": " + fileContent);
  return fileContent;
}

/**
 * Create a .txt file with the given content in the specified folder.
 * Return the shareable link to the new text file.
 */
function createTextFile(fileName, fileContent, outputFolder) {
  if (!fileContent) return '';

  var textBlob = Utilities.newBlob(fileContent, 'text/plain', fileName + '.txt');
  var newTextFile = outputFolder.createFile(textBlob);
  Logger.log("Created .txt file for " + fileName + " in folder: " + outputFolder.getName());

  var textFileLink = 'https://drive.google.com/file/d/' + newTextFile.getId() + '/view?usp=sharing';
  Logger.log("Link to .txt file: " + textFileLink);

  return textFileLink;
}

/**
 * Look for (or create) an "Extraction Log" spreadsheet in `outputFolder`,
 * then append a row with the original PDF link, text file link, file content,
 * and extracted fields: sum, num, and date.
 */
function appendExtractionLog(file, fileContent, extracted, textFileLink, outputFolder, spreadsheet) {
  var fileId = file.getId();
  var originalLink = 'https://drive.google.com/file/d/' + fileId + '/view?usp=sharing';

  var sheet = spreadsheet.getActiveSheet();

  Logger.log("Writing to spreadsheet: " + spreadsheet.getName() + " (" + spreadsheet.getId() + ")");
  Logger.log("Spreadsheet link: https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/edit?usp=sharing");

  var rowData = [
    file.getName(),
    originalLink,
    textFileLink,
    fileContent,
    extracted["sum"],
    extracted["num"],
    extracted["d"]
  ];
  try {
    sheet.appendRow(rowData);
    Logger.log("Appended new row: " + JSON.stringify(rowData));
  } catch (e) {
    Logger.log("Error appending row: " + e.toString());
  }
}

/**
 * Finds or creates a Google Sheet named "Extraction Log" in `outputFolder`.
 * Returns the Spreadsheet object.
 */
function findOrCreateSheetInFolder(outputFolder) {
  var sheetName = "Extraction Log";
  var existingSheets = outputFolder.getFilesByType(MimeType.GOOGLE_SHEETS);

  while (existingSheets.hasNext()) {
    var sheetFile = existingSheets.next();
    if (sheetFile.getName() === sheetName) {
      var spreadsheet = SpreadsheetApp.openById(sheetFile.getId());
      Logger.log("Found existing spreadsheet: " + sheetFile.getName() + " (" + sheetFile.getId() + ")");
      return spreadsheet;
    }
  }

  // Not found: create a new Spreadsheet in the output folder
  var newSS = SpreadsheetApp.create(sheetName);
  var newSSFile = DriveApp.getFileById(newSS.getId());
  outputFolder.addFile(newSSFile);
  DriveApp.removeFile(newSSFile); // remove from root

  // Optionally add a header row
  var sheet = newSS.getActiveSheet();
  sheet.appendRow(["File Name", "Original PDF link", "Text file link", "Extracted text", "Sum", "Num", "Date"]);

  Logger.log("Created new spreadsheet: " + newSS.getName() + " (" + newSS.getId() + ")");
  return newSS;
}

/**
 * Return the first named group "value" from a regex match, or null if none found.
 */
function matchit(name, pattern, content) {
  var matches = pattern.exec(content);
  if (matches && matches.groups && matches.groups.value) {
    Logger.log("Found matches for " + name + ": " + matches.groups.value);
    return matches.groups.value;
  }
  Logger.log("No match found for " + name);
  return null;

}
