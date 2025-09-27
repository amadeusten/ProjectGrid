/**
 * @OnlyCurrentDoc
 * Asset Management System with Google Drive Integration
 */

// Configuration
const CONFIG = {
  LOG_SHEET_NAME: 'AssetLog',
  LOG_ID_PREFIX: 'ASSET',
  MASTER_SHEET_NAME: 'MasterBBLMW25', // Main data sheet name
  DRIVE_FOLDER_ID: '1p816AjdED6d6uQc2yeb_X8hzpGlYOmPC', // Replace with actual folder ID
  
  // Dropdown data sheets
  DROPDOWN_SHEETS: {
    items: 'ItemsList',
    materials: 'MaterialsList', 
    statuses: 'StatusesList',
    venues: 'VenuesList',
    areas: 'AreasList'
  },
  
  // Material ID prefixes
  MATERIAL_ID_MAP: {
    'Gatorplast 3/16"': 'A',
    'Gatorplast 1/2"': 'B',
    'Wall & Window Vinyl - Matte': 'C',
    'Bar Wrap Vinyl - Matte': 'D',
    'Heavy Cardstock': 'E',
    'Static Cling - White': 'F',
    'Static Cling - Clear': 'G',
    'PVC 1/8"': 'H',
    'Fabrication': 'I'
    // Add more materials and their prefixes as needed
  }
};

// Main application namespace
const assetApp = {
  showDialog: function() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('AssetForm')
        .setWidth(600)
        .setHeight(900);
    const ui = SpreadsheetApp.getUi();
    ui.showModalDialog(htmlOutput, 'BBLMW25 : Add New Asset');
  },

  openForEdit: function(logId) {
    const formData = projectSheet.getLoggedFormData(logId, CONFIG.LOG_SHEET_NAME);
    if (formData) {
      const htmlOutput = HtmlService.createHtmlOutputFromFile('AssetForm')
          .setWidth(600)
          .setHeight(900);
      
      const htmlContent = htmlOutput.getContent();
      const modifiedContent = htmlContent.replace(
        '<script>',
        `<script>window.editFormData = ${JSON.stringify(formData)};`
      );
      
      const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent)
          .setWidth(600)
          .setHeight(900);
      
      const ui = SpreadsheetApp.getUi();
      ui.showModalDialog(modifiedOutput, 'Edit Asset');
    } else {
      this.showDialog();
    }
  },

  addToProject: function(assetData) {
    try {
      return projectSheet.addProjectItem(assetData, CONFIG.LOG_ID_PREFIX, CONFIG.LOG_SHEET_NAME);
    } catch (e) {
      console.error("Error in assetApp.addToProject: " + e.toString());
      return {
        success: false,
        message: `Error adding to project: ${e.toString()}`,
        rowNumber: null,
        logId: null
      };
    }
  }
};

// Dropdown management namespace
const dropdownManager = {
  getDropdownValues: function(fieldName) {
    const sheetName = CONFIG.DROPDOWN_SHEETS[fieldName];
    if (!sheetName) return [];
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.getRange(1, 1).setValue(fieldName.charAt(0).toUpperCase() + fieldName.slice(1));
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    return values.filter(row => row[0] && row[0].toString().trim() !== "")
                 .map(row => row[0].toString().trim());
  },

  addDropdownValue: function(fieldName, newValue) {
    const sheetName = CONFIG.DROPDOWN_SHEETS[fieldName];
    if (!sheetName) return { success: false, message: 'Invalid field name' };
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.getRange(1, 1).setValue(fieldName.charAt(0).toUpperCase() + fieldName.slice(1));
    }
    
    // Check if value already exists
    const existingValues = this.getDropdownValues(fieldName);
    if (existingValues.includes(newValue)) {
      return { success: false, message: 'Value already exists' };
    }
    
    // Add new value
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1).setValue(newValue);
    
    return { success: true, message: 'Value added successfully', value: newValue };
  },

  updateDropdownValue: function(fieldName, oldValue, newValue) {
    const sheetName = CONFIG.DROPDOWN_SHEETS[fieldName];
    if (!sheetName) return { success: false, message: 'Invalid field name' };
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) return { success: false, message: 'Sheet not found' };
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'No values to update' };
    
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === oldValue) {
        sheet.getRange(i + 2, 1).setValue(newValue);
        
        // Update in main sheet
        this.updateInMainSheet(fieldName, oldValue, newValue);
        
        return { success: true, message: 'Value updated successfully' };
      }
    }
    
    return { success: false, message: 'Value not found' };
  },

  deleteDropdownValue: function(fieldName, value) {
    const sheetName = CONFIG.DROPDOWN_SHEETS[fieldName];
    if (!sheetName) return { success: false, message: 'Invalid field name' };
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) return { success: false, message: 'Sheet not found' };
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'No values to delete' };
    
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === value) {
        sheet.deleteRow(i + 2);
        return { success: true, message: 'Value deleted successfully' };
      }
    }
    
    return { success: false, message: 'Value not found' };
  },

  updateInMainSheet: function(fieldName, oldValue, newValue) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = spreadsheet.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    
    if (!mainSheet) return;
    
    const columnMap = {
      'items': 7,    // Column G
      'materials': 8, // Column H
      'statuses': 4,  // Column D
      'venues': 11,   // Column K
      'areas': 2      // Column B
    };
    
    const column = columnMap[fieldName];
    if (!column) return;
    
    const lastRow = mainSheet.getLastRow();
    if (lastRow <= 1) return;
    
    const range = mainSheet.getRange(2, column, lastRow - 1, 1);
    const values = range.getValues();
    
    let updated = false;
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === oldValue) {
        values[i][0] = newValue;
        updated = true;
      }
    }
    
    if (updated) {
      range.setValues(values);
    }
  },

  getAllDropdownData: function() {
    return {
      items: this.getDropdownValues('items'),
      materials: this.getDropdownValues('materials'),
      statuses: this.getDropdownValues('statuses'),
      venues: this.getDropdownValues('venues'),
      areas: this.getDropdownValues('areas')
    };
  }
};

// Drive management namespace
const driveManager = {
  getFolderStructure: function() {
    try {
      const primaryFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
      const folders = [];
      
      folders.push({
        id: primaryFolder.getId(),
        name: primaryFolder.getName()
      });
      
      const subfolders = primaryFolder.getFolders();
      while (subfolders.hasNext()) {
        const folder = subfolders.next();
        folders.push({
          id: folder.getId(),
          name: folder.getName()
        });
      }
      
      return folders;
    } catch (e) {
      console.error('Error getting folder structure: ' + e.toString());
      return [];
    }
  },

  uploadFile: function(fileData, folderId, fileName) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      const blob = Utilities.newBlob(
        Utilities.base64Decode(fileData.split(',')[1]),
        fileData.split(';')[0].split(':')[1],
        fileName
      );
      
      const file = folder.createFile(blob);
      return {
        success: true,
        fileId: file.getId(),
        fileUrl: file.getUrl(),
        fileName: file.getName()
      };
    } catch (e) {
      console.error('Error uploading file: ' + e.toString());
      return {
        success: false,
        message: 'Error uploading file: ' + e.toString()
      };
    }
  }
};

// Project sheet management
const projectSheet = {
  getActiveSheet: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    
    // If MasterBBLMW25 doesn't exist, create it
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONFIG.MASTER_SHEET_NAME);
      // Set up headers
      sheet.getRange(1, 1, 1, 16).setValues([[
        'ID', 'Area', 'Asset', 'Status', 'Dimensions', 'Quantity', 'Item', 'Material',
        'Due Date', 'Strike Date', 'Venue', 'Location', 'Artwork', 'Double Sided', 'Diecut', 'Edit'
      ]]);
      sheet.getRange(1, 1, 1, 16).setFontWeight('bold');
    }
    
    return sheet;
  },

  getNextId: function(material) {
    const sheet = this.getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    const prefix = CONFIG.MATERIAL_ID_MAP[material] || 'Z';
    let maxNumber = 0;
    
    for (let i = 1; i < values.length; i++) {
      const cellValue = values[i][0];
      if (cellValue && typeof cellValue === 'string' && cellValue.startsWith(prefix)) {
        const numberPart = cellValue.substring(prefix.length);
        const number = parseInt(numberPart);
        if (!isNaN(number) && number > maxNumber) {
          maxNumber = number;
        }
      }
    }
    
    const nextNumber = maxNumber + 1;
    return `${prefix}${nextNumber}`;
  },

  formatDate: function(dateString) {
    if (!dateString) return '';
    const date = new Date(dateString);
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                    'July', 'August', 'September', 'October', 'November', 'December'];
    
    return `${days[date.getDay()]}, ${months[date.getMonth()]} ${date.getDate()}, ${date.getFullYear()}`;
  },

  logFormData: function(formData, projectRowNumber, logIdPrefix, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = spreadsheet.getSheetByName(logSheetName);
      
      if (!logSheet) {
        logSheet = spreadsheet.insertSheet(logSheetName);
        logSheet.hideSheet();
        logSheet.getRange(1, 1, 1, 4).setValues([
          ['LogID', 'ProjectRow', 'Timestamp', 'FormData']
        ]);
      }
      
      const logId = `${logIdPrefix}_${Date.now()}_${projectRowNumber}`;
      const timestamp = new Date();
      const formDataWithRow = { ...formData, originalRowNumber: projectRowNumber };
      const formDataJson = JSON.stringify(formDataWithRow);
      
      const lastLogRow = logSheet.getLastRow();
      logSheet.getRange(lastLogRow + 1, 1, 1, 4).setValues([
        [logId, projectRowNumber, timestamp, formDataJson]
      ]);
      
      return logId;
    } catch (error) {
      console.error('Error logging form data:', error);
      return null;
    }
  },

  updateLogFormData: function(formData, projectRowNumber, logIdPrefix, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = spreadsheet.getSheetByName(logSheetName);
      
      if (!logSheet) {
        return this.logFormData(formData, projectRowNumber, logIdPrefix, logSheetName);
      }
      
      const dataRange = logSheet.getDataRange();
      const values = dataRange.getValues();
      let existingRowIndex = -1;
      
      for (let i = 1; i < values.length; i++) {
        if (values[i][1] === projectRowNumber) {
          existingRowIndex = i + 1;
          break;
        }
      }
      
      const logId = `${logIdPrefix}_${Date.now()}_${projectRowNumber}`;
      const timestamp = new Date();
      const formDataWithRow = { ...formData, originalRowNumber: projectRowNumber };
      const formDataJson = JSON.stringify(formDataWithRow);
      
      if (existingRowIndex > 0) {
        logSheet.getRange(existingRowIndex, 1, 1, 4).setValues([
          [logId, projectRowNumber, timestamp, formDataJson]
        ]);
      } else {
        const lastLogRow = logSheet.getLastRow();
        logSheet.getRange(lastLogRow + 1, 1, 1, 4).setValues([
          [logId, projectRowNumber, timestamp, formDataJson]
        ]);
      }
      
      return logId;
    } catch (error) {
      console.error('Error updating log data:', error);
      return null;
    }
  },

  getLoggedFormData: function(logId, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = spreadsheet.getSheetByName(logSheetName);
      
      if (!logSheet) return null;
      
      const dataRange = logSheet.getDataRange();
      const values = dataRange.getValues();
      
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === logId) {
          return JSON.parse(values[i][3]);
        }
      }
      
      return null;
    } catch (error) {
      console.error('Error retrieving form data:', error);
      return null;
    }
  },

  addProjectItem: function(assetData, logIdPrefix, logSheetName) {
    try {
      const sheet = this.getActiveSheet();
      const originalRowNumber = assetData.originalRowNumber || (assetData.formData && assetData.formData.originalRowNumber);
      
      if (originalRowNumber && originalRowNumber > 0) {
        return this.updateProjectItem(assetData, logIdPrefix, logSheetName);
      }
      
      const lastRow = sheet.getLastRow();
      const nextRow = lastRow + 1;
      
      const assetId = this.getNextId(assetData.material);
      const dueDate = this.formatDate(assetData.dueDate);
      const strikeDate = this.formatDate(assetData.strikeDate);
      const dimensions = `${assetData.width}" x ${assetData.height}"`;
      
      // Column mapping per blueprint
      const rowData = [
        assetId,                    // A: ID
        assetData.area || '',       // B: Area
        assetData.asset || '',      // C: Asset
        'New Asset',                // D: Status (default)
        dimensions,                 // E: Dimensions (Width x Height)
        assetData.quantity || '',   // F: Quantity
        assetData.item || '',       // G: Item
        assetData.material || '',   // H: Material
        dueDate,                    // I: Due Date
        strikeDate,                 // J: Strike Date
        assetData.venue || '',      // K: Venue
        assetData.location || '',   // L: Location
        assetData.artworkUrl || '', // M: Artwork (file URL)
        assetData.doubleSided ? 'TRUE' : 'FALSE', // N: Double Sided
        assetData.dieCut ? 'TRUE' : 'FALSE',      // O: Diecut
        'Edit'                      // P: Edit instruction
      ];
      
      const range = sheet.getRange(nextRow, 1, 1, rowData.length);
      range.setValues([rowData]);
      
      // Set "New Asset" formatting (blue text in Status column D)
      sheet.getRange(nextRow, 4).setFontColor('#0062e2');
      
      // Log form data
      const logId = this.logFormData(assetData.formData || assetData, nextRow, logIdPrefix, logSheetName);
      
      if (logId) {
        const editCell = sheet.getRange(nextRow, 16);
        editCell.setNote(`LogID: ${logId}\n\nTo edit this item:\n1. Select this cell\n2. Go to Assets > Edit Selected Item`);
        editCell.setBackground('#e3f2fd');
        editCell.setFontColor('#1976d2');
        editCell.setFontWeight('bold');
      }
      
      return {
        success: true,
        message: `Asset added to row ${nextRow}`,
        rowNumber: nextRow,
        logId: logId,
        isUpdate: false
      };
      
    } catch (error) {
      console.error('Error adding project item:', error);
      return {
        success: false,
        message: `Error adding item: ${error.message}`,
        rowNumber: null,
        logId: null,
        isUpdate: false
      };
    }
  },

  updateProjectItem: function(assetData, logIdPrefix, logSheetName) {
    try {
      const sheet = this.getActiveSheet();
      const rowNum = parseInt(assetData.originalRowNumber);
      
      if (!rowNum || isNaN(rowNum) || rowNum < 1) {
        throw new Error(`Invalid row number: ${assetData.originalRowNumber}`);
      }
      
      const existingId = sheet.getRange(rowNum, 1).getValue();
      const dueDate = this.formatDate(assetData.dueDate);
      const strikeDate = this.formatDate(assetData.strikeDate);
      const dimensions = `${assetData.width}" x ${assetData.height}"`;
      
      const rowData = [
        existingId,                 // A: ID (keep existing)
        assetData.area || '',       // B: Area
        assetData.asset || '',      // C: Asset
        assetData.status || 'New Asset', // D: Status
        dimensions,                 // E: Dimensions
        assetData.quantity || '',   // F: Quantity
        assetData.item || '',       // G: Item
        assetData.material || '',   // H: Material
        dueDate,                    // I: Due Date
        strikeDate,                 // J: Strike Date
        assetData.venue || '',      // K: Venue
        assetData.location || '',   // L: Location
        assetData.artworkUrl || '', // M: Artwork
        assetData.doubleSided ? 'TRUE' : 'FALSE', // N: Double Sided
        assetData.dieCut ? 'TRUE' : 'FALSE',      // O: Diecut
        'Edit'                      // P: Edit
      ];
      
      const range = sheet.getRange(rowNum, 1, 1, 15);
      range.setValues([rowData.slice(0, 15)]);
      
      const logId = this.updateLogFormData(assetData.formData || assetData, rowNum, logIdPrefix, logSheetName);
      
      if (logId) {
        const editCell = sheet.getRange(rowNum, 16);
        editCell.setNote(`LogID: ${logId}\n\nTo edit this item:\n1. Select this cell\n2. Go to Assets > Edit Selected Item\n\nLast updated: ${new Date().toLocaleString()}`);
      }
      
      return {
        success: true,
        message: `Asset updated in row ${rowNum}`,
        rowNumber: rowNum,
        logId: logId,
        isUpdate: true
      };
      
    } catch (error) {
      console.error('Error updating project item:', error);
      return {
        success: false,
        message: `Error updating item: ${error.message}`,
        rowNumber: null,
        logId: null,
        isUpdate: false
      };
    }
  }
};

// Menu and trigger functions
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Assets')
      .addItem('Add New Asset', 'openAssetApp')
      .addSeparator()
      .addItem('Edit Selected Item', 'editSelectedItem')
      .addToUi();
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  
  // Only run on MasterBBLMW25 sheet
  if (sheet.getName() !== CONFIG.MASTER_SHEET_NAME) return;
  
  const range = e.range;
  
  // Check if Status column (D) was edited
  if (range.getColumn() === 4 && range.getRow() > 1) {
    const newStatus = range.getValue();
    
    if (newStatus && newStatus !== 'New Asset') {
      // Reset text color to black
      range.setFontColor('#000000');
      
      // Sort by ID (Column A), but only non-"New Asset" rows
      sortNonNewAssets();
    }
  }
}

function sortNonNewAssets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(CONFIG.MASTER_SHEET_NAME);
  
  if (!sheet) return;
  
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return; // No data to sort
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 16);
  const values = dataRange.getValues();
  
  // Separate "New Asset" rows and other rows
  const newAssetRows = [];
  const otherRows = [];
  
  values.forEach((row, index) => {
    if (row[3] === 'New Asset') { // Column D (Status)
      newAssetRows.push({ row: row, originalIndex: index + 2 });
    } else {
      otherRows.push({ row: row, originalIndex: index + 2 });
    }
  });
  
  // Sort other rows by ID (Column A)
  otherRows.sort((a, b) => {
    const idA = a.row[0].toString();
    const idB = b.row[0].toString();
    return idA.localeCompare(idB);
  });
  
  // Write back sorted data
  const sortedData = [...otherRows.map(r => r.row), ...newAssetRows.map(r => r.row)];
  dataRange.setValues(sortedData);
}

function openAssetApp() {
  assetApp.showDialog();
}

function editSelectedItem() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    const activeCell = sheet.getActiveCell();
    
    const column = activeCell.getColumn();
    if (column !== 16) { // Column P
      SpreadsheetApp.getUi().alert(
        'Edit Item', 
        'Please select an "Edit" cell first, then try again.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const cellNote = activeCell.getNote();
    if (!cellNote || !cellNote.includes('LogID:')) {
      SpreadsheetApp.getUi().alert(
        'Edit Item', 
        'No edit data found for this item.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const logIdMatch = cellNote.match(/LogID:\s*([^\n\r]+)/);
    if (!logIdMatch) {
      SpreadsheetApp.getUi().alert(
        'Edit Item', 
        'Could not find LogID in the selected cell.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const logId = logIdMatch[1].trim();
    assetApp.openForEdit(logId);
    
  } catch (error) {
    console.error('Error in editSelectedItem:', error);
    SpreadsheetApp.getUi().alert(
      'Error', 
      'An error occurred while trying to edit the item: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// Exposed functions for client-side calls
function getDropdownData() {
  return dropdownManager.getAllDropdownData();
}

function addNewDropdownValue(fieldName, value) {
  return dropdownManager.addDropdownValue(fieldName, value);
}

function updateDropdownValue(fieldName, oldValue, newValue) {
  return dropdownManager.updateDropdownValue(fieldName, oldValue, newValue);
}

function deleteDropdownValue(fieldName, value) {
  return dropdownManager.deleteDropdownValue(fieldName, value);
}

function getDriveFolders() {
  return driveManager.getFolderStructure();
}

function uploadFileToDrive(fileData, folderId, fileName) {
  return driveManager.uploadFile(fileData, folderId, fileName);
}

function addAssetToProject(assetData) {
  return assetApp.addToProject(assetData);
}
