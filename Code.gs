function handleEdit(e) {
  const sheet = e.source.getActiveSheet();
  
  if (sheet.getName() !== CONFIG.MASTER_SHEET_NAME) return;
  
  const range = e.range;
  const row = range.getRow();
  
  // Don't process header row
  if (row < 2) return;
  
  // Ensure checkboxes exist in columns N and O for edited row
  ensureCheckboxesInRow(sheet, row);
  
  // Handle status column formatting and sorting
  if (range.getColumn() === CONFIG.COLUMN_MAP.STATUS) {
    const newStatus = range.getValue();
    const rowRange = sheet.getRange(row, 1, 1, 16);
    
    if (newStatus === 'New Asset') {
      // Set entire row to blue
      rowRange.setFontColor('#0062e2');
    } else if (newStatus === 'Requires Attn - HX') {
      // Set entire row to pink
      rowRange.setFontColor('#FF2093');
      // Show comment dialog for HX team notification
      showHXCommentDialog(sheet, row);
    } else if (newStatus && newStatus !== 'New Asset' && newStatus !== 'Requires Attn - HX') {
      // Reset to black text
      rowRange.setFontColor('#000000');
    }
    
    // Always sort after status change
    ensureNewAssetsAtTop(sheet);
  }
  
  // Also check if any other column was edited - if the status is already "New Asset", ensure formatting
  if (range.getColumn() !== CONFIG.COLUMN_MAP.STATUS && range.getColumn() !== CONFIG.COLUMN_MAP.EDIT) {
    const statusCell = sheet.getRange(row, CONFIG.COLUMN_MAP.STATUS);
    const currentStatus = statusCell.getValue();
    
    if (currentStatus === 'New Asset') {
      const rowRange = sheet.getRange(row, 1, 1, 16);
      rowRange.setFontColor('#0062e2');
    } else if (currentStatus === 'Requires Attn - HX') {
      const rowRange = sheet.getRange(row, 1, 1, 16);
      rowRange.setFontColor('#FF2093');
    }
  }
  
  // Log any edit to the row (unless it's column P being edited)
  if (range.getColumn() !== CONFIG.COLUMN_MAP.EDIT) {
    logRowEdit(sheet, row);
  }
}

function ensureCheckboxesInRow(sheet, row) {
  try {
    const doubleSidedCell = sheet.getRange(row, CONFIG.COLUMN_MAP.DOUBLE_SIDED);
    const diecutCell = sheet.getRange(row, CONFIG.COLUMN_MAP.DIECUT);
    
    // Check if checkbox exists in Double Sided column (N)
    if (doubleSidedCell.getDataValidation() === null) {
      doubleSidedCell.insertCheckboxes();
    }
    
    // Check if checkbox exists in Diecut column (O)
    if (diecutCell.getDataValidation() === null) {
      diecutCell.insertCheckboxes();
    }
  } catch (error) {
    console.error('Error ensuring checkboxes:', error);
  }
}/**
 * @OnlyCurrentDoc
 * Asset Management System with Google Drive Integration and Two-Way Sync
 */

// Configuration
const CONFIG = {
  LOG_SHEET_NAME: 'AssetLog',
  LOG_ID_PREFIX: 'ASSET',
  MASTER_SHEET_NAME: 'MasterBBLMW25',
  DRIVE_FOLDER_ID: '1p816AjdED6d6uQc2yeb_X8hzpGlYOmPC',
  
  // Column mapping for MasterBBLMW25 sheet
  COLUMN_MAP: {
    ID: 1,           // A
    AREA: 2,         // B
    ASSET: 3,        // C
    STATUS: 4,       // D
    DIMENSIONS: 5,   // E
    QUANTITY: 6,     // F
    ITEM: 7,         // G
    MATERIAL: 8,     // H
    DUE_DATE: 9,     // I
    STRIKE_DATE: 10, // J
    VENUE: 11,       // K
    LOCATION: 12,    // L
    ARTWORK: 13,     // M
    DOUBLE_SIDED: 14,// N
    DIECUT: 15,      // O
    EDIT: 16         // P
  },
  
  DROPDOWN_SHEETS: {
    items: 'ItemsList',
    materials: 'MaterialsList', 
    statuses: 'StatusesList',
    venues: 'VenuesList',
    areas: 'AreasList'
  },
  
  MATERIAL_ID_MAP: {
    'Adhesive Vinyl - Matte': 'A',
    'Foamcore - 1/4"': 'B',
    'Foamcore - 1/2"': 'C',
    'Gatorplast - 1/4"': 'D',
    'Gatorplast - 1/2"': 'E',
    'Cardstock - Heavy': 'F',
    'Cardstock - Regular': 'G',
    'Fabrication': 'H',
    'Fabric': 'I'
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

  openForEdit: function(rowNumber) {
    const formData = projectSheet.getRowData(rowNumber);
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
      SpreadsheetApp.getUi().alert('Error', 'Could not load row data.', SpreadsheetApp.getUi().ButtonSet.OK);
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
    
    const existingValues = this.getDropdownValues(fieldName);
    if (existingValues.includes(newValue)) {
      return { success: false, message: 'Value already exists' };
    }
    
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
      'items': CONFIG.COLUMN_MAP.ITEM,
      'materials': CONFIG.COLUMN_MAP.MATERIAL,
      'statuses': CONFIG.COLUMN_MAP.STATUS,
      'venues': CONFIG.COLUMN_MAP.VENUE,
      'areas': CONFIG.COLUMN_MAP.AREA
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
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONFIG.MASTER_SHEET_NAME);
      sheet.getRange(1, 1, 1, 16).setValues([[
        'ID', 'Area', 'Asset', 'Status', 'Dimensions', 'Quantity', 'Item', 'Material',
        'Due Date', 'Strike Date', 'Venue', 'Location', 'Artwork', 'Double Sided', 'Diecut', 'Edit'
      ]]);
      sheet.getRange(1, 1, 1, 16).setFontWeight('bold');
      
      // Apply dropdown validations to entire columns
      this.applyColumnDropdowns(sheet);
    }
    
    return sheet;
  },

  applyColumnDropdowns: function(sheet) {
    const maxRows = 1000; // Apply to first 1000 rows
    
    // Get dropdown values
    const items = dropdownManager.getDropdownValues('items');
    const materials = dropdownManager.getDropdownValues('materials');
    const statuses = dropdownManager.getDropdownValues('statuses');
    const venues = dropdownManager.getDropdownValues('venues');
    const areas = dropdownManager.getDropdownValues('areas');
    
    // Apply data validation to Item column (G) - entire column
    if (items.length > 0) {
      const itemRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(items, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.ITEM, maxRows, 1).setDataValidation(itemRule);
    }
    
    // Apply data validation to Material column (H) - entire column
    if (materials.length > 0) {
      const materialRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(materials, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.MATERIAL, maxRows, 1).setDataValidation(materialRule);
    }
    
    // Apply data validation to Status column (D) - entire column
    if (statuses.length > 0) {
      const statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(statuses, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.STATUS, maxRows, 1).setDataValidation(statusRule);
    }
    
    // Apply data validation to Venue column (K) - entire column
    if (venues.length > 0) {
      const venueRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(venues, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.VENUE, maxRows, 1).setDataValidation(venueRule);
    }
    
    // Apply data validation to Area column (B) - entire column
    if (areas.length > 0) {
      const areaRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(areas, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.AREA, maxRows, 1).setDataValidation(areaRule);
    }
  },

  getRowData: function(rowNumber) {
    try {
      const sheet = this.getActiveSheet();
      const row = parseInt(rowNumber);
      
      if (!row || isNaN(row) || row < 2) {
        return null;
      }
      
      const lastRow = sheet.getLastRow();
      if (row > lastRow) {
        return null;
      }
      
      const range = sheet.getRange(row, 1, 1, 16);
      const values = range.getValues()[0];
      
      // Parse dimensions (format: "Width" x "Height")
      let width = '';
      let height = '';
      const dimensions = values[CONFIG.COLUMN_MAP.DIMENSIONS - 1];
      if (dimensions) {
        const dimMatch = dimensions.toString().match(/^([\d.]+)"\s*x\s*([\d.]+)"$/);
        if (dimMatch) {
          width = dimMatch[1];
          height = dimMatch[2];
        }
      }
      
      // Parse dates to YYYY-MM-DD format
      const parseDateToISO = (dateValue) => {
        if (!dateValue) return '';
        
        // If it's already a Date object
        if (dateValue instanceof Date) {
          return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        
        // If it's a string in the format "Day, Month Date, Year"
        const dateStr = dateValue.toString();
        const dateMatch = dateStr.match(/\w+,\s+(\w+)\s+(\d+),\s+(\d+)/);
        if (dateMatch) {
          const month = dateMatch[1];
          const day = dateMatch[2];
          const year = dateMatch[3];
          const date = new Date(`${month} ${day}, ${year}`);
          return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        
        return '';
      };
      
      const formData = {
        originalRowNumber: row,
        item: values[CONFIG.COLUMN_MAP.ITEM - 1] || '',
        material: values[CONFIG.COLUMN_MAP.MATERIAL - 1] || '',
        asset: values[CONFIG.COLUMN_MAP.ASSET - 1] || '',
        quantity: values[CONFIG.COLUMN_MAP.QUANTITY - 1] || '',
        width: width,
        height: height,
        dieCut: (values[CONFIG.COLUMN_MAP.DIECUT - 1] === true || values[CONFIG.COLUMN_MAP.DIECUT - 1] === 'TRUE'),
        doubleSided: (values[CONFIG.COLUMN_MAP.DOUBLE_SIDED - 1] === true || values[CONFIG.COLUMN_MAP.DOUBLE_SIDED - 1] === 'TRUE'),
        status: values[CONFIG.COLUMN_MAP.STATUS - 1] || '',
        dueDate: parseDateToISO(values[CONFIG.COLUMN_MAP.DUE_DATE - 1]),
        strikeDate: parseDateToISO(values[CONFIG.COLUMN_MAP.STRIKE_DATE - 1]),
        venue: values[CONFIG.COLUMN_MAP.VENUE - 1] || '',
        area: values[CONFIG.COLUMN_MAP.AREA - 1] || '',
        location: values[CONFIG.COLUMN_MAP.LOCATION - 1] || '',
        artworkUrl: values[CONFIG.COLUMN_MAP.ARTWORK - 1] || ''
      };
      
      return formData;
    } catch (error) {
      console.error('Error getting row data:', error);
      return null;
    }
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

  getMaterialPrefix: function(material) {
    return CONFIG.MATERIAL_ID_MAP[material] || 'Z';
  },

  getNextIdForMaterial: function(material) {
    return this.getNextId(material);
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
      
      const rowData = [
        assetId,
        assetData.area || '',
        assetData.asset || '',
        'New Asset',
        dimensions,
        assetData.quantity || '',
        assetData.item || '',
        assetData.material || '',
        dueDate,
        strikeDate,
        assetData.venue || '',
        assetData.location || '',
        assetData.artworkUrl || '',
        assetData.doubleSided || false,  // Will be checkbox
        assetData.dieCut || false,       // Will be checkbox
        'Edit'
      ];
      
      const range = sheet.getRange(nextRow, 1, 1, rowData.length);
      range.setValues([rowData]);
      
      // Insert checkboxes for Double Sided and Diecut columns
      sheet.getRange(nextRow, CONFIG.COLUMN_MAP.DOUBLE_SIDED).insertCheckboxes();
      sheet.getRange(nextRow, CONFIG.COLUMN_MAP.DIECUT).insertCheckboxes();
      
      // Set entire row to blue text for "New Asset" status
      range.setFontColor('#0062e2');
      
      const logId = this.logFormData(assetData.formData || assetData, nextRow, logIdPrefix, logSheetName);
      
      if (logId) {
        const editCell = sheet.getRange(nextRow, CONFIG.COLUMN_MAP.EDIT);
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
      
      const existingId = sheet.getRange(rowNum, CONFIG.COLUMN_MAP.ID).getValue();
      const dueDate = this.formatDate(assetData.dueDate);
      const strikeDate = this.formatDate(assetData.strikeDate);
      const dimensions = `${assetData.width}" x ${assetData.height}"`;
      
      const rowData = [
        existingId,
        assetData.area || '',
        assetData.asset || '',
        assetData.status || 'New Asset',
        dimensions,
        assetData.quantity || '',
        assetData.item || '',
        assetData.material || '',
        dueDate,
        strikeDate,
        assetData.venue || '',
        assetData.location || '',
        assetData.artworkUrl || '',
        assetData.doubleSided || false,  // Will be checkbox
        assetData.dieCut || false,       // Will be checkbox
        'Edit'
      ];
      
      const range = sheet.getRange(rowNum, 1, 1, 15);
      range.setValues([rowData.slice(0, 15)]);
      
      // Ensure checkboxes exist in columns N and O
      const doubleSidedCell = sheet.getRange(rowNum, CONFIG.COLUMN_MAP.DOUBLE_SIDED);
      const diecutCell = sheet.getRange(rowNum, CONFIG.COLUMN_MAP.DIECUT);
      
      // Check if checkbox exists, if not insert it
      if (doubleSidedCell.getDataValidation() === null) {
        doubleSidedCell.insertCheckboxes();
      }
      if (diecutCell.getDataValidation() === null) {
        diecutCell.insertCheckboxes();
      }
      
      // Apply formatting based on status
      const rowRange = sheet.getRange(rowNum, 1, 1, 16);
      if (assetData.status === 'New Asset') {
        rowRange.setFontColor('#0062e2');
      } else {
        rowRange.setFontColor('#000000');
      }
      
      const logId = this.updateLogFormData(assetData.formData || assetData, rowNum, logIdPrefix, logSheetName);
      
      if (logId) {
        const editCell = sheet.getRange(rowNum, CONFIG.COLUMN_MAP.EDIT);
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
      .addItem('Edit Selected Row', 'editSelectedRow')
      .addSeparator()
      .addItem('Setup Auto-Sort', 'setupAutoSort')
      .addItem('Setup File Renaming Trigger', 'setupFileRenamingTrigger')
      .addToUi();
  
  // Ensure triggers are set up
  setupAutoSort();
}

function setupAutoSort() {
  // Remove existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEditInstallable' || 
        trigger.getHandlerFunction() === 'onChangeInstallable') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new installable edit trigger
  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  
  // Create new installable change trigger (catches pastes, imports, etc.)
  ScriptApp.newTrigger('onChangeInstallable')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();
}

function setupFileRenamingTrigger() {
  try {
    // Remove existing file renaming triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'checkAndRenameFiles') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Create time-driven trigger to check for files every 5 minutes
    ScriptApp.newTrigger('checkAndRenameFiles')
      .timeBased()
      .everyMinutes(5)
      .create();
    
    SpreadsheetApp.getUi().alert(
      'File Renaming Trigger Setup',
      'File renaming trigger has been set up successfully!\n\nFiles in the "0 - Initial Files" folder will be automatically renamed every 5 minutes.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to setup file renaming trigger: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function checkAndRenameFiles() {
  try {
    // Get the main drive folder
    const mainFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    
    // Find the "0 - Initial Files" subfolder
    const subfolders = mainFolder.getFoldersByName(CONFIG.INITIAL_FILES_FOLDER_NAME);
    
    if (!subfolders.hasNext()) {
      console.log('Initial Files folder not found');
      return;
    }
    
    const initialFilesFolder = subfolders.next();
    
    // Get the spreadsheet to lookup asset names
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
    if (!sheet) return;
    
    // Get all files in the folder
    const files = initialFilesFolder.getFiles();
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      // Extract ID from filename (everything before the first dot or end of string)
      const idMatch = fileName.match(/^([A-Z]\d+)/);
      
      if (idMatch) {
        const assetId = idMatch[1];
        
        // Find the asset in the spreadsheet
        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        
        for (let i = 1; i < values.length; i++) {
          const rowId = values[i][CONFIG.COLUMN_MAP.ID - 1];
          
          if (rowId === assetId) {
            const assetName = values[i][CONFIG.COLUMN_MAP.ASSET - 1];
            
            if (assetName) {
              // Get file extension
              const lastDotIndex = fileName.lastIndexOf('.');
              const extension = lastDotIndex > -1 ? fileName.substring(lastDotIndex) : '';
              
              // Create new filename: ID_PROJECTCODE_AssetName.ext
              const sanitizedAssetName = assetName.toString().replace(/[^a-zA-Z0-9_-]/g, '_');
              const newFileName = `${assetId}_${CONFIG.PROJECT_CODE}_${sanitizedAssetName}${extension}`;
              
              // Only rename if the name is different
              if (fileName !== newFileName) {
                file.setName(newFileName);
                console.log(`Renamed: ${fileName} -> ${newFileName}`);
              }
            }
            break;
          }
        }
      }
    }
  } catch (error) {
    console.error('Error in checkAndRenameFiles:', error);
  }
}

function onEditInstallable(e) {
  handleEdit(e);
}

function onChangeInstallable(e) {
  // On any change, ensure New Assets are sorted to top
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
  if (sheet) {
    ensureNewAssetsAtTop(sheet);
  }
}

function handleEdit(e) {
  const sheet = e.source.getActiveSheet();
  
  if (sheet.getName() !== CONFIG.MASTER_SHEET_NAME) return;
  
  const range = e.range;
  const row = range.getRow();
  
  // Don't process header row
  if (row < 2) return;
  
  // Handle status column formatting and sorting
  if (range.getColumn() === CONFIG.COLUMN_MAP.STATUS) {
    const newStatus = range.getValue();
    const rowRange = sheet.getRange(row, 1, 1, 16);
    
    if (newStatus === 'New Asset') {
      // Set entire row to blue
      rowRange.setFontColor('#0062e2');
    } else if (newStatus === 'Requires Attn - HX') {
      // Set entire row to pink
      rowRange.setFontColor('#FF2093');
      // Show comment dialog for HX team notification
      showHXCommentDialog(sheet, row);
    } else if (newStatus && newStatus !== 'New Asset' && newStatus !== 'Requires Attn - HX') {
      // Reset to black text
      rowRange.setFontColor('#000000');
    }
    
    // Always sort after status change
    ensureNewAssetsAtTop(sheet);
  }
  
  // Also check if any other column was edited - if the status is already "New Asset", ensure formatting
  if (range.getColumn() !== CONFIG.COLUMN_MAP.STATUS && range.getColumn() !== CONFIG.COLUMN_MAP.EDIT) {
    const statusCell = sheet.getRange(row, CONFIG.COLUMN_MAP.STATUS);
    const currentStatus = statusCell.getValue();
    
    if (currentStatus === 'New Asset') {
      const rowRange = sheet.getRange(row, 1, 1, 16);
      rowRange.setFontColor('#0062e2');
    } else if (currentStatus === 'Requires Attn - HX') {
      const rowRange = sheet.getRange(row, 1, 1, 16);
      rowRange.setFontColor('#FF2093');
    }
  }
  
  // Log any edit to the row (unless it's column P being edited)
  if (range.getColumn() !== CONFIG.COLUMN_MAP.EDIT) {
    logRowEdit(sheet, row);
  }
}

function showHXCommentDialog(sheet, row) {
  try {
    // Get asset details
    const rowData = sheet.getRange(row, 1, 1, 16).getValues()[0];
    const assetId = rowData[CONFIG.COLUMN_MAP.ID - 1];
    const assetName = rowData[CONFIG.COLUMN_MAP.ASSET - 1];
    const material = rowData[CONFIG.COLUMN_MAP.MATERIAL - 1];
    const item = rowData[CONFIG.COLUMN_MAP.ITEM - 1];
    
    // Create the dialog
    const htmlOutput = HtmlService.createHtmlOutputFromFile('HXCommentDialog')
        .setWidth(420)
        .setHeight(280);
    
    // Inject asset data into the dialog
    const htmlContent = htmlOutput.getContent();
    const modifiedContent = htmlContent.replace(
      '<script>',
      `<script>window.assetData = ${JSON.stringify({
        row: row,
        assetId: assetId || 'N/A',
        assetName: assetName || 'N/A',
        item: item || 'N/A',
        material: material || 'N/A'
      })};`
    );
    
    const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent)
        .setWidth(420)
        .setHeight(280);
    
    const ui = SpreadsheetApp.getUi();
    ui.showModalDialog(modifiedOutput, '⚠️ HX Attention Required');
    
  } catch (error) {
    console.error('Error showing HX comment dialog:', error);
  }
}

function sendHXNotification(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
    const row = data.row;
    
    // Send email notifications
    const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    const sheetName = sheet.getName();
    const cellReference = `${sheetName}!D${row}`;
    const directLink = `${spreadsheetUrl}#gid=${sheet.getSheetId()}&range=D${row}`;
    
    const emailSubject = `🚨 HX Attention Required - Asset ${data.assetId}`;
    const emailBody = `
      <html>
        <body style="font-family: Arial, sans-serif; color: #333;">
          <h2 style="color: #FF2093;">⚠️ Asset Requires Attention</h2>
          
          <div style="background-color: #fef2f2; padding: 15px; border-left: 4px solid #FF2093; margin: 20px 0;">
            <p><strong>Asset Name:</strong> ${data.assetName}</p>
            <p><strong>Asset ID:</strong> ${data.assetId}</p>
            <p><strong>Item:</strong> ${data.item}</p>
            <p><strong>Material:</strong> ${data.material}</p>
            <p><strong>Location:</strong> ${cellReference}</p>
          </div>
          
          <div style="background-color: #f9fafb; padding: 15px; border-radius: 6px; margin: 20px 0;">
            <p style="margin: 0;"><strong>Comment:</strong></p>
            <p style="margin: 10px 0 0 0; white-space: pre-wrap;">${data.comment}</p>
          </div>
          
          <p>This asset has been flagged as <strong style="color: #FF2093;">"Requires Attn - HX"</strong> and needs your immediate attention.</p>
          
          <p style="margin-top: 30px;">
            <a href="${directLink}" style="background-color: #FF2093; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px;">View Asset in Spreadsheet</a>
          </p>
          
          <hr style="margin-top: 30px; border: none; border-top: 1px solid #ddd;">
          <p style="font-size: 12px; color: #666;">
            This is an automated notification from the BBLMW25 Asset Management System.
          </p>
        </body>
      </html>
    `;
    
    // Send to diana@hxecute.com
    try {
      MailApp.sendEmail({
        to: 'diana@hxecute.com',
        subject: emailSubject,
        htmlBody: emailBody
      });
    } catch (emailError) {
      console.error('Error sending email to diana@hxecute.com:', emailError);
    }
    
    // Send to heather@hxecute.com
    try {
      MailApp.sendEmail({
        to: 'heather@hxecute.com',
        subject: emailSubject,
        htmlBody: emailBody
      });
    } catch (emailError) {
      console.error('Error sending email to heather@hxecute.com:', emailError);
    }
    
    return { success: true, message: 'Notifications sent successfully' };
    
  } catch (error) {
    console.error('Error sending HX notification:', error);
    return { success: false, message: error.toString() };
  }
}

function createHXAttentionComment(sheet, row) {
  // This function is deprecated - replaced by showHXCommentDialog
  // Kept for backwards compatibility
  showHXCommentDialog(sheet, row);
}

function ensureNewAssetsAtTop(sheet) {
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return; // No data to sort
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 16);
  const values = dataRange.getValues();
  
  // Separate "New Asset" rows and other rows
  const newAssetRows = [];
  const otherRows = [];
  
  values.forEach((row, index) => {
    const status = row[CONFIG.COLUMN_MAP.STATUS - 1];
    const rowNum = index + 2;
    
    if (status === 'New Asset') {
      newAssetRows.push(row);
      // Ensure blue formatting
      sheet.getRange(rowNum, 1, 1, 16).setFontColor('#0062e2');
    } else if (status === 'Requires Attn - HX') {
      otherRows.push(row);
      // Ensure pink formatting
      sheet.getRange(rowNum, 1, 1, 16).setFontColor('#FF2093');
    } else {
      otherRows.push(row);
      // Ensure black formatting for other statuses
      sheet.getRange(rowNum, 1, 1, 16).setFontColor('#000000');
    }
  });
  
  // Sort other rows by ID (Column A)
  otherRows.sort((a, b) => {
    const idA = a[0].toString();
    const idB = b[0].toString();
    return idA.localeCompare(idB);
  });
  
  // Put "New Asset" rows at the TOP, then sorted other rows
  const sortedData = [...newAssetRows, ...otherRows];
  
  // Only update if order has changed
  const currentOrder = values.map(row => row.join('|'));
  const newOrder = sortedData.map(row => row.join('|'));
  
  if (currentOrder.join('||') !== newOrder.join('||')) {
    dataRange.setValues(sortedData);
  }
}

function onEdit(e) {
  // Keep simple trigger for backward compatibility
  handleEdit(e);
}

function logRowEdit(sheet, row) {
  try {
    // Get all row data
    const rowData = sheet.getRange(row, 1, 1, 16).getValues()[0];
    
    // Columns to check for content (excluding dropdown-only columns and checkboxes)
    // We'll check: ID (A), Asset (C), Dimensions (E), Quantity (F), Due Date (I), Strike Date (J), Location (L), Artwork (M)
    const columnsToCheck = [
      CONFIG.COLUMN_MAP.ID - 1,          // A: ID
      CONFIG.COLUMN_MAP.ASSET - 1,       // C: Asset
      CONFIG.COLUMN_MAP.DIMENSIONS - 1,  // E: Dimensions
      CONFIG.COLUMN_MAP.QUANTITY - 1,    // F: Quantity
      CONFIG.COLUMN_MAP.DUE_DATE - 1,    // I: Due Date
      CONFIG.COLUMN_MAP.STRIKE_DATE - 1, // J: Strike Date
      CONFIG.COLUMN_MAP.LOCATION - 1,    // L: Location
      CONFIG.COLUMN_MAP.ARTWORK - 1      // M: Artwork
    ];
    
    // Check if row has any meaningful content (excluding dropdown columns)
    const hasContent = columnsToCheck.some(index => {
      const cell = rowData[index];
      return cell !== '' && cell !== null && cell !== undefined;
    });
    
    if (!hasContent) return;
    
    // Parse the row data into form format
    const parseDateToISO = (dateValue) => {
      if (!dateValue) return '';
      
      if (dateValue instanceof Date) {
        return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      
      const dateStr = dateValue.toString();
      const dateMatch = dateStr.match(/\w+,\s+(\w+)\s+(\d+),\s+(\d+)/);
      if (dateMatch) {
        const month = dateMatch[1];
        const day = dateMatch[2];
        const year = dateMatch[3];
        const date = new Date(`${month} ${day}, ${year}`);
        return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      
      return '';
    };
    
    // Parse dimensions
    let width = '';
    let height = '';
    const dimensions = rowData[CONFIG.COLUMN_MAP.DIMENSIONS - 1];
    if (dimensions) {
      const dimMatch = dimensions.toString().match(/^([\d.]+)"\s*x\s*([\d.]+)"$/);
      if (dimMatch) {
        width = dimMatch[1];
        height = dimMatch[2];
      }
    }
    
    const formData = {
      originalRowNumber: row,
      item: rowData[CONFIG.COLUMN_MAP.ITEM - 1] || '',
      material: rowData[CONFIG.COLUMN_MAP.MATERIAL - 1] || '',
      asset: rowData[CONFIG.COLUMN_MAP.ASSET - 1] || '',
      quantity: rowData[CONFIG.COLUMN_MAP.QUANTITY - 1] || '',
      width: width,
      height: height,
      dieCut: (rowData[CONFIG.COLUMN_MAP.DIECUT - 1] === true),
      doubleSided: (rowData[CONFIG.COLUMN_MAP.DOUBLE_SIDED - 1] === true),
      status: rowData[CONFIG.COLUMN_MAP.STATUS - 1] || '',
      dueDate: parseDateToISO(rowData[CONFIG.COLUMN_MAP.DUE_DATE - 1]),
      strikeDate: parseDateToISO(rowData[CONFIG.COLUMN_MAP.STRIKE_DATE - 1]),
      venue: rowData[CONFIG.COLUMN_MAP.VENUE - 1] || '',
      area: rowData[CONFIG.COLUMN_MAP.AREA - 1] || '',
      location: rowData[CONFIG.COLUMN_MAP.LOCATION - 1] || '',
      artworkUrl: rowData[CONFIG.COLUMN_MAP.ARTWORK - 1] || ''
    };
    
    // Log the form data
    const logId = projectSheet.updateLogFormData(formData, row, CONFIG.LOG_ID_PREFIX, CONFIG.LOG_SHEET_NAME);
    
    // Set "Edit" in column P with note
    const editCell = sheet.getRange(row, CONFIG.COLUMN_MAP.EDIT);
    editCell.setValue('Edit');
    
    if (logId) {
      editCell.setNote(`LogID: ${logId}\n\nTo edit this item:\n1. Select this cell\n2. Go to Assets > Edit Selected Row\n\nLast updated: ${new Date().toLocaleString()}`);
      editCell.setBackground('#e3f2fd');
      editCell.setFontColor('#1976d2');
      editCell.setFontWeight('bold');
    }
    
  } catch (error) {
    console.error('Error logging row edit:', error);
  }
}

function sortNonNewAssets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(CONFIG.MASTER_SHEET_NAME);
  
  if (!sheet) return;
  
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return;
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 16);
  const values = dataRange.getValues();
  
  const newAssetRows = [];
  const otherRows = [];
  
  values.forEach((row, index) => {
    if (row[CONFIG.COLUMN_MAP.STATUS - 1] === 'New Asset') {
      newAssetRows.push({ row: row, originalIndex: index + 2 });
    } else {
      otherRows.push({ row: row, originalIndex: index + 2 });
    }
  });
  
  otherRows.sort((a, b) => {
    const idA = a.row[0].toString();
    const idB = b.row[0].toString();
    return idA.localeCompare(idB);
  });
  
  const sortedData = [...otherRows.map(r => r.row), ...newAssetRows.map(r => r.row)];
  dataRange.setValues(sortedData);
}

function openAssetApp() {
  assetApp.showDialog();
}

function editSelectedRow() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    
    // Check if we're on the master sheet
    if (sheet.getName() !== CONFIG.MASTER_SHEET_NAME) {
      SpreadsheetApp.getUi().alert(
        'Edit Row', 
        `Please select a row in the "${CONFIG.MASTER_SHEET_NAME}" sheet to edit.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const activeCell = sheet.getActiveCell();
    const rowNumber = activeCell.getRow();
    
    // Check if it's a valid data row (not header)
    if (rowNumber < 2) {
      SpreadsheetApp.getUi().alert(
        'Edit Row', 
        'Please select a data row (not the header row) to edit.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Open the form with the row data
    assetApp.openForEdit(rowNumber);
    
  } catch (error) {
    console.error('Error in editSelectedRow:', error);
    SpreadsheetApp.getUi().alert(
      'Error', 
      'An error occurred while trying to edit the row: ' + error.message,
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

function getMaterialPrefix(material) {
  return projectSheet.getMaterialPrefix(material);
}

function getNextIdForMaterial(material) {
  return projectSheet.getNextIdForMaterial(material);
}

function getRowDataForEdit(rowNumber) {
  return projectSheet.getRowData(rowNumber);
}
