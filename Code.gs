/**
 * @OnlyCurrentDoc
 * Asset Management System with Google Drive Integration and Two-Way Sync
 */

// Configuration
const CONFIG = {
  LOG_SHEET_NAME: 'AssetLog',
  LOG_ID_PREFIX: 'ASSET',
  MASTER_SHEET_NAME: 'MasterBBLMW25',
  MATERIAL_ID_MAP_SHEET: 'MaterialIDMap',
  DRIVE_FOLDER_ID: '1p816AjdED6d6uQc2yeb_X8hzpGlYOmPC',
  INITIAL_FILES_FOLDER_NAME: '0 - Initial Files',
  THUMBNAIL_FOLDER_NAME: '1 - Thumbnail Images',
  PROJECT_CODE: '19126BB',
  
  COLUMN_MAP: {
    ID: 1, AREA: 2, ASSET: 3, STATUS: 4, DIMENSIONS: 5, QUANTITY: 6,
    ITEM: 7, MATERIAL: 8, DUE_DATE: 9, STRIKE_DATE: 10, VENUE: 11,
    LOCATION: 12, ARTWORK: 13, IMAGE_LINK: 14, DOUBLE_SIDED: 15, DIECUT: 16, 
    PRODUCTION_STATUS: 17, EDIT: 18
  },
  
  DROPDOWN_SHEETS: {
    items: 'ItemsList',
    materials: 'MaterialsList', 
    statuses: 'StatusesList',
    venues: 'VenuesList',
    areas: 'AreasList',
    productionStatuses: 'ProductionStatusesList'
  },
  
  INITIAL_MATERIAL_ID_MAP: {
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

const assetApp = {
  showDialog: function() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('AssetForm')
        .setWidth(600).setHeight(900);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'BBLMW25 : Add New Asset');
  },

  openForEdit: function(rowNumber) {
    const formData = projectSheet.getRowData(rowNumber);
    if (formData) {
      const htmlOutput = HtmlService.createHtmlOutputFromFile('AssetForm')
          .setWidth(600).setHeight(900);
      const htmlContent = htmlOutput.getContent();
      const modifiedContent = htmlContent.replace('<script>', `<script>window.editFormData = ${JSON.stringify(formData)};`);
      const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent).setWidth(600).setHeight(900);
      SpreadsheetApp.getUi().showModalDialog(modifiedOutput, 'Edit Asset');
    } else {
      SpreadsheetApp.getUi().alert('Error', 'Could not load row data.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  },

  addToProject: function(assetData) {
    try {
      return projectSheet.addProjectItem(assetData, CONFIG.LOG_ID_PREFIX, CONFIG.LOG_SHEET_NAME);
    } catch (e) {
      console.error("Error in assetApp.addToProject: " + e.toString());
      return { success: false, message: `Error adding to project: ${e.toString()}`, rowNumber: null, logId: null };
    }
  }
};

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
    return values.filter(row => row[0] && row[0].toString().trim() !== "").map(row => row[0].toString().trim());
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
    if (existingValues.includes(newValue)) return { success: false, message: 'Value already exists' };
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1).setValue(newValue);
    if (fieldName === 'materials') materialIDManager.assignIDToMaterial(newValue);
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
        if (fieldName === 'materials') materialIDManager.updateMaterialName(oldValue, newValue);
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
        if (fieldName === 'materials') materialIDManager.deleteMaterial(value);
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
      'items': CONFIG.COLUMN_MAP.ITEM, 'materials': CONFIG.COLUMN_MAP.MATERIAL,
      'statuses': CONFIG.COLUMN_MAP.STATUS, 'venues': CONFIG.COLUMN_MAP.VENUE, 
      'areas': CONFIG.COLUMN_MAP.AREA, 'productionStatuses': CONFIG.COLUMN_MAP.PRODUCTION_STATUS
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
    if (updated) range.setValues(values);
  },

  getAllDropdownData: function() {
    return {
      items: this.getDropdownValues('items'), materials: this.getDropdownValues('materials'),
      statuses: this.getDropdownValues('statuses'), venues: this.getDropdownValues('venues'), 
      areas: this.getDropdownValues('areas'), productionStatuses: this.getDropdownValues('productionStatuses')
    };
  }
};

const materialIDManager = {
  getOrCreateMaterialIDSheet: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CONFIG.MATERIAL_ID_MAP_SHEET);
    
    if (!sheet) {
      // Create new sheet
      sheet = spreadsheet.insertSheet(CONFIG.MATERIAL_ID_MAP_SHEET);
      sheet.getRange(1, 1, 1, 2).setValues([['Material', 'ID Prefix']]);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
      
      // Add initial data
      const initialData = [];
      for (const [material, prefix] of Object.entries(CONFIG.INITIAL_MATERIAL_ID_MAP)) {
        initialData.push([material, prefix]);
      }
      if (initialData.length > 0) {
        sheet.getRange(2, 1, initialData.length, 2).setValues(initialData);
      }
      
      // Hide the sheet
      sheet.hideSheet();
    } else {
      // Ensure existing sheet is hidden
      if (!sheet.isSheetHidden()) {
        sheet.hideSheet();
      }
    }
    
    return sheet;
  },

  getMaterialIDMap: function() {
    const sheet = this.getOrCreateMaterialIDSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return {};
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const map = {};
    values.forEach(row => {
      if (row[0] && row[1]) map[row[0].toString().trim()] = row[1].toString().trim();
    });
    return map;
  },

  getNextAvailableLetter: function() {
    const sheet = this.getOrCreateMaterialIDSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 'A';
    const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const usedLetters = values.map(row => row[0]).filter(letter => letter);
    let maxCharCode = 64;
    usedLetters.forEach(letter => {
      const charCode = letter.charCodeAt(0);
      if (charCode > maxCharCode) maxCharCode = charCode;
    });
    return String.fromCharCode(maxCharCode + 1);
  },

  assignIDToMaterial: function(materialName) {
    const sheet = this.getOrCreateMaterialIDSheet();
    const map = this.getMaterialIDMap();
    if (map[materialName]) return map[materialName];
    const nextLetter = this.getNextAvailableLetter();
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[materialName, nextLetter]]);
    return nextLetter;
  },

  getMaterialPrefix: function(material) {
    const map = this.getMaterialIDMap();
    return map[material] || 'Z';
  },

  updateMaterialName: function(oldName, newName) {
    const sheet = this.getOrCreateMaterialIDSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === oldName) {
        sheet.getRange(i + 2, 1).setValue(newName);
        return;
      }
    }
  },

  deleteMaterial: function(materialName) {
    const sheet = this.getOrCreateMaterialIDSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === materialName) {
        sheet.deleteRow(i + 2);
        return;
      }
    }
  }
};

const driveManager = {
  getFolderStructure: function() {
    try {
      const primaryFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
      const folders = [{ id: primaryFolder.getId(), name: primaryFolder.getName() }];
      const subfolders = primaryFolder.getFolders();
      while (subfolders.hasNext()) {
        const folder = subfolders.next();
        folders.push({ id: folder.getId(), name: folder.getName() });
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
      const blob = Utilities.newBlob(Utilities.base64Decode(fileData.split(',')[1]), fileData.split(';')[0].split(':')[1], fileName);
      const file = folder.createFile(blob);
      return { success: true, fileId: file.getId(), fileUrl: file.getUrl(), fileName: file.getName() };
    } catch (e) {
      console.error('Error uploading file: ' + e.toString());
      return { success: false, message: 'Error uploading file: ' + e.toString() };
    }
  },

  updateImageLinks: function() {
    try {
      const primaryFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
      const thumbnailFolders = primaryFolder.getFoldersByName(CONFIG.THUMBNAIL_FOLDER_NAME);
      
      if (!thumbnailFolders.hasNext()) {
        return { success: false, message: `Thumbnail folder "${CONFIG.THUMBNAIL_FOLDER_NAME}" not found.` };
      }
      
      const thumbnailFolder = thumbnailFolders.next();
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
      
      if (!sheet) {
        return { success: false, message: 'Master sheet not found.' };
      }
      
      const files = thumbnailFolder.getFiles();
      let updateCount = 0;
      
      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();
        
        // Extract asset ID from filename (e.g., A01.jpg -> A01)
        const fileBaseName = fileName.substring(0, fileName.lastIndexOf('.')) || fileName;
        const assetId = fileBaseName;
        
        // Find matching row in spreadsheet
        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        
        for (let i = 1; i < values.length; i++) {
          const rowId = values[i][CONFIG.COLUMN_MAP.ID - 1];
          if (rowId === assetId) {
            // Update the image link column
            const imageLink = file.getUrl();
            sheet.getRange(i + 1, CONFIG.COLUMN_MAP.IMAGE_LINK).setValue(imageLink);
            updateCount++;
            break;
          }
        }
      }
      
      return { success: true, message: `Updated ${updateCount} image links.`, updateCount: updateCount };
    } catch (e) {
      console.error('Error updating image links: ' + e.toString());
      return { success: false, message: 'Error updating image links: ' + e.toString() };
    }
  }
};

const projectSheet = {
  getActiveSheet: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONFIG.MASTER_SHEET_NAME);
      const headers = ['ID', 'Area', 'Asset', 'Status', 'Dimensions', 'Quantity', 'Item', 'Material', 'Due Date', 'Strike Date', 'Venue', 'Location', 'Artwork', 'Image Link', 'Double Sided', 'Diecut', 'Production Status', 'Edit'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      this.applyColumnDropdowns(sheet);
      this.setupConditionalFormatting(sheet);
    }
    return sheet;
  },

  applyColumnDropdowns: function(sheet) {
    const maxRows = 1000;
    const items = dropdownManager.getDropdownValues('items');
    const materials = dropdownManager.getDropdownValues('materials');
    const statuses = dropdownManager.getDropdownValues('statuses');
    const venues = dropdownManager.getDropdownValues('venues');
    const areas = dropdownManager.getDropdownValues('areas');
    const productionStatuses = dropdownManager.getDropdownValues('productionStatuses');
    
    if (items.length > 0) {
      const itemRule = SpreadsheetApp.newDataValidation().requireValueInList(items, true).setAllowInvalid(true).build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.ITEM, maxRows, 1).setDataValidation(itemRule);
    }
    if (materials.length > 0) {
      const materialRule = SpreadsheetApp.newDataValidation().requireValueInList(materials, true).setAllowInvalid(true).build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.MATERIAL, maxRows, 1).setDataValidation(materialRule);
    }
    if (statuses.length > 0) {
      const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(statuses, true).setAllowInvalid(true).build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.STATUS, maxRows, 1).setDataValidation(statusRule);
    }
    if (venues.length > 0) {
      const venueRule = SpreadsheetApp.newDataValidation().requireValueInList(venues, true).setAllowInvalid(true).build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.VENUE, maxRows, 1).setDataValidation(venueRule);
    }
    if (areas.length > 0) {
      const areaRule = SpreadsheetApp.newDataValidation().requireValueInList(areas, true).setAllowInvalid(true).build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.AREA, maxRows, 1).setDataValidation(areaRule);
    }
    if (productionStatuses.length > 0) {
      const productionStatusRule = SpreadsheetApp.newDataValidation().requireValueInList(productionStatuses, true).setAllowInvalid(true).build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.PRODUCTION_STATUS, maxRows, 1).setDataValidation(productionStatusRule);
    }
  },

  setupConditionalFormatting: function(sheet) {
    // Clear existing conditional formatting rules
    sheet.clearConditionalFormatRules();
    
    // Add conditional formatting for Status = "Delivered"
    const range = sheet.getRange(2, 1, 1000, CONFIG.COLUMN_MAP.EDIT); // Entire row range
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$D2="Delivered"`) // Status column is D (column 4)
      .setFontColor('#38AE74')
      .setRanges([range])
      .build();
    
    sheet.setConditionalFormatRules([rule]);
  },

  getRowData: function(rowNumber) {
    try {
      const sheet = this.getActiveSheet();
      const row = parseInt(rowNumber);
      if (!row || isNaN(row) || row < 2) return null;
      const lastRow = sheet.getLastRow();
      if (row > lastRow) return null;
      const range = sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT);
      const values = range.getValues()[0];
      let width = '', height = '';
      const dimensions = values[CONFIG.COLUMN_MAP.DIMENSIONS - 1];
      if (dimensions) {
        const dimMatch = dimensions.toString().match(/^([\d.]+)"\s*x\s*([\d.]+)"$/);
        if (dimMatch) { width = dimMatch[1]; height = dimMatch[2]; }
      }
      const parseDateToISO = (dateValue) => {
        if (!dateValue) return '';
        if (dateValue instanceof Date) return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const dateStr = dateValue.toString();
        const dateMatch = dateStr.match(/\w+,\s+(\w+)\s+(\d+),\s+(\d+)/);
        if (dateMatch) {
          const date = new Date(`${dateMatch[1]} ${dateMatch[2]}, ${dateMatch[3]}`);
          return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        return '';
      };
      return {
        originalRowNumber: row, item: values[CONFIG.COLUMN_MAP.ITEM - 1] || '', material: values[CONFIG.COLUMN_MAP.MATERIAL - 1] || '',
        asset: values[CONFIG.COLUMN_MAP.ASSET - 1] || '', quantity: values[CONFIG.COLUMN_MAP.QUANTITY - 1] || '', width: width, height: height,
        dieCut: (values[CONFIG.COLUMN_MAP.DIECUT - 1] === true), doubleSided: (values[CONFIG.COLUMN_MAP.DOUBLE_SIDED - 1] === true),
        status: values[CONFIG.COLUMN_MAP.STATUS - 1] || '', dueDate: parseDateToISO(values[CONFIG.COLUMN_MAP.DUE_DATE - 1]),
        strikeDate: parseDateToISO(values[CONFIG.COLUMN_MAP.STRIKE_DATE - 1]), venue: values[CONFIG.COLUMN_MAP.VENUE - 1] || '',
        area: values[CONFIG.COLUMN_MAP.AREA - 1] || '', location: values[CONFIG.COLUMN_MAP.LOCATION - 1] || '', 
        artworkUrl: values[CONFIG.COLUMN_MAP.ARTWORK - 1] || '', imageLink: values[CONFIG.COLUMN_MAP.IMAGE_LINK - 1] || '',
        productionStatus: values[CONFIG.COLUMN_MAP.PRODUCTION_STATUS - 1] || ''
      };
    } catch (error) {
      console.error('Error getting row data:', error);
      return null;
    }
  },

  getNextId: function(material, isEdit = false, currentId = null) {
    const sheet = this.getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const prefix = materialIDManager.getMaterialPrefix(material);
    
    // For edit mode, use the existing ID
    if (isEdit && currentId) {
      return currentId;
    }
    
    let maxNumber = 0;
    for (let i = 1; i < values.length; i++) {
      const cellValue = values[i][0];
      if (cellValue && typeof cellValue === 'string' && cellValue.startsWith(prefix)) {
        const numberPart = cellValue.substring(prefix.length);
        const number = parseInt(numberPart);
        if (!isNaN(number) && number > maxNumber) maxNumber = number;
      }
    }
    
    // Format with leading zero for values 1-9
    const nextNumber = maxNumber + 1;
    const formattedNumber = nextNumber < 10 ? `0${nextNumber}` : nextNumber.toString();
    return `${prefix}${formattedNumber}`;
  },

  getMaterialPrefix: function(material) { return materialIDManager.getMaterialPrefix(material); },
  getNextIdForMaterial: function(material) { return this.getNextId(material); },

  formatDate: function(dateString) {
    if (!dateString) return '';
    const date = new Date(dateString);
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    return `${days[date.getDay()]}, ${months[date.getMonth()]} ${date.getDate()}, ${date.getFullYear()}`;
  },

  logFormData: function(formData, projectRowNumber, logIdPrefix, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = spreadsheet.getSheetByName(logSheetName);
      if (!logSheet) {
        logSheet = spreadsheet.insertSheet(logSheetName);
        logSheet.hideSheet();
        logSheet.getRange(1, 1, 1, 4).setValues([['LogID', 'ProjectRow', 'Timestamp', 'FormData']]);
      }
      const logId = `${logIdPrefix}_${Date.now()}_${projectRowNumber}`;
      const timestamp = new Date();
      const formDataWithRow = { ...formData, originalRowNumber: projectRowNumber };
      const formDataJson = JSON.stringify(formDataWithRow);
      const lastLogRow = logSheet.getLastRow();
      logSheet.getRange(lastLogRow + 1, 1, 1, 4).setValues([[logId, projectRowNumber, timestamp, formDataJson]]);
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
      if (!logSheet) return this.logFormData(formData, projectRowNumber, logIdPrefix, logSheetName);
      const dataRange = logSheet.getDataRange();
      const values = dataRange.getValues();
      let existingRowIndex = -1;
      for (let i = 1; i < values.length; i++) {
        if (values[i][1] === projectRowNumber) { existingRowIndex = i + 1; break; }
      }
      const logId = `${logIdPrefix}_${Date.now()}_${projectRowNumber}`;
      const timestamp = new Date();
      const formDataWithRow = { ...formData, originalRowNumber: projectRowNumber };
      const formDataJson = JSON.stringify(formDataWithRow);
      if (existingRowIndex > 0) {
        logSheet.getRange(existingRowIndex, 1, 1, 4).setValues([[logId, projectRowNumber, timestamp, formDataJson]]);
      } else {
        const lastLogRow = logSheet.getLastRow();
        logSheet.getRange(lastLogRow + 1, 1, 1, 4).setValues([[logId, projectRowNumber, timestamp, formDataJson]]);
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
      if (originalRowNumber && originalRowNumber > 0) return this.updateProjectItem(assetData, logIdPrefix, logSheetName);
      
      // Insert at the bottom as before
      const lastRow = sheet.getLastRow();
      const nextRow = lastRow + 1;
      
      const assetId = this.getNextId(assetData.material);
      const dueDate = this.formatDate(assetData.dueDate);
      const strikeDate = this.formatDate(assetData.strikeDate);
      const dimensions = `${assetData.width}" x ${assetData.height}"`;
      const rowData = [
        assetId, assetData.area || '', assetData.asset || '', 'New Asset', dimensions, assetData.quantity || '', 
        assetData.item || '', assetData.material || '', dueDate, strikeDate, assetData.venue || '', 
        assetData.location || '', assetData.artworkUrl || '', '', assetData.doubleSided || false, 
        assetData.dieCut || false, '', 'Edit'
      ];
      const range = sheet.getRange(nextRow, 1, 1, rowData.length);
      range.setValues([rowData]);
      sheet.getRange(nextRow, CONFIG.COLUMN_MAP.DOUBLE_SIDED).insertCheckboxes();
      sheet.getRange(nextRow, CONFIG.COLUMN_MAP.DIECUT).insertCheckboxes();
      range.setFontColor('#0062e2');
      
      const logId = this.logFormData(assetData.formData || assetData, nextRow, logIdPrefix, logSheetName);
      if (logId) {
        const editCell = sheet.getRange(nextRow, CONFIG.COLUMN_MAP.EDIT);
        editCell.setNote(`LogID: ${logId}\n\nTo edit this item:\n1. Select this cell\n2. Go to Assets > Edit Selected Row`);
        editCell.setBackground('#e3f2fd');
        editCell.setFontColor('#1976d2');
        editCell.setFontWeight('bold');
      }
      
      // Apply dropdown validations to the new row
      this.applyDropdownsToRow(sheet, nextRow);
      
      // Force sort immediately after adding to move "New Asset" to top
      this.forceSort(sheet);
      
      return { success: true, message: `Asset added to row ${nextRow}`, rowNumber: nextRow, logId: logId, isUpdate: false };
    } catch (error) {
      console.error('Error adding project item:', error);
      return { success: false, message: `Error adding item: ${error.message}`, rowNumber: null, logId: null, isUpdate: false };
    }
  },

  applyDropdownsToRow: function(sheet, rowNum) {
    try {
      const items = dropdownManager.getDropdownValues('items');
      const materials = dropdownManager.getDropdownValues('materials');
      const statuses = dropdownManager.getDropdownValues('statuses');
      const venues = dropdownManager.getDropdownValues('venues');
      const areas = dropdownManager.getDropdownValues('areas');
      const productionStatuses = dropdownManager.getDropdownValues('productionStatuses');
      
      if (items.length > 0) {
        const itemRule = SpreadsheetApp.newDataValidation().requireValueInList(items, true).setAllowInvalid(true).build();
        sheet.getRange(rowNum, CONFIG.COLUMN_MAP.ITEM).setDataValidation(itemRule);
      }
      if (materials.length > 0) {
        const materialRule = SpreadsheetApp.newDataValidation().requireValueInList(materials, true).setAllowInvalid(true).build();
        sheet.getRange(rowNum, CONFIG.COLUMN_MAP.MATERIAL).setDataValidation(materialRule);
      }
      if (statuses.length > 0) {
        const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(statuses, true).setAllowInvalid(true).build();
        sheet.getRange(rowNum, CONFIG.COLUMN_MAP.STATUS).setDataValidation(statusRule);
      }
      if (venues.length > 0) {
        const venueRule = SpreadsheetApp.newDataValidation().requireValueInList(venues, true).setAllowInvalid(true).build();
        sheet.getRange(rowNum, CONFIG.COLUMN_MAP.VENUE).setDataValidation(venueRule);
      }
      if (areas.length > 0) {
        const areaRule = SpreadsheetApp.newDataValidation().requireValueInList(areas, true).setAllowInvalid(true).build();
        sheet.getRange(rowNum, CONFIG.COLUMN_MAP.AREA).setDataValidation(areaRule);
      }
      if (productionStatuses.length > 0) {
        const productionStatusRule = SpreadsheetApp.newDataValidation().requireValueInList(productionStatuses, true).setAllowInvalid(true).build();
        sheet.getRange(rowNum, CONFIG.COLUMN_MAP.PRODUCTION_STATUS).setDataValidation(productionStatusRule);
      }
    } catch (error) {
      console.error('Error applying dropdowns to row:', error);
    }
  },

  updateProjectItem: function(assetData, logIdPrefix, logSheetName) {
    try {
      const sheet = this.getActiveSheet();
      const rowNum = parseInt(assetData.originalRowNumber);
      if (!rowNum || isNaN(rowNum) || rowNum < 1) throw new Error(`Invalid row number: ${assetData.originalRowNumber}`);
      const existingId = sheet.getRange(rowNum, CONFIG.COLUMN_MAP.ID).getValue();
      
      // For edit mode, generate filename using existing ID instead of next sequential ID
      const assetId = existingId || this.getNextId(assetData.material, true, existingId);
      
      const dueDate = this.formatDate(assetData.dueDate);
      const strikeDate = this.formatDate(assetData.strikeDate);
      const dimensions = `${assetData.width}" x ${assetData.height}"`;
      const rowData = [
        assetId, assetData.area || '', assetData.asset || '', assetData.status || 'New Asset', dimensions, 
        assetData.quantity || '', assetData.item || '', assetData.material || '', dueDate, strikeDate, 
        assetData.venue || '', assetData.location || '', assetData.artworkUrl || '', assetData.imageLink || '', 
        assetData.doubleSided || false, assetData.dieCut || false, assetData.productionStatus || '', 'Edit'
      ];
      const range = sheet.getRange(rowNum, 1, 1, rowData.length - 1); // Exclude Edit column from bulk update
      range.setValues([rowData.slice(0, -1)]);
      const doubleSidedCell = sheet.getRange(rowNum, CONFIG.COLUMN_MAP.DOUBLE_SIDED);
      const diecutCell = sheet.getRange(rowNum, CONFIG.COLUMN_MAP.DIECUT);
      if (doubleSidedCell.getDataValidation() === null) doubleSidedCell.insertCheckboxes();
      if (diecutCell.getDataValidation() === null) diecutCell.insertCheckboxes();
      const rowRange = sheet.getRange(rowNum, 1, 1, CONFIG.COLUMN_MAP.EDIT);
      if (assetData.status === 'New Asset') rowRange.setFontColor('#0062e2');
      else if (assetData.status === 'Delivered') rowRange.setFontColor('#38AE74');
      else if (assetData.status === 'On Hold') rowRange.setFontColor('#f7c831');
      else rowRange.setFontColor('#000000');
      const logId = this.updateLogFormData(assetData.formData || assetData, rowNum, logIdPrefix, logSheetName);
      if (logId) {
        const editCell = sheet.getRange(rowNum, CONFIG.COLUMN_MAP.EDIT);
        editCell.setNote(`LogID: ${logId}\n\nTo edit this item:\n1. Select this cell\n2. Go to Assets > Edit Selected Row\n\nLast updated: ${new Date().toLocaleString()}`);
      }
      
      // Force sort after updating in case status changed
      this.forceSort(sheet);
      
      return { success: true, message: `Asset updated in row ${rowNum}`, rowNumber: rowNum, logId: logId, isUpdate: true };
    } catch (error) {
      console.error('Error updating project item:', error);
      return { success: false, message: `Error updating item: ${error.message}`, rowNumber: null, logId: null, isUpdate: false };
    }
  },

  forceSort: function(sheet) {
    try {
      // Use a small delay to ensure all data is written before sorting
      Utilities.sleep(100);
      ensureNewAssetsAtTop(sheet);
    } catch (error) {
      console.error('Error in forceSort:', error);
    }
  }
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('BBLMW25')
    .addItem('Add New Asset', 'openAssetApp')
    .addSeparator()
    .addItem('Edit Selected Row', 'editSelectedRow')
    .addSeparator()
    .addItem('Reorder Asset', 'openReorderAssetApp')
    .addSeparator()
    .addSubMenu(ui.createMenu('Additional')
      .addItem('Edit All Dropdowns', 'openDropdownEditor')
      .addItem('Update All File Names', 'updateFileNames')
      .addItem('Update Image Links', 'updateImageLinks'))
    .addToUi();
  setupAutoSort();
  materialIDManager.getOrCreateMaterialIDSheet();
  
  // Initialize Production Status dropdown sheet with default values
  initializeProductionStatusSheet();
}

function initializeProductionStatusSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName('ProductionStatusesList');
    
    if (!sheet) {
      // Create the sheet
      sheet = spreadsheet.insertSheet('ProductionStatusesList');
      sheet.getRange(1, 1).setValue('Production Statuses');
      sheet.getRange(1, 1).setFontWeight('bold');
      
      // Add initial values
      const initialValues = [
        ['Processing'],
        ['Printing'], 
        ['Cutting'],
        ['Finishing'],
        ['Ready'],
        ['Picked Up']
      ];
      
      sheet.getRange(2, 1, initialValues.length, 1).setValues(initialValues);
      sheet.hideSheet();
    }
  } catch (error) {
    console.error('Error initializing Production Status sheet:', error);
  }
}

function openDropdownEditor() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('DropdownEditor').setTitle('Dropdown Editor').setWidth(320);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function setupAutoSort() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEditInstallable' || trigger.getHandlerFunction() === 'onChangeInstallable') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  ScriptApp.newTrigger('onChangeInstallable').forSpreadsheet(SpreadsheetApp.getActive()).onChange().create();
}

function updateFileNames() {
  try {
    const ui = SpreadsheetApp.getUi();
    const mainFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const subfolders = mainFolder.getFoldersByName(CONFIG.INITIAL_FILES_FOLDER_NAME);
    if (!subfolders.hasNext()) {
      ui.alert('Folder Not Found', `The "${CONFIG.INITIAL_FILES_FOLDER_NAME}" folder was not found.`, ui.ButtonSet.OK);
      return;
    }
    const initialFilesFolder = subfolders.next();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
    if (!sheet) { ui.alert('Error', 'Master sheet not found.', ui.ButtonSet.OK); return; }
    const files = initialFilesFolder.getFiles();
    let renamedCount = 0, skippedCount = 0, linksAddedCount = 0;
    const renamedFiles = [];
    const fileMap = {}; // Map of asset IDs to file URLs
    
    // First pass: Rename files and build file map
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const idMatch = fileName.match(/^([A-Z]\d+)/);
      if (idMatch) {
        const assetId = idMatch[1];
        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        for (let i = 1; i < values.length; i++) {
          const rowId = values[i][CONFIG.COLUMN_MAP.ID - 1];
          if (rowId === assetId) {
            const assetName = values[i][CONFIG.COLUMN_MAP.ASSET - 1];
            if (assetName) {
              const lastDotIndex = fileName.lastIndexOf('.');
              const extension = lastDotIndex > -1 ? fileName.substring(lastDotIndex) : '';
              const sanitizedAssetName = assetName.toString().replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_-]/g, '_');
              const newFileName = `${assetId}_${CONFIG.PROJECT_CODE}_${sanitizedAssetName}${extension}`;
              if (fileName !== newFileName) {
                file.setName(newFileName);
                renamedFiles.push(`${fileName} → ${newFileName}`);
                renamedCount++;
              } else skippedCount++;
              
              // Store file URL for this asset ID
              fileMap[assetId] = file.getUrl();
            }
            break;
          }
        }
      }
    }
    
    // Second pass: Insert Drive links into Artwork column
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      const rowId = values[i][CONFIG.COLUMN_MAP.ID - 1];
      if (rowId && fileMap[rowId]) {
        const artworkCell = sheet.getRange(i + 1, CONFIG.COLUMN_MAP.ARTWORK);
        const currentValue = artworkCell.getValue();
        
        // Only update if cell is empty or different from the file URL
        if (!currentValue || currentValue !== fileMap[rowId]) {
          artworkCell.setValue(fileMap[rowId]);
          linksAddedCount++;
        }
      }
    }
    
    let message = `File Update Complete!\n\n✅ Renamed: ${renamedCount} file(s)\n⏭️ Skipped: ${skippedCount} file(s) (already correct)\n🔗 Links Added: ${linksAddedCount} artwork link(s)\n`;
    if (renamedFiles.length > 0) {
      message += `\nRenamed Files:\n${renamedFiles.slice(0, 10).join('\n')}`;
      if (renamedFiles.length > 10) message += `\n... and ${renamedFiles.length - 10} more`;
    }
    ui.alert('Update File Names', message, ui.ButtonSet.OK);
  } catch (error) {
    console.error('Error in updateFileNames:', error);
    SpreadsheetApp.getUi().alert('Error', 'Failed to update file names: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function updateImageLinks() {
  try {
    const result = driveManager.updateImageLinks();
    const ui = SpreadsheetApp.getUi();
    
    if (result.success) {
      ui.alert('Update Image Links', `Successfully updated ${result.updateCount} image links from the "${CONFIG.THUMBNAIL_FOLDER_NAME}" folder.`, ui.ButtonSet.OK);
    } else {
      ui.alert('Error', result.message, ui.ButtonSet.OK);
    }
  } catch (error) {
    console.error('Error in updateImageLinks:', error);
    SpreadsheetApp.getUi().alert('Error', 'Failed to update image links: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function onEditInstallable(e) { handleEdit(e); }

function onChangeInstallable(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
  if (sheet) ensureNewAssetsAtTop(sheet);
}

function handleEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.MASTER_SHEET_NAME) return;
  const range = e.range;
  const row = range.getRow();
  if (row < 2) return;
  ensureCheckboxesInRow(sheet, row);
  if (range.getColumn() === CONFIG.COLUMN_MAP.STATUS) {
    const newStatus = range.getValue();
    const rowRange = sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT);
    if (newStatus === 'New Asset') rowRange.setFontColor('#0062e2');
    else if (newStatus === 'Requires Attn - HX') {
      rowRange.setFontColor('#FF2093');
      showHXCommentDialog(sheet, row);
    } else if (newStatus === 'Delivered') rowRange.setFontColor('#38AE74');
    else if (newStatus === 'On Hold') rowRange.setFontColor('#f7c831');
    else if (newStatus && newStatus !== 'New Asset' && newStatus !== 'Requires Attn - HX' && newStatus !== 'Delivered' && newStatus !== 'On Hold') rowRange.setFontColor('#000000');
    ensureNewAssetsAtTop(sheet);
  }
  if (range.getColumn() !== CONFIG.COLUMN_MAP.STATUS && range.getColumn() !== CONFIG.COLUMN_MAP.EDIT) {
    const statusCell = sheet.getRange(row, CONFIG.COLUMN_MAP.STATUS);
    const currentStatus = statusCell.getValue();
    if (currentStatus === 'New Asset') sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT).setFontColor('#0062e2');
    else if (currentStatus === 'Requires Attn - HX') sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT).setFontColor('#FF2093');
    else if (currentStatus === 'Delivered') sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT).setFontColor('#38AE74');
    else if (currentStatus === 'On Hold') sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT).setFontColor('#f7c831');
  }
  if (range.getColumn() !== CONFIG.COLUMN_MAP.EDIT) logRowEdit(sheet, row);
}

function ensureCheckboxesInRow(sheet, row) {
  try {
    const doubleSidedCell = sheet.getRange(row, CONFIG.COLUMN_MAP.DOUBLE_SIDED);
    const diecutCell = sheet.getRange(row, CONFIG.COLUMN_MAP.DIECUT);
    const doubleSidedValue = doubleSidedCell.getValue();
    const diecutValue = diecutCell.getValue();
    const doubleSidedValidation = doubleSidedCell.getDataValidation();
    if (doubleSidedValidation === null || doubleSidedValidation.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      doubleSidedCell.insertCheckboxes();
      if (doubleSidedValue === 'TRUE' || doubleSidedValue === true) doubleSidedCell.setValue(true);
      else if (doubleSidedValue === 'FALSE' || doubleSidedValue === false) doubleSidedCell.setValue(false);
    }
    const diecutValidation = diecutCell.getDataValidation();
    if (diecutValidation === null || diecutValidation.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      diecutCell.insertCheckboxes();
      if (diecutValue === 'TRUE' || diecutValue === true) diecutCell.setValue(true);
      else if (diecutValue === 'FALSE' || diecutValue === false) diecutCell.setValue(false);
    }
  } catch (error) {
    console.error('Error ensuring checkboxes:', error);
  }
}

function showHXCommentDialog(sheet, row) {
  try {
    const rowData = sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT).getValues()[0];
    const assetId = rowData[CONFIG.COLUMN_MAP.ID - 1];
    const assetName = rowData[CONFIG.COLUMN_MAP.ASSET - 1];
    const material = rowData[CONFIG.COLUMN_MAP.MATERIAL - 1];
    const item = rowData[CONFIG.COLUMN_MAP.ITEM - 1];
    const htmlOutput = HtmlService.createHtmlOutputFromFile('HXCommentDialog').setWidth(420).setHeight(280);
    const htmlContent = htmlOutput.getContent();
    const modifiedContent = htmlContent.replace('<script>', `<script>window.assetData = ${JSON.stringify({ row: row, assetId: assetId || 'N/A', assetName: assetName || 'N/A', item: item || 'N/A', material: material || 'N/A' })};`);
    const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent).setWidth(420).setHeight(280);
    SpreadsheetApp.getUi().showModalDialog(modifiedOutput, '⚠️ HX Attention Required');
  } catch (error) {
    console.error('Error showing HX comment dialog:', error);
  }
}

function sendHXNotification(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
    const row = data.row;
    const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    const sheetName = sheet.getName();
    const cellReference = `${sheetName}!D${row}`;
    const directLink = `${spreadsheetUrl}#gid=${sheet.getSheetId()}&range=D${row}`;
    const emailSubject = `🚨 HX Attention Required - Asset ${data.assetId}`;
    const emailBody = `<html><body style="font-family: Arial, sans-serif; color: #333;"><h2 style="color: #FF2093;">⚠️ Asset Requires Attention</h2><div style="background-color: #fef2f2; padding: 15px; border-left: 4px solid #FF2093; margin: 20px 0;"><p><strong>Asset Name:</strong> ${data.assetName}</p><p><strong>Asset ID:</strong> ${data.assetId}</p><p><strong>Item:</strong> ${data.item}</p><p><strong>Material:</strong> ${data.material}</p><p><strong>Location:</strong> ${cellReference}</p></div><div style="background-color: #f9fafb; padding: 15px; border-radius: 6px; margin: 20px 0;"><p style="margin: 0;"><strong>Comment:</strong></p><p style="margin: 10px 0 0 0; white-space: pre-wrap;">${data.comment}</p></div><p>This asset has been flagged as <strong style="color: #FF2093;">"Requires Attn - HX"</strong> and needs your immediate attention.</p><p style="margin-top: 30px;"><a href="${directLink}" style="background-color: #FF2093; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px;">View Asset in Spreadsheet</a></p><hr style="margin-top: 30px; border: none; border-top: 1px solid #ddd;"><p style="font-size: 12px; color: #666;">This is an automated notification from the BBLMW25 Asset Management System.</p></body></html>`;
    try { MailApp.sendEmail({ to: 'diana@hxecute.com', subject: emailSubject, htmlBody: emailBody }); } catch (emailError) { console.error('Error sending email to diana@hxecute.com:', emailError); }
    try { MailApp.sendEmail({ to: 'heather@hxecute.com', subject: emailSubject, htmlBody: emailBody }); } catch (emailError) { console.error('Error sending email to heather@hxecute.com:', emailError); }
    return { success: true, message: 'Notifications sent successfully' };
  } catch (error) {
    console.error('Error sending HX notification:', error);
    return { success: false, message: error.toString() };
  }
}

function logRowEdit(sheet, row) {
  try {
    const rowData = sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT).getValues()[0];
    const columnsToCheck = [CONFIG.COLUMN_MAP.ID - 1, CONFIG.COLUMN_MAP.ASSET - 1, CONFIG.COLUMN_MAP.DIMENSIONS - 1, CONFIG.COLUMN_MAP.QUANTITY - 1, CONFIG.COLUMN_MAP.DUE_DATE - 1, CONFIG.COLUMN_MAP.STRIKE_DATE - 1, CONFIG.COLUMN_MAP.LOCATION - 1, CONFIG.COLUMN_MAP.ARTWORK - 1];
    const hasContent = columnsToCheck.some(index => { const cell = rowData[index]; return cell !== '' && cell !== null && cell !== undefined; });
    if (!hasContent) return;
    const parseDateToISO = (dateValue) => {
      if (!dateValue) return '';
      if (dateValue instanceof Date) return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const dateStr = dateValue.toString();
      const dateMatch = dateStr.match(/\w+,\s+(\w+)\s+(\d+),\s+(\d+)/);
      if (dateMatch) {
        const date = new Date(`${dateMatch[1]} ${dateMatch[2]}, ${dateMatch[3]}`);
        return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      return '';
    };
    let width = '', height = '';
    const dimensions = rowData[CONFIG.COLUMN_MAP.DIMENSIONS - 1];
    if (dimensions) {
      const dimMatch = dimensions.toString().match(/^([\d.]+)"\s*x\s*([\d.]+)"$/);
      if (dimMatch) { width = dimMatch[1]; height = dimMatch[2]; }
    }
    const formData = {
      originalRowNumber: row, item: rowData[CONFIG.COLUMN_MAP.ITEM - 1] || '', material: rowData[CONFIG.COLUMN_MAP.MATERIAL - 1] || '',
      asset: rowData[CONFIG.COLUMN_MAP.ASSET - 1] || '', quantity: rowData[CONFIG.COLUMN_MAP.QUANTITY - 1] || '', width: width, height: height,
      dieCut: (rowData[CONFIG.COLUMN_MAP.DIECUT - 1] === true), doubleSided: (rowData[CONFIG.COLUMN_MAP.DOUBLE_SIDED - 1] === true),
      status: rowData[CONFIG.COLUMN_MAP.STATUS - 1] || '', dueDate: parseDateToISO(rowData[CONFIG.COLUMN_MAP.DUE_DATE - 1]),
      strikeDate: parseDateToISO(rowData[CONFIG.COLUMN_MAP.STRIKE_DATE - 1]), venue: rowData[CONFIG.COLUMN_MAP.VENUE - 1] || '',
      area: rowData[CONFIG.COLUMN_MAP.AREA - 1] || '', location: rowData[CONFIG.COLUMN_MAP.LOCATION - 1] || '', 
      artworkUrl: rowData[CONFIG.COLUMN_MAP.ARTWORK - 1] || '', imageLink: rowData[CONFIG.COLUMN_MAP.IMAGE_LINK - 1] || '',
      productionStatus: rowData[CONFIG.COLUMN_MAP.PRODUCTION_STATUS - 1] || ''
    };
    const logId = projectSheet.updateLogFormData(formData, row, CONFIG.LOG_ID_PREFIX, CONFIG.LOG_SHEET_NAME);
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

function ensureNewAssetsAtTop(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    const dataRange = sheet.getRange(2, 1, lastRow - 1, CONFIG.COLUMN_MAP.EDIT);
    const values = dataRange.getValues();
    const newAssetRows = [], otherRows = [];
    values.forEach((row, index) => {
      const status = row[CONFIG.COLUMN_MAP.STATUS - 1];
      if (status === 'New Asset') newAssetRows.push(row);
      else otherRows.push(row);
    });
    otherRows.sort((a, b) => {
      const idA = a[0] ? a[0].toString() : '';
      const idB = b[0] ? b[0].toString() : '';
      return idA.localeCompare(idB);
    });
    const sortedData = [...newAssetRows, ...otherRows];
    const currentOrder = values.map(row => row.join('|'));
    const newOrder = sortedData.map(row => row.join('|'));
    if (currentOrder.join('||') !== newOrder.join('||')) {
      dataRange.setValues(sortedData);
      for (let i = 0; i < sortedData.length; i++) {
        const rowNum = i + 2;
        const status = sortedData[i][CONFIG.COLUMN_MAP.STATUS - 1];
        const rowRange = sheet.getRange(rowNum, 1, 1, CONFIG.COLUMN_MAP.EDIT);
        if (status === 'New Asset') rowRange.setFontColor('#0062e2');
        else if (status === 'Requires Attn - HX') rowRange.setFontColor('#FF2093');
        else if (status === 'Delivered') rowRange.setFontColor('#38AE74');
        else rowRange.setFontColor('#000000');
      }
    }
  } catch (error) {
    console.error('Error in ensureNewAssetsAtTop:', error);
  }
}

function sortNonNewAssets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(CONFIG.MASTER_SHEET_NAME);
  if (!sheet) return;
  ensureNewAssetsAtTop(sheet);
}

function onEdit(e) { handleEdit(e); }
function openAssetApp() { assetApp.showDialog(); }

function editSelectedRow() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    if (sheet.getName() !== CONFIG.MASTER_SHEET_NAME) {
      SpreadsheetApp.getUi().alert('Edit Row', `Please select a row in the "${CONFIG.MASTER_SHEET_NAME}" sheet to edit.`, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    const activeCell = sheet.getActiveCell();
    const rowNumber = activeCell.getRow();
    if (rowNumber < 2) {
      SpreadsheetApp.getUi().alert('Edit Row', 'Please select a data row (not the header row) to edit.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    assetApp.openForEdit(rowNumber);
  } catch (error) {
    console.error('Error in editSelectedRow:', error);
    SpreadsheetApp.getUi().alert('Error', 'An error occurred while trying to edit the row: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function getDropdownData() { return dropdownManager.getAllDropdownData(); }
function addNewDropdownValue(fieldName, value) { return dropdownManager.addDropdownValue(fieldName, value); }
function updateDropdownValue(fieldName, oldValue, newValue) { return dropdownManager.updateDropdownValue(fieldName, oldValue, newValue); }
function deleteDropdownValue(fieldName, value) { return dropdownManager.deleteDropdownValue(fieldName, value); }
function getDriveFolders() { return driveManager.getFolderStructure(); }
function uploadFileToDrive(fileData, folderId, fileName) { return driveManager.uploadFile(fileData, folderId, fileName); }
function addAssetToProject(assetData) { return assetApp.addToProject(assetData); }
function getMaterialPrefix(material) { return projectSheet.getMaterialPrefix(material); }
function getNextIdForMaterial(material) { return projectSheet.getNextIdForMaterial(material); }
function getRowDataForEdit(rowNumber) { 
  const rowData = projectSheet.getRowData(rowNumber);
  if (rowData) {
    // Get the ID from the spreadsheet for filename generation
    const sheet = projectSheet.getActiveSheet();
    const idValue = sheet.getRange(rowNumber, CONFIG.COLUMN_MAP.ID).getValue();
    return { ...rowData, id: idValue };
  }
  return null;
}

// Reorder Asset Functions
function openReorderAssetApp() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('ReorderAssetForm')
      .setWidth(600).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Reorder Asset');
}

function getAvailableAssetsForReorder() {
  try {
    const sheet = projectSheet.getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) return [];
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, CONFIG.COLUMN_MAP.EDIT);
    const values = dataRange.getValues();
    
    const availableAssets = [];
    const seenAssets = new Set();
    
    for (let i = 0; i < values.length; i++) {
      const assetId = values[i][CONFIG.COLUMN_MAP.ID - 1];
      const assetName = values[i][CONFIG.COLUMN_MAP.ASSET - 1];
      
      if (!assetName) continue;
      
      // Check if this is a concatenated/reordered asset (contains " - " pattern)
      // We want to exclude assets that match the pattern: "ID - Asset Name"
      const isConcatenated = /^[A-Z]\d+\s*-\s*.+/.test(assetName.toString());
      
      // Only add if not concatenated and not already seen
      if (!isConcatenated && !seenAssets.has(assetName)) {
        seenAssets.add(assetName);
        availableAssets.push({
          name: assetName,
          rowNumber: i + 2 // Actual row number in sheet
        });
      }
    }
    
    // Sort alphabetically
    availableAssets.sort((a, b) => a.name.localeCompare(b.name));
    
    return availableAssets;
  } catch (error) {
    console.error('Error getting available assets:', error);
    throw new Error('Failed to load assets: ' + error.message);
  }
}

function reorderAsset(reorderData) {
  try {
    const sheet = projectSheet.getActiveSheet();
    const originalRowNumber = parseInt(reorderData.originalRowNumber);
    
    if (!originalRowNumber || originalRowNumber < 2) {
      return { success: false, message: 'Invalid asset selection' };
    }
    
    // Get all data from the original row
    const originalRange = sheet.getRange(originalRowNumber, 1, 1, CONFIG.COLUMN_MAP.EDIT);
    const originalData = originalRange.getValues()[0];
    
    // Extract original values
    const originalId = originalData[CONFIG.COLUMN_MAP.ID - 1];
    const originalAsset = originalData[CONFIG.COLUMN_MAP.ASSET - 1];
    const originalMaterial = originalData[CONFIG.COLUMN_MAP.MATERIAL - 1];
    
    // Generate new ID based on material
    const newId = projectSheet.getNextId(originalMaterial);
    
    // Create concatenated asset name: "OriginalID - OriginalAssetName"
    const newAssetName = `${originalId} - ${originalAsset}`;
    
    // Insert new row at the bottom
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    
    // Build row data with modifications
    const rowData = [
      newId, // New sequential ID
      originalData[CONFIG.COLUMN_MAP.AREA - 1] || '', // Area
      newAssetName, // Concatenated asset name
      'New Asset', // Status - always "New Asset"
      originalData[CONFIG.COLUMN_MAP.DIMENSIONS - 1] || '', // Dimensions
      reorderData.quantity, // New quantity from form
      originalData[CONFIG.COLUMN_MAP.ITEM - 1] || '', // Item
      originalMaterial || '', // Material
      originalData[CONFIG.COLUMN_MAP.DUE_DATE - 1] || '', // Due Date
      originalData[CONFIG.COLUMN_MAP.STRIKE_DATE - 1] || '', // Strike Date
      originalData[CONFIG.COLUMN_MAP.VENUE - 1] || '', // Venue
      originalData[CONFIG.COLUMN_MAP.LOCATION - 1] || '', // Location
      originalData[CONFIG.COLUMN_MAP.ARTWORK - 1] || '', // Artwork URL
      originalData[CONFIG.COLUMN_MAP.IMAGE_LINK - 1] || '', // Image Link
      originalData[CONFIG.COLUMN_MAP.DOUBLE_SIDED - 1] || false, // Double Sided
      originalData[CONFIG.COLUMN_MAP.DIECUT - 1] || false, // Diecut
      '', // Production Status - empty
      'Edit' // Edit column
    ];
    
    // Write the new row
    const range = sheet.getRange(newRow, 1, 1, rowData.length);
    range.setValues([rowData]);
    
    // Set up checkboxes
    sheet.getRange(newRow, CONFIG.COLUMN_MAP.DOUBLE_SIDED).insertCheckboxes();
    sheet.getRange(newRow, CONFIG.COLUMN_MAP.DIECUT).insertCheckboxes();
    
    // Apply "New Asset" blue color
    range.setFontColor('#0062e2');
    
    // Apply dropdown validations to the new row
    projectSheet.applyDropdownsToRow(sheet, newRow);
    
    // Set up Edit cell formatting
    const editCell = sheet.getRange(newRow, CONFIG.COLUMN_MAP.EDIT);
    editCell.setBackground('#e3f2fd');
    editCell.setFontColor('#1976d2');
    editCell.setFontWeight('bold');
    editCell.setNote(`Reordered from: ${originalAsset} (ID: ${originalId})\nCreated: ${new Date().toLocaleString()}`);
    
    // Log the reorder in the AssetLog
    const logData = {
      originalRowNumber: newRow,
      reorderedFrom: originalId,
      originalAsset: originalAsset,
      newQuantity: reorderData.quantity
    };
    projectSheet.logFormData(logData, newRow, CONFIG.LOG_ID_PREFIX, CONFIG.LOG_SHEET_NAME);
    
    // Force sort to move "New Asset" to top
    projectSheet.forceSort(sheet);
    
    return { 
      success: true, 
      message: `Asset reordered successfully! New ID: ${newId}`,
      newId: newId,
      rowNumber: newRow
    };
    
  } catch (error) {
    console.error('Error reordering asset:', error);
    return { success: false, message: 'Error reordering asset: ' + error.message };
  }
}
