/**
 * =================================================================
 *DYNAMIC QA SHEET AUTOMATION FRAMEWORK
 * =================================================================
**/

// =================================================================
// SECTION 0: CONFIGURATION
// =================================================================

const CONFIG = {
    API_BASE_URL: "https://product-counter-api.onrender.com",
    PROP_KEYS: {
        SHOP_ID: 'SHOP_ID',
        JOB_ID: 'PC_FETCH_JOB_ID',
        JOB_STATUS: 'PC_FETCH_STATUS',
        SHEET_NAME: 'PC_FETCH_SHEET_NAME'
    },
    SHEET_ROWS: {
        KEYWORD_GEN: 14,
        MANUAL: 20,
        SEMANTIC: 21,
        STAGING: 36
    },
    HEADER_NAMES: {
        KEYWORD: 'keyword',
        PC: 'pc',
        SEARCH_VOLUME: 'search volume',
        URL: 'url',
        TITLE: 'title',
        PAGE_NAME: 'page_name',
        PRODUCT_IDS: 'product_ids'
    },
    POSSIBLE_KEYWORD_HEADERS: ['keyword', 'keywords', 'original_keyword', 'query', 'cleanup_keyword'],
    POSSIBLE_SV_HEADERS: ['search volume', 'traffic', 'sv', 'volume'],
    TOAST_TITLE: 'QA Sheet Automation'
};

// Global cache for sheet headers
let headerCache = null;

// =================================================================
// SECTION 1: SHEET CREATION AND MENU INSTALLATION
// =================================================================

/**
Adds a new menu item for trigger cleanup.
 */
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Automation Tools')
        .addItem('Create QA Sheet', 'createQaSheet')
        .addSeparator()
        .addItem('Delete ALL Triggers (Troubleshooting)', 'deleteAllProjectTriggers')
        .addToUi();
}

function onOpenQaSheet() {
    createQaSheetMenu();
    initializeFirstRunWorkflow();
}

function createQaSheetMenu() {
    SpreadsheetApp.getUi()
        .createMenu('QA Tools') // Renamed for clarity
        .addItem('Start PC Fetch (All Keywords)', 'startProductCountJobAll')
        .addItem('Start PC Fetch (Filtered Keywords Only)', 'startProductCountJobFiltered')
        .addSeparator()
        .addItem('Validate Selected Keywords', 'validateSelectedKeywords')
        .addItem('Validate Selected URLs', 'validateSelectedUrls')
        .addToUi();
}

function initializeFirstRunWorkflow() {
    const docProperties = PropertiesService.getDocumentProperties();
    if (docProperties.getProperty(CONFIG.PROP_KEYS.JOB_ID)) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`Manual Product Count fetch is already in progress.`, 'Automation Status', 10);
    }
}


function installPermanentOnOpenTrigger(spreadsheetId) {
    const triggerFunctionName = 'onOpenQaSheet';
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
        if (trigger.getTriggerSourceId() === spreadsheetId && trigger.getHandlerFunction() === triggerFunctionName) {
            ScriptApp.deleteTrigger(trigger);
        }
    }
    ScriptApp.newTrigger(triggerFunctionName)
        .forSpreadsheet(spreadsheetId)
        .onOpen()
        .create();
}


function createQaSheet() {
    const ui = SpreadsheetApp.getUi();
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeColumn = activeSheet.getActiveCell().getColumn();
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const sheetName = activeSheet.getName();

    const keywordGenCellA1 = activeSheet.getRange(CONFIG.SHEET_ROWS.KEYWORD_GEN, activeColumn).getA1Notation();
    const manualCellA1 = activeSheet.getRange(CONFIG.SHEET_ROWS.MANUAL, activeColumn).getA1Notation();
    const semanticCellA1 = activeSheet.getRange(CONFIG.SHEET_ROWS.SEMANTIC, activeColumn).getA1Notation();
    const stagingCellA1 = activeSheet.getRange(CONFIG.SHEET_ROWS.STAGING, activeColumn).getA1Notation();


    const rangesToFetch = [
        `'${sheetName}'!${manualCellA1}`,
        `'${sheetName}'!${semanticCellA1}`,
        `'${sheetName}'!${stagingCellA1}`,
        `'${sheetName}'!${keywordGenCellA1}`
    ];

    try {
        const response = Sheets.Spreadsheets.get(spreadsheetId, {
            ranges: rangesToFetch,
            fields: 'sheets(data(rowData(values(chipRuns(chip(richLinkProperties(uri)))))))'
        });

        const getUriFromResponse = (data) => (data.rowData?.[0]?.values?.[0]?.chipRuns?.[0]?.chip?.richLinkProperties?.uri) || null;

        const manualUrl = getUriFromResponse(response.sheets[0].data[0]);
        const semanticUrl = getUriFromResponse(response.sheets[0].data[1]);
        const stagingUrl = getUriFromResponse(response.sheets[0].data[2]);
        let keywordGenUrl = getUriFromResponse(response.sheets[0].data[3]);

        if (!semanticUrl) {
            ui.alert(`Error: Could not find a link in cell ${semanticCellA1} (Semantic). This is a required file.`);
            return;
        }
        if (!keywordGenUrl) {
             if (ui.alert('Keyword Gen Sheet Missing', `The Keyword Gen sheet was not found in cell ${keywordGenCellA1}. Proceed without Product IDs and calculated PC?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
        }

        if (!manualUrl && ui.alert('Manual Sheet Missing', 'The Manual Dedupe sheet was not found. Proceed without search volumes?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
        if (!stagingUrl && ui.alert('Staging Sheet Missing', 'The Staging sheet was not found. Proceed without URL/Title data?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

        const semanticFileId = semanticUrl.match(/[-\w]{25,}/)[0];
        const sourceSpreadsheet = SpreadsheetApp.openById(semanticFileId);
        const sourceSheetName = sourceSpreadsheet.getName();
        const nameParts = sourceSheetName.split(' ');
        if (nameParts.length < 2) {
            ui.alert(`Error: Linked sheet's name "${sourceSheetName}" has an unexpected format.`);
            return;
        }

        const outputFileName = `${nameParts[0]} ${nameParts[1]} QA`;
        if (ui.alert(`Create QA sheet for "${nameParts[0]}"?`, `This will create a new file named "${outputFileName}".\n\nContinue?`, ui.ButtonSet.YES_NO) === ui.Button.YES) {
            processAndEnrichSheet(semanticUrl, manualUrl, stagingUrl, keywordGenUrl, outputFileName);
        }
    } catch (e) {
        ui.alert('An error occurred during setup: ' + e.toString());
    }
}


function processAndEnrichSheet(semanticSourceUrl, manualSourceUrl, stagingSourceUrl, keywordGenSourceUrl, outputFileName) {
    const ui = SpreadsheetApp.getUi();
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    try {
        activeSpreadsheet.toast('Step 1/7: Creating copy...', CONFIG.TOAST_TITLE, -1);
        const semanticFileId = semanticSourceUrl.match(/[-\w]{25,}/)[0];
        const qaSpreadsheet = SpreadsheetApp.openById(semanticFileId).copy(outputFileName);
        const qaSpreadsheetId = qaSpreadsheet.getId();
        const qaSpreadsheetUrl = qaSpreadsheet.getUrl();
        const newSS = SpreadsheetApp.openById(qaSpreadsheetId);

        activeSpreadsheet.toast('Step 2/7: Reformatting sheet...', CONFIG.TOAST_TITLE, -1);
        const qaSheet = rebuildAndOrderSheet(newSS, "Combined Final Data", "QA Data");
        if (!qaSheet) return;
        headerCache = null;

        activeSpreadsheet.toast('Step 3/7: Accessing keyword gen sheet...', CONFIG.TOAST_TITLE, -1);
        let keywordGenMap = null;
        if (keywordGenSourceUrl) {
            try { keywordGenMap = createKeywordGenMap(SpreadsheetApp.openById(keywordGenSourceUrl.match(/[-\w]{25,}/)[0])); }
            catch (e) { Logger.log(`Could not process optional keyword gen sheet. Error: ${e.toString()}`); }
        }

        activeSpreadsheet.toast('Step 4/7: Accessing manual sheet...', CONFIG.TOAST_TITLE, -1);
        let svMap = null;
        if (manualSourceUrl) {
            try { svMap = createSearchVolumeMap(SpreadsheetApp.openById(manualSourceUrl.match(/[-\w]{25,}/)[0])); }
            catch (e) { Logger.log(`Could not process optional manual sheet. Error: ${e.toString()}`); }
        }

        activeSpreadsheet.toast('Step 5/7: Accessing staging sheet...', CONFIG.TOAST_TITLE, -1);
        let stagingMap = null;
        if (stagingSourceUrl) {
            try { stagingMap = createStagingDataMap(SpreadsheetApp.openById(stagingSourceUrl.match(/[-\w]{25,}/)[0])); }
            catch (e) { Logger.log(`Could not process optional staging sheet. Error: ${e.toString()}`); }
        }
        
        activeSpreadsheet.toast('Step 6/7: Matching data & calculating PC...', CONFIG.TOAST_TITLE, -1);
        enrichAndFinalizeSheet(qaSheet, svMap, stagingMap, keywordGenMap);

        activeSpreadsheet.toast('Step 7/7: Hiding columns & Finalizing...', CONFIG.TOAST_TITLE, -1);
        hideQaColumns(qaSheet);
        installPermanentOnOpenTrigger(qaSpreadsheetId);
        
        activeSpreadsheet.toast('Process complete! New QA sheet is ready.', CONFIG.TOAST_TITLE, 10);
        showCompletionDialog(outputFileName, qaSpreadsheetUrl);

    } catch (e) {
        ui.alert('A critical error occurred: ' + e.toString());
    }
}


// =================================================================
// SECTION 2: AUTOMATED & MANUAL PC FETCHER LOGIC
// =================================================================

/**
 * Menu handler to fetch PCs for ALL keywords in the sheet.
 * This method is fast as it ignores filters.
 */
function startProductCountJobAll() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const headers = getSheetHeaders(sheet);
    const keywordColIndex = headers.indexOf(CONFIG.HEADER_NAMES.KEYWORD);
    
    if (keywordColIndex === -1) {
        SpreadsheetApp.getUi().alert(`Error: "${CONFIG.HEADER_NAMES.KEYWORD}" column not found in this sheet.`);
        return;
    }

    const keywords = sheet.getRange(2, keywordColIndex + 1, sheet.getLastRow() - 1).getValues().flat().filter(kw => kw && String(kw).trim());
    executePcJob(keywords, "All");
}

/**
 * Menu handler to fetch PCs only for keywords visible in the current filter.
 * This method is slower as it must check each row individually.
 */
function startProductCountJobFiltered() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const headers = getSheetHeaders(sheet);
    const keywordColIndex = headers.indexOf(CONFIG.HEADER_NAMES.KEYWORD);

    if (keywordColIndex === -1) {
        SpreadsheetApp.getUi().alert(`Error: "${CONFIG.HEADER_NAMES.KEYWORD}" column not found in this sheet.`);
        return;
    }
    
    const keywords = [];
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        ss.toast('Analyzing filtered rows... this may take a moment.', 'Processing', 5);
        const keywordColData = sheet.getRange(2, keywordColIndex + 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < keywordColData.length; i++) {
            const rowNum = i + 2;
            if (!sheet.isRowHiddenByFilter(rowNum)) {
                const kw = keywordColData[i][0];
                if (kw && typeof kw === 'string' && kw.trim() !== '') {
                    keywords.push(kw.trim());
                }
            }
        }
    }
    
    executePcJob(keywords, "Filtered");
}

/**
 * Core function that executes the Product Count job.
 * @param {string[]} keywords The array of keywords to process.
 * @param {string} type The type of job for logging ('All' or 'Filtered').
 */
function executePcJob(keywords, type) {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const docProperties = PropertiesService.getDocumentProperties();

    if (docProperties.getProperty(CONFIG.PROP_KEYS.JOB_ID)) {
        if (ui.alert('Job In Progress', 'A job is already running. Start a new one anyway?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    }
    
    let shopId; // Define shopId variable
    const response = ui.prompt('Shop ID Required', 'Please enter the Shop ID to fetch product counts:', ui.ButtonSet.OK_CANCEL);
    const button = response.getSelectedButton();
    const responseText = response.getResponseText();

    if (button === ui.Button.OK && responseText.trim() !== '') {
        shopId = responseText.trim();
    } else {
        ui.alert('Shop ID is required to start the job. Operation cancelled.');
        return;
    }

    if (keywords.length === 0) {
        ui.alert(`No ${type === 'Filtered' ? 'visible' : ''} keywords found to process.`);
        return;
    }

    ss.toast(`Sending ${keywords.length} keywords from "${ss.getActiveSheet().getName()}"...`, `${type} Job Start`, 10);
    const payload = { shop_id: shopId, keywords: keywords };
    const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };

    try {
        const response = UrlFetchApp.fetch(`${CONFIG.API_BASE_URL}/start_job`, options);
        const responseText = response.getContentText();
        if (response.getResponseCode() === 200) {
            const jobId = JSON.parse(responseText).job_id;
            docProperties.setProperties({
                [CONFIG.PROP_KEYS.JOB_ID]: jobId,
                [CONFIG.PROP_KEYS.SHEET_NAME]: ss.getActiveSheet().getName()
            });
            docProperties.deleteProperty(CONFIG.PROP_KEYS.JOB_STATUS);
            ss.toast(`Manual Job Started! ID: ${jobId}. Results will populate automatically.`, 'Job Sent', 20);
            createPollingTrigger();
        } else {
            ui.alert(`Error starting job: ${responseText}`);
        }
    } catch (e) {
        ui.alert(`A network error occurred: ${e.toString()}`);
    }
}


function createPollingTrigger() {
    deletePollingTrigger();
    ScriptApp.newTrigger('pollForResults').timeBased().everyMinutes(1).create();
    Logger.log('Polling trigger created.');
}

function deletePollingTrigger() {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
        if (trigger.getHandlerFunction() === 'pollForResults') {
            ScriptApp.deleteTrigger(trigger);
            Logger.log('Deleted existing polling trigger.');
        }
    }
}


function pollForResults() {
    const docProperties = PropertiesService.getDocumentProperties();
    const properties = docProperties.getProperties();
    const jobId = properties[CONFIG.PROP_KEYS.JOB_ID];
    const targetSheetName = properties[CONFIG.PROP_KEYS.SHEET_NAME];

    const cleanup = () => {
        docProperties.deleteProperty(CONFIG.PROP_KEYS.JOB_ID);
        docProperties.deleteProperty(CONFIG.PROP_KEYS.SHEET_NAME);
        deletePollingTrigger();
        headerCache = null;
    };

    if (!jobId) {
        deletePollingTrigger();
        return;
    }
    if (!targetSheetName) {
        Logger.log(`Job ID ${jobId} found, but no target sheet name. Cleaning up.`);
        cleanup();
        return;
    }

    Logger.log(`Polling for Job ID: ${jobId} on sheet: ${targetSheetName}`);
    const options = { method: 'get', muteHttpExceptions: true };

    try {
        const response = UrlFetchApp.fetch(`${CONFIG.API_BASE_URL}/get_results/${jobId}`, options);
        const data = JSON.parse(response.getContentText());
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        switch (data.status) {
            case "complete":
                Logger.log(`Job ${jobId} is complete. Writing results.`);
                const sheet = ss.getSheetByName(targetSheetName);
                if (!sheet) {
                    Logger.log(`Target sheet "${targetSheetName}" no longer exists. Aborting.`);
                    cleanup();
                    return;
                }

                let headers = getSheetHeaders(sheet);
                let pcColIndex = headers.indexOf(CONFIG.HEADER_NAMES.PC);

                if (pcColIndex === -1) {
                    Logger.log(`'PC' column not found on sheet "${targetSheetName}". Creating it.`);
                    pcColIndex = sheet.getLastColumn();
                    sheet.getRange(1, pcColIndex + 1).setValue(CONFIG.HEADER_NAMES.PC.toUpperCase());
                    headerCache = null;
                }

                const results = data.results.map(count => [count ?? 0]);
                if (results.length > 0) {
                    sheet.getRange(2, pcColIndex + 1, results.length, 1).setValues(results);
                }

                applyPcFilter(sheet);
                docProperties.setProperty(CONFIG.PROP_KEYS.JOB_STATUS, 'COMPLETE');
                cleanup();
                ss.toast(`Product Counts for "${targetSheetName}" populated!`, "Automation Complete", 30);
                break;

            case "processing":
                Logger.log(`Job ${jobId} is still processing.`);
                break;

            case "failed":
            case "not_found":
                Logger.log(`Job ${jobId} failed or was not found. Status: ${data.status}.`);
                docProperties.setProperty(CONFIG.PROP_KEYS.JOB_STATUS, `FAILED: ${data.status}`);
                cleanup();
                break;

            default:
                Logger.log(`Unknown server status: ${data.status}`);
                break;
        }
    } catch (e) {
        Logger.log(`Polling network error: ${e.toString()}`);
    }
}


function applyPcFilter(sheet) {
    if (sheet.getFilter()) {
        sheet.getFilter().remove();
    }
    const headers = getSheetHeaders(sheet);
    const pcColIndex = headers.indexOf(CONFIG.HEADER_NAMES.PC);

    if (pcColIndex !== -1) {
        const range = sheet.getDataRange();
        const filter = range.createFilter();
        const filterCriteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(['0', '1', '2']).build();
        filter.setColumnFilterCriteria(pcColIndex + 1, filterCriteria);
        Logger.log(`Filter applied to PC column on sheet "${sheet.getName()}".`);
    } else {
        Logger.log(`Could not find PC column to apply filter on sheet "${sheet.getName()}".`);
    }
}


// =================================================================
// SECTION 3: UTILITY AND TROUBLESHOOTING FUNCTIONS
// =================================================================

function deleteAllProjectTriggers() {
    const ui = SpreadsheetApp.getUi();
    const confirmation = ui.alert(
        'Delete ALL Triggers?',
        'This will delete ALL triggers for this script, including all polling and onOpen triggers. This is used for troubleshooting if you get a "too many triggers" error. Are you sure you want to continue?',
        ui.ButtonSet.YES_NO
    );

    if (confirmation !== ui.Button.YES) {
        ui.alert('No triggers were deleted.');
        return;
    }

    try {
        const triggers = ScriptApp.getProjectTriggers();
        if (triggers.length === 0) {
            ui.alert('There were no triggers to delete.');
            return;
        }

        let deletedCount = 0;
        for (const trigger of triggers) {
            ScriptApp.deleteTrigger(trigger);
            deletedCount++;
            Logger.log(`Deleted trigger with ID: ${trigger.getUniqueId()}`);
        }
        ui.alert(`Successfully deleted ${deletedCount} trigger(s).`);
    } catch (e) {
        ui.alert('An error occurred while deleting triggers: ' + e.toString());
    }
}


function getSheetHeaders(sheet) {
    const spreadsheetId = sheet.getParent().getId();
    const sheetName = sheet.getName();
    if (headerCache && headerCache.spreadsheetId === spreadsheetId && headerCache.sheetName === sheetName) {
        return headerCache.headers;
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().toLowerCase().trim());
    headerCache = { headers, spreadsheetId, sheetName };
    return headers;
}

function showCompletionDialog(fileName, fileUrl) {
    const html = `<style>body{font-family:'Roboto',sans-serif;padding:15px}a{display:inline-block;padding:10px 20px;background-color:#4285F4;color:white;text-decoration:none;border-radius:4px;margin-top:15px}a:hover{background-color:#357ae8}</style><body><h3>Automation Complete!</h3><p>The new QA sheet "<b>${fileName}</b>" is ready.</p><a href="${fileUrl}" target="_blank">Open New Sheet</a></body>`;
    const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(350).setHeight(180);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Process Finished');
}

function rebuildAndOrderSheet(spreadsheet, sourceSheetName, targetSheetName) {
    const finalHeaders = ['Keyword', 'sensical editor', 'assortment editor', 'corrected_keyword', 'pluralised_keyword', 'cleanup_keyword', 'regional_keyword', 'Title', 'page_name', 'keyword_sources', 'is_subjective', 'is_sensical', 'product_ids', 'PC', 'Search Volume', 'Current Ranking(if applicable)', 'Sensical QA', 'sandbox link', 'Plural', 'Assortment QA', 'Comment', 'Completeness QA- comments', 'RE', 'Nearest Removals', 'Google Dupes', 'URL'];
    const sourceSheet = spreadsheet.getSheetByName(sourceSheetName);
    if (!sourceSheet) {
        SpreadsheetApp.getUi().alert(`Error: Source tab "${sourceSheetName}" not found.`);
        return null;
    }
    const sourceData = sourceSheet.getDataRange().getValues();
    if (sourceData.length < 1) return null;

    const sourceHeaders = sourceData.shift().map(h => h.toString().trim().toLowerCase());
    const sourceHeaderMap = new Map(sourceHeaders.map((h, i) => [h, i]));
    const getKeywordSourceIndex = () => CONFIG.POSSIBLE_KEYWORD_HEADERS.map(h => sourceHeaderMap.get(h)).find(i => i !== undefined);

    const targetSheet = spreadsheet.insertSheet(targetSheetName, 0);
    targetSheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);

    const targetData = sourceData.map(sourceRow => {
        const newRow = new Array(finalHeaders.length).fill('');
        const keywordIndex = getKeywordSourceIndex();
        if (keywordIndex !== undefined) newRow[0] = sourceRow[keywordIndex];
        finalHeaders.forEach((h, i) => {
            if (i === 0) return;
            const sourceIndex = sourceHeaderMap.get(h.toLowerCase());
            if (sourceIndex !== undefined) newRow[i] = sourceRow[sourceIndex];
        });
        return newRow;
    });

    if (targetData.length > 0) {
        targetSheet.getRange(2, 1, targetData.length, targetData[0].length).setValues(targetData);
    }

    spreadsheet.getSheets().forEach(s => {
        if (s.getName() !== targetSheetName) spreadsheet.deleteSheet(s);
    });
    return targetSheet;
}

function hideQaColumns(sheet) {
    const headersToHide = ["corrected_keyword", "pluralised_keyword", "cleanup_keyword", "regional_keyword", "keyword_sources", "is_subjective", "is_sensical", "Current Ranking(if applicable)"];
    const headers = getSheetHeaders(sheet);
    headers.forEach((header, i) => {
        if (headersToHide.includes(header)) {
            try { sheet.hideColumns(i + 1); }
            catch (e) { /* Ignore errors */ }
        }
    });
}

function enrichAndFinalizeSheet(sheet, svMap, stagingMap, keywordGenMap) {
    const headers = getSheetHeaders(sheet);
    const colIndices = {
        keyword: headers.indexOf(CONFIG.HEADER_NAMES.KEYWORD),
        sv: headers.indexOf(CONFIG.HEADER_NAMES.SEARCH_VOLUME),
        url: headers.indexOf(CONFIG.HEADER_NAMES.URL),
        title: headers.indexOf(CONFIG.HEADER_NAMES.TITLE),
        pageName: headers.indexOf(CONFIG.HEADER_NAMES.PAGE_NAME),
        productIds: headers.indexOf(CONFIG.HEADER_NAMES.PRODUCT_IDS),
        pc: headers.indexOf(CONFIG.HEADER_NAMES.PC)
    };

    if (colIndices.keyword === -1) return;
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    const data = dataRange.getValues();

    data.forEach(row => {
        const searchKey = row[colIndices.keyword]?.toString().toLowerCase().trim();
        if (!searchKey) return;
        
        let svPopulated = false;

        if (keywordGenMap?.has(searchKey)) {
            const genData = keywordGenMap.get(searchKey);
            const productIdsString = genData.productIds;

            if (colIndices.productIds !== -1 && productIdsString) {
                row[colIndices.productIds] = productIdsString;
            }

            if (colIndices.pc !== -1) {
                let count = 0;
                if (productIdsString && productIdsString.length > 2) {
                    try {
                        // Count pairs of single quotes to get the number of IDs
                        count = (productIdsString.match(/'/g) || []).length / 2;
                    } catch (e) {
                        Logger.log(`Could not parse product_ids for key "${searchKey}": ${productIdsString}`);
                        count = 0;
                    }
                }
                row[colIndices.pc] = count;
            }

            if (colIndices.sv !== -1 && genData.sv) {
                 row[colIndices.sv] = parseInt(genData.sv, 10) || null;
                 svPopulated = true;
            }
        } else if (colIndices.pc !== -1) {
             row[colIndices.pc] = 0;
        }

        if (!svPopulated && svMap && colIndices.sv !== -1) {
            const sv = svMap.get(searchKey);
            if (sv !== undefined) row[colIndices.sv] = parseInt(sv, 10) || null;
        }

        if (stagingMap?.has(searchKey)) {
            const d = stagingMap.get(searchKey);
            if (colIndices.url !== -1 && d.url) row[colIndices.url] = d.url;
            if (colIndices.title !== -1 && d.title) row[colIndices.title] = d.title;
            if (colIndices.pageName !== -1 && d.page_name) row[colIndices.pageName] = d.page_name;
        }
    });
    dataRange.setValues(data);
    sortSheetByTraffic(sheet);
    applyPcFilter(sheet);
}

function createSearchVolumeMap(spreadsheet) {
    const svMap = new Map();
    for (const sheet of spreadsheet.getSheets()) {
        try {
            const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().toLowerCase().trim());
            const keywordCol = headers.findIndex(h => CONFIG.POSSIBLE_KEYWORD_HEADERS.includes(h));
            const svCol = headers.findIndex(h => CONFIG.POSSIBLE_SV_HEADERS.includes(h));
            if (keywordCol !== -1 && svCol !== -1) {
                const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
                data.forEach(row => {
                    const key = row[keywordCol]?.toString().toLowerCase().trim();
                    if (key && row[svCol]) svMap.set(key, row[svCol]);
                });
                return svMap;
            }
        } catch (e) { continue; }
    }
    return svMap.size > 0 ? svMap : null;
}

function createStagingDataMap(spreadsheet) {
    const URL_HEADERS = ['url', 'urls', 'page url', 'page_url', 'link'];
    const TITLE_HEADERS = ['title'];
    const PAGENAME_HEADERS = ['name', 'page_name', 'pagename'];
    const stagingMap = new Map();
    for (const sheet of spreadsheet.getSheets()) {
        try {
            const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().toLowerCase().trim());
            const keywordCol = headers.findIndex(h => CONFIG.POSSIBLE_KEYWORD_HEADERS.includes(h));
            if (keywordCol === -1) continue;

            const urlCol = headers.findIndex(h => URL_HEADERS.includes(h));
            const titleCol = headers.findIndex(h => TITLE_HEADERS.includes(h));
            const pageNameCol = headers.findIndex(h => PAGENAME_HEADERS.includes(h));
            const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

            data.forEach(row => {
                const key = row[keywordCol]?.toString().toLowerCase().trim();
                if (key) {
                    const d = stagingMap.get(key) || {};
                    if (urlCol !== -1 && row[urlCol] && !d.url) d.url = row[urlCol];
                    if (titleCol !== -1 && row[titleCol] && !d.title) d.title = row[titleCol];
                    if (pageNameCol !== -1 && row[pageNameCol] && !d.page_name) d.page_name = row[pageNameCol];
                    if (Object.keys(d).length > 0) stagingMap.set(key, d);
                }
            });
        } catch (e) { continue; }
    }
    return stagingMap.size > 0 ? stagingMap : null;
}

function createKeywordGenMap(spreadsheet) {
    // --- NO CHANGES ABOVE THIS LINE IN THIS FUNCTION ---
    const PRODUCT_IDS_HEADERS = ['product_ids', 'product ids'];
    // CHANGED: Define headers we will look for
    const PRIMARY_KEYWORD_HEADER = 'cleanup_keyword';
    const FALLBACK_KEYWORD_HEADER = 'original_keyword';

    const keywordGenMap = new Map();
    for (const sheet of spreadsheet.getSheets()) {
        try {
            const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().toLowerCase().trim());
            
            // CHANGED: Find the index of both the primary and fallback keyword columns
            const primaryKeywordCol = headers.indexOf(PRIMARY_KEYWORD_HEADER);
            const fallbackKeywordCol = headers.indexOf(FALLBACK_KEYWORD_HEADER);
            
            // CHANGED: Continue to the next sheet if NEITHER keyword column is found.
            if (primaryKeywordCol === -1 && fallbackKeywordCol === -1) continue;

            const productIdsCol = headers.findIndex(h => PRODUCT_IDS_HEADERS.includes(h));
            const svCol = headers.findIndex(h => CONFIG.POSSIBLE_SV_HEADERS.includes(h));
            
            // This sheet must have product IDs to be useful for this function
            if (productIdsCol === -1) continue;

            const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

            data.forEach(row => {
                let key = null;

                // CHANGED: Implement the fallback logic.
                // 1. Try to get the key from the primary 'cleanup_keyword' column.
                if (primaryKeywordCol !== -1 && row[primaryKeywordCol] && row[primaryKeywordCol].toString().trim() !== '') {
                    key = row[primaryKeywordCol].toString().toLowerCase().trim();
                } 
                // 2. If that fails, fall back to the 'original_keyword' column.
                else if (fallbackKeywordCol !== -1 && row[fallbackKeywordCol] && row[fallbackKeywordCol].toString().trim() !== '') {
                    key = row[fallbackKeywordCol].toString().toLowerCase().trim();
                }
                
                if (key) { // If we successfully found a key from either column
                     const entry = {
                        productIds: row[productIdsCol] || null,
                        sv: (svCol !== -1 && row[svCol]) ? row[svCol] : null
                    };
                    keywordGenMap.set(key, entry);
                }
            });

            if (keywordGenMap.size > 0) return keywordGenMap;

        } catch (e) {
            Logger.log(`Error processing sheet in Keyword Gen file: ${e.toString()}`);
            continue;
        }
    }
    return keywordGenMap.size > 0 ? keywordGenMap : null;
}

function sortSheetByTraffic(sheet) {
    if (sheet.getLastRow() < 2) return;
    const headers = getSheetHeaders(sheet);
    const sortColIndex = headers.indexOf(CONFIG.HEADER_NAMES.SEARCH_VOLUME);
    if (sortColIndex === -1) return;
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort({ column: sortColIndex + 1, ascending: false });
}

/**
 * Checks if a keyword contains only allowed "normal text" characters.
 * This version is updated to allow periods (.) and all international/accented letters.
 * Allowed: letters (any language), numbers, whitespace, single quotes, double quotes, hyphens, periods.
 * @param {string} keyword The keyword to validate.
 * @returns {boolean} True if valid, false otherwise.
 */
function isValidKeyword(keyword) {
    if (typeof keyword !== 'string' || !keyword.trim()) {
        return false;
    }
    // UPDATED REGEX:
    // \p{L} -> Matches any Unicode letter (e.g., a, z, â, é)
    // \p{N} -> Matches any Unicode number (e.g., 0, 9)
    // \.    -> Matches a literal period (the backslash is to "escape" it)
    // u     -> The "unicode" flag, required for \p{L} and \p{N} to work.
    const allowedPattern = /^[\p{L}\p{N}\s'"\.-]*$/u;
    return allowedPattern.test(keyword);
}


/**
 * Validates a URL based on specific formatting rules.
 * @param {string} url The URL to validate.
 * @returns {{isValid: boolean, reason: string}} An object indicating validity and a reason for failure.
 */
function validateUrl(url) {
    if (typeof url !== 'string' || !url.trim()) {
        return { isValid: false, reason: "URL is empty or not a string." };
    }
    
    const trimmedUrl = url.trim();

    if (!trimmedUrl.startsWith("https://")) {
        return { isValid: false, reason: "Does not start with 'https://'." };
    }

    if (trimmedUrl.includes('--')) {
        return { isValid: false, reason: "Contains consecutive hyphens ('--')." };
    }
    
    // This regex looks for any character that is NOT a Unicode letter, a number, or one of the allowed symbols.
    // The 'u' flag is crucial for it to correctly handle international characters (\p{L}).
    const invalidCharPattern = /[^\p{L}\p{N}_.:\/-]/u; 
    const match = trimmedUrl.match(invalidCharPattern);
    
    if (match) {
        return { isValid: false, reason: `Contains invalid character: '${match[0]}'.` };
    }

    return { isValid: true, reason: "Valid" };
}

// =================================================================
// SECTION 4: DATA VALIDATION TOOLS (NEW SECTION)
// =================================================================

/**
 * Generic handler to validate data in the user's selected range.
 * @param {function} validationFunction The specific validation function to run (e.g., isValidKeyword or validateUrl).
 * @param {string} outputSheetName The name of the sheet to output errors to (e.g., "faulty_keywords").
 * @param {string} entityName The singular name of the item being validated (e.g., "Keyword").
 * @param {string} entityNamePlural The plural name of the item being validated (e.g., "Keywords").
 */
function runValidationOnSelection(validationFunction, outputSheetName, entityName, entityNamePlural) {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeRange = ss.getActiveRange();

    if (!activeRange) {
        ui.alert('Please select a range of cells to validate first.');
        return;
    }

    const values = activeRange.getValues();
    const faultyItems = [];

    values.forEach((row, rowIndex) => {
        row.forEach((cellValue, colIndex) => {
            if (cellValue && typeof cellValue === 'string' && cellValue.trim() !== '') {
                const result = validationFunction(cellValue);
                // The handler adapts to the two different return types of our validation functions
                const isInvalid = (typeof result === 'boolean' && !result) || (typeof result === 'object' && !result.isValid);

                if (isInvalid) {
                    const reason = typeof result === 'object' ? result.reason : 'Contains disallowed characters.';
                    const cellNotation = activeRange.offset(rowIndex, colIndex, 1, 1).getA1Notation();
                    faultyItems.push([cellValue, reason, cellNotation]);
                }
            }
        });
    });

    if (faultyItems.length === 0) {
        ss.toast(`Validation Complete: All ${values.length * values[0].length} selected ${entityNamePlural.toLowerCase()} are valid.`, CONFIG.TOAST_TITLE, 10);
    } else {
        let sheet = ss.getSheetByName(outputSheetName);
        if (!sheet) {
            sheet = ss.insertSheet(outputSheetName);
        } else {
            sheet.clear();
        }

        const headers = [`Invalid ${entityName}`, 'Reason', 'Original Cell'];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
        sheet.getRange(2, 1, faultyItems.length, faultyItems[0].length).setValues(faultyItems);
        sheet.autoResizeColumns(1, headers.length);
        sheet.activate();
        ui.alert(`Validation Complete`, `${faultyItems.length} invalid ${entityNamePlural.toLowerCase()} found. See the "${outputSheetName}" sheet for details.`, ui.ButtonSet.OK);
    }
}

/**
 * Menu function to trigger keyword validation on the selected range.
 */
function validateSelectedKeywords() {
    runValidationOnSelection(isValidKeyword, 'faulty_keywords', 'Keyword', 'Keywords');
}

/**
 * Menu function to trigger URL validation on the selected range.
 */
function validateSelectedUrls() {
    runValidationOnSelection(validateUrl, 'faulty_urls', 'URL', 'URLs');
}
