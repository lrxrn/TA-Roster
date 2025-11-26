const { Builder, Browser, By, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const path = require('path');
const fs = require('fs');
const { ServiceBuilder } = require('selenium-webdriver/chrome');
const XLSX = require('xlsx');

// Parse command line arguments
const args = process.argv.slice(2);
const SKIP_DOWNLOAD = args.includes('--skip-download') || args.includes('-s');

// Load configuration
const config = require('./config.json');

// Configuration from config.json
const DOWNLOAD_URL = config.roster.downloadUrl;
const ROSTER_FOLDER = path.join(__dirname, config.roster.folderName);
const HISTORY_FOLDER = path.join(ROSTER_FOLDER, config.roster.historyFolderName);
const OUTPUT_FILENAME = config.roster.outputFileName;
const JSON_OUTPUT_FILENAME = 'Roster.json';
const RETENTION_DAYS = config.cleanup.retentionDays;
const DOWNLOAD_TIMEOUT = config.download.timeoutMs;
const HEADLESS_MODE = config.download.headless;

// Get current date/time in DD-MM-YYYY-HH-MM format
function getDateTimeString() {
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = now.getFullYear();
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    return `_${day}-${month}-${year}-${hours}-${minutes}`;
}

// Create required folders if they don't exist
function ensureFolders() {
    if (!fs.existsSync(ROSTER_FOLDER)) {
        fs.mkdirSync(ROSTER_FOLDER, { recursive: true });
        console.log('Created roster folder:', ROSTER_FOLDER);
    }
    if (!fs.existsSync(HISTORY_FOLDER)) {
        fs.mkdirSync(HISTORY_FOLDER, { recursive: true });
        console.log('Created history folder:', HISTORY_FOLDER);
    }
}

// Clean up old files from history folder (older than retention period)
function cleanupOldFiles() {
    console.log(`Cleaning up files older than ${RETENTION_DAYS} days...`);
    
    if (!fs.existsSync(HISTORY_FOLDER)) {
        console.log('History folder does not exist, skipping cleanup');
        return;
    }
    
    const now = Date.now();
    const retentionMs = RETENTION_DAYS * 24 * 60 * 60 * 1000;
    const files = fs.readdirSync(HISTORY_FOLDER);
    
    let deletedCount = 0;
    for (const file of files) {
        const filePath = path.join(HISTORY_FOLDER, file);
        const stats = fs.statSync(filePath);
        
        // Clean up both Excel files and JSON files
        const isTargetFile = file.endsWith('.xlsx') || file.endsWith('.xls') || file.endsWith('.json');
        
        if (isTargetFile && now - stats.mtimeMs > retentionMs) {
            fs.unlinkSync(filePath);
            console.log(`Deleted old file: ${file}`);
            deletedCount++;
        }
    }
    
    console.log(`Cleanup complete. Deleted ${deletedCount} old file(s).`);
}

// Check if a sheet name is a roster sheet (date format ddmmyyyy or name "TA")
function isRosterSheet(sheetName) {
    // Check if sheet name is "TA"
    if (sheetName.toUpperCase() === 'TA') {
        return true;
    }
    
    // Check if sheet name matches ddmmyyyy format (8 digits)
    const datePattern = /^\d{8}$/;
    return datePattern.test(sheetName);
}

// Convert Excel date serial number to JavaScript Date
function excelDateToJSDate(serial) {
    // Excel dates start from 1900-01-01 (but Excel incorrectly treats 1900 as a leap year)
    const excelEpoch = new Date(1899, 11, 30);
    const days = Math.floor(serial);
    const milliseconds = Math.round((serial - days) * 86400000);
    return new Date(excelEpoch.getTime() + days * 86400000 + milliseconds);
}

// Get Monday of the week from a date
function getMondayOfWeek(date) {
    const d = new Date(date);
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1); // Adjust when day is Sunday
    const monday = new Date(d.setDate(diff));
    monday.setHours(0, 0, 0, 0);
    return monday;
}

// Format date as DD-MM-YYYY
function formatDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}-${month}-${year}`;
}

// Get day name from date
function getDayName(date) {
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    return days[date.getDay()];
}

// Check if cell has yellow background (FFFF00)
function isYellowCell(cell) {
    if (!cell || !cell.s || !cell.s.fgColor) {
        return false;
    }
    
    const color = cell.s.fgColor.rgb;
    if (!color) {
        return false;
    }
    
    // Check for yellow color (FFFF00 or similar)
    const normalizedColor = color.toUpperCase();
    return normalizedColor === 'FFFFFF00' || normalizedColor === 'FFFF00';
}

// Parse time string to {start, end} object
function parseTimeRange(timeStr) {
    if (!timeStr || typeof timeStr !== 'string') {
        return null;
    }
    
    const parts = timeStr.split('-').map(t => t.trim());
    if (parts.length !== 2) {
        return null;
    }
    
    // Convert 0815 to 08:15
    const formatTime = (t) => {
        const cleaned = t.replace(/[^0-9]/g, '');
        if (cleaned.length === 4) {
            return `${cleaned.substring(0, 2)}:${cleaned.substring(2, 4)}`;
        }
        return t;
    };
    
    return {
        start: formatTime(parts[0]),
        end: formatTime(parts[1])
    };
}

// Parse a single day's roster from the worksheet
function parseDayRoster(worksheet, dateRow) {
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    
    // Read the date cell (Excel date number)
    const dateCell = worksheet[XLSX.utils.encode_cell({ r: dateRow, c: 0 })];
    if (!dateCell || !dateCell.v || typeof dateCell.v !== 'number') {
        return null;
    }
    
    // Convert Excel date to JS Date and get day name
    const date = excelDateToJSDate(dateCell.v);
    const dayName = getDayName(date);
    const dateFormatted = formatDate(date);
    
    // Header row is immediately after date row
    const headerRow = dateRow + 1;
    
    // Verify this is a header row (should have "Shift" in column A)
    const headerCheck = worksheet[XLSX.utils.encode_cell({ r: headerRow, c: 0 })];
    if (!headerCheck || headerCheck.v.toString().toLowerCase() !== 'shift') {
        return null;
    }
    
    // Read headers (starting from column 2, which is column C - the first shift type)
    const headers = [];
    for (let col = 2; col <= range.e.c; col++) {
        const headerCell = worksheet[XLSX.utils.encode_cell({ r: headerRow, c: col })];
        if (headerCell && headerCell.v) {
            headers.push({ col, name: headerCell.v.toString() });
        }
    }
    
    if (headers.length === 0) {
        return null;
    }
    
    // Parse shifts (starting from row after header)
    const shifts = [];
    let currentRow = headerRow + 1;
    
    while (currentRow <= range.e.r) {
        const shiftNameCell = worksheet[XLSX.utils.encode_cell({ r: currentRow, c: 0 })];
        
        // Stop if we hit an empty row or a cell that looks like a date number (next day)
        if (!shiftNameCell || !shiftNameCell.v) {
            // Check next few rows for another date
            let foundNextDate = false;
            for (let i = 1; i <= 3; i++) {
                const nextCell = worksheet[XLSX.utils.encode_cell({ r: currentRow + i, c: 0 })];
                if (nextCell && nextCell.v && typeof nextCell.v === 'number') {
                    foundNextDate = true;
                    break;
                }
            }
            if (foundNextDate || currentRow - headerRow > 10) {
                break;
            }
            currentRow++;
            continue;
        }
        
        const shiftName = shiftNameCell.v.toString().trim();
        
        // If we hit another date number, stop
        if (typeof shiftNameCell.v === 'number') {
            break;
        }
        
        // Read shift time from column B
        const timeCell = worksheet[XLSX.utils.encode_cell({ r: currentRow, c: 1 })];
        const timeRange = timeCell && timeCell.v ? parseTimeRange(timeCell.v.toString()) : null;
        
        if (!timeRange) {
            currentRow++;
            continue;
        }
        
        // Read assignments for this shift
        const assignments = [];
        
        for (const header of headers) {
            const cellAddress = XLSX.utils.encode_cell({ r: currentRow, c: header.col });
            const cell = worksheet[cellAddress];
            
            if (cell && cell.v) {
                const personName = cell.v.toString().trim();
                
                // Skip if name is empty
                if (personName) {
                    // Check if cell has yellow background
                    const isAssigned = isYellowCell(cell);
                    
                    assignments.push({
                        type: header.name,
                        person: personName,
                        assignType: isAssigned ? 'assigned' : 'requested'
                    });
                }
            }
        }
        
        // Only add shift if it has assignments
        if (assignments.length > 0) {
            shifts.push({
                name: shiftName,
                time: timeRange,
                assignments
            });
        }
        
        currentRow++;
    }
    
    return shifts.length > 0 ? { day: dayName, date: dateFormatted, shifts } : null;
}

// Wait for a file to be downloaded (check for new or modified files in the folder)
async function waitForDownload(downloadPath, startTime, timeout = DOWNLOAD_TIMEOUT) {
    const timeoutEnd = Date.now() + timeout;
    
    while (Date.now() < timeoutEnd) {
        const currentFiles = fs.readdirSync(downloadPath);
        
        // Check for any downloading files (.crdownload)
        const downloadingFiles = currentFiles.filter(f => f.endsWith('.crdownload'));
        if (downloadingFiles.length > 0) {
            console.log('Download in progress:', downloadingFiles[0]);
            await new Promise(resolve => setTimeout(resolve, 1000));
            continue;
        }
        
        // Look for Excel files that were modified after we started (ignore main roster file)
        const excelFiles = currentFiles.filter(f => 
            (f.endsWith('.xlsx') || f.endsWith('.xls')) && 
            !f.includes('.crdownload') && 
            !f.includes('.tmp') &&
            f !== OUTPUT_FILENAME
        );
        
        for (const file of excelFiles) {
            const filePath = path.join(downloadPath, file);
            const stats = fs.statSync(filePath);
            
            // Check if file was modified after we started the download
            if (stats.mtimeMs >= startTime) {
                console.log(`Found recently modified file: ${file} (modified at ${stats.mtime})`);
                // Wait a moment to ensure file is fully written
                await new Promise(resolve => setTimeout(resolve, 2000));
                return file;
            }
        }
        
        // Wait a bit before checking again
        await new Promise(resolve => setTimeout(resolve, 1000));
    }
    
    throw new Error('Download timeout - file not downloaded within expected time');
}

// Move file to history with date/time suffix and copy to main roster file
function processDownloadedFile(downloadPath, filename) {
    const oldPath = path.join(downloadPath, filename);
    const ext = path.extname(filename);
    const baseName = path.basename(filename, ext);
    const dateTimeSuffix = getDateTimeString();
    const historyFilename = `${baseName}${dateTimeSuffix}${ext}`;
    const historyPath = path.join(HISTORY_FOLDER, historyFilename);
    const mainRosterPath = path.join(ROSTER_FOLDER, OUTPUT_FILENAME);
    
    // Move to history folder with date/time suffix
    fs.renameSync(oldPath, historyPath);
    console.log(`Moved to history: ${historyFilename}`);
    
    // Copy the file to main roster location
    fs.copyFileSync(historyPath, mainRosterPath);
    console.log(`Updated main roster file: ${OUTPUT_FILENAME}`);
    
    return { historyFile: historyFilename, mainFile: OUTPUT_FILENAME };
}

async function downloadExcelFile() {
    // Ensure folders exist
    ensureFolders();
    
    // Clean up old files first
    cleanupOldFiles();
    
    // Record the start time to detect new/modified files
    const startTime = Date.now();
    console.log('Starting download process at:', new Date(startTime).toISOString());
    
    // Configure Chrome options
    const chromeOptions = new chrome.Options();
    
    // Set download directory preferences
    chromeOptions.setUserPreferences({
        'download.default_directory': ROSTER_FOLDER,
        'download.prompt_for_download': false,
        'download.directory_upgrade': true,
        'safebrowsing.enabled': false,
        'safebrowsing.disable_download_protection': true,
        'plugins.always_open_pdf_externally': true
    });
    
    // Run in headless mode based on config (required for GitHub Actions)
    if (HEADLESS_MODE) {
        chromeOptions.addArguments('--headless=new');
    }
    chromeOptions.addArguments('--disable-gpu');
    chromeOptions.addArguments('--no-sandbox');
    chromeOptions.addArguments('--disable-dev-shm-usage');
    chromeOptions.addArguments('--disable-software-rasterizer');
    chromeOptions.addArguments('--window-size=1920,1080');
    
    let driver;
    
    try {
        console.log('Starting Chrome browser...');
        
        // Use Chrome's built-in driver service
        const service = new ServiceBuilder();
        
        driver = await new Builder()
            .forBrowser(Browser.CHROME)
            .setChromeOptions(chromeOptions)
            .setChromeService(service)
            .build();
        
        console.log('Browser started successfully!');
        
        console.log('Navigating to SharePoint download link...');
        await driver.get(DOWNLOAD_URL);
        
        // Wait for the page to load and download to start
        console.log('Waiting for download to start...');
        
        // Give some time for download to begin
        await driver.sleep(3000);
        
        // Wait for the file to be downloaded
        console.log('Waiting for file download to complete...');
        const downloadedFile = await waitForDownload(ROSTER_FOLDER, startTime, DOWNLOAD_TIMEOUT);
        console.log(`Downloaded: ${downloadedFile}`);
        
        // Process the downloaded file (move to history, update main roster)
        const result = processDownloadedFile(ROSTER_FOLDER, downloadedFile);
        
        console.log('Download completed successfully!');
        console.log(`History file: ${path.join(HISTORY_FOLDER, result.historyFile)}`);
        console.log(`Main roster: ${path.join(ROSTER_FOLDER, result.mainFile)}`);
        
    } catch (error) {
        console.error('Error during download:', error.message);
        throw error;
    } finally {
        if (driver) {
            console.log('Closing browser...');
            await driver.quit();
        }
    }

    return path.join(ROSTER_FOLDER, OUTPUT_FILENAME);
}

// Get the latest Excel file from the roster folder
function getLatestRosterFile() {
    if (!fs.existsSync(ROSTER_FOLDER)) {
        throw new Error(`Roster folder does not exist: ${ROSTER_FOLDER}`);
    }
    
    const files = fs.readdirSync(ROSTER_FOLDER);
    const excelFiles = files.filter(f => 
        (f.endsWith('.xlsx') || f.endsWith('.xls')) && 
        !f.startsWith('~') && 
        !f.includes('.tmp')
    );
    
    if (excelFiles.length === 0) {
        throw new Error(`No Excel files found in roster folder: ${ROSTER_FOLDER}`);
    }
    
    // Get the most recently modified file
    let latestFile = null;
    let latestTime = 0;
    
    for (const file of excelFiles) {
        const filePath = path.join(ROSTER_FOLDER, file);
        const stats = fs.statSync(filePath);
        
        if (stats.mtimeMs > latestTime) {
            latestTime = stats.mtimeMs;
            latestFile = filePath;
        }
    }
    
    if (!latestFile) {
        throw new Error('No valid Excel file found in roster folder');
    }
    
    console.log(`Using latest roster file: ${path.basename(latestFile)}`);
    console.log(`Last modified: ${new Date(latestTime).toISOString()}`);
    
    return latestFile;
}

async function parseExceltoJson(filePath) {
    console.log('Parsing Excel to JSON...');
    console.log('Reading file:', filePath);
    
    try {
        // Read the Excel file
        const workbook = XLSX.readFile(filePath, { cellStyles: true });
        const allRosterData = [];
        
        // Process each sheet
        for (const sheetName of workbook.SheetNames) {
            console.log(`Checking sheet: ${sheetName}`);
            
            // Check if this is a roster sheet
            if (!isRosterSheet(sheetName)) {
                console.log(`Skipping non-roster sheet: ${sheetName}`);
                continue;
            }
            
            console.log(`Processing roster sheet: ${sheetName}`);
            const worksheet = workbook.Sheets[sheetName];
            
            // Get week start date from cell A3
            const a3Cell = worksheet['A3'];
            if (!a3Cell || !a3Cell.v) {
                console.log(`Warning: Cell A3 is empty in sheet ${sheetName}, skipping`);
                continue;
            }
            
            // Convert Excel date to JavaScript Date
            let weekDate;
            if (typeof a3Cell.v === 'number') {
                weekDate = excelDateToJSDate(a3Cell.v);
            } else {
                weekDate = new Date(a3Cell.v);
            }
            
            const mondayOfWeek = getMondayOfWeek(weekDate);
            console.log(`Week start (Monday): ${formatDate(mondayOfWeek)}`);
            
            // Parse all days in the roster
            const range = XLSX.utils.decode_range(worksheet['!ref']);
            const dayData = [];
            
            // Find all date rows (cells in column A that are numbers and likely dates)
            let currentRow = 0;
            
            while (currentRow <= range.e.r) {
                const cellAddress = XLSX.utils.encode_cell({ r: currentRow, c: 0 });
                const cell = worksheet[cellAddress];
                
                // Check if this cell contains a date number (Excel date)
                if (cell && cell.v && typeof cell.v === 'number') {
                    // Verify it's likely a date (between 40000 and 50000 represents dates around 2009-2036)
                    if (cell.v > 40000 && cell.v < 50000) {
                        console.log(`Found date at row ${currentRow}: ${cell.v}`);
                        const dayRoster = parseDayRoster(worksheet, currentRow);
                        
                        if (dayRoster) {
                            dayData.push(dayRoster);
                            console.log(`Parsed ${dayRoster.day} with ${dayRoster.shifts.length} shifts`);
                        }
                    }
                }
                
                currentRow++;
            }
            
            // Add this week's data if we found any days
            if (dayData.length > 0) {
                allRosterData.push({
                    week: formatDate(mondayOfWeek),
                    parsedTimestamp: new Date().toISOString(),
                    data: dayData
                });
                console.log(`Parsed ${dayData.length} days from sheet ${sheetName}`);
            }
        }
        
        if (allRosterData.length === 0) {
            console.log('No roster data found in the Excel file');
            return;
        }
        
        // Write JSON to main roster folder
        const mainJsonPath = path.join(ROSTER_FOLDER, JSON_OUTPUT_FILENAME);
        fs.writeFileSync(mainJsonPath, JSON.stringify(allRosterData, null, 2));
        console.log(`Written JSON to: ${mainJsonPath}`);
        
        // Write JSON to history folder with timestamp
        const dateTimeSuffix = getDateTimeString();
        const historyJsonFilename = `Roster${dateTimeSuffix}.json`;
        const historyJsonPath = path.join(HISTORY_FOLDER, historyJsonFilename);
        fs.writeFileSync(historyJsonPath, JSON.stringify(allRosterData, null, 2));
        console.log(`Written JSON to history: ${historyJsonPath}`);
        
        console.log(`Successfully parsed ${allRosterData.length} week(s) from roster`);
        
    } catch (error) {
        console.error('Error parsing Excel file:', error.message);
        throw error;
    }
}

// Run the download
if (SKIP_DOWNLOAD) {
    console.log('Skipping download (--skip-download flag detected)');
    console.log('Looking for existing roster file...\n');
    
    try {
        const latestFile = getLatestRosterFile();
        
        console.log('\n--- Starting Excel Parsing ---\n');
        parseExceltoJson(latestFile)
            .then(() => {
                console.log('\nScript completed successfully');
                process.exit(0);
            })
            .catch((error) => {
                console.error('Parsing failed:', error.message);
                process.exit(1);
            });
    } catch (error) {
        console.error('Error: Cannot skip download -', error.message);
        console.error('Please run without --skip-download flag to download a new roster file.');
        process.exit(1);
    }
} else {
    downloadExcelFile()
        .then(async (filePath) => {
            console.log('\n--- Starting Excel Parsing ---\n');
            await parseExceltoJson(filePath);
            console.log('\nScript completed successfully');
            process.exit(0);
        })
        .catch((error) => {
            console.error('Script failed:', error);
            process.exit(1);
        });
}
