const { Builder, Browser, By, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const path = require('path');
const fs = require('fs');
const { ServiceBuilder } = require('selenium-webdriver/chrome');

// Load configuration
const config = require('./config.json');

// Configuration from config.json
const DOWNLOAD_URL = config.roster.downloadUrl;
const ROSTER_FOLDER = path.join(__dirname, config.roster.folderName);
const HISTORY_FOLDER = path.join(ROSTER_FOLDER, config.roster.historyFolderName);
const OUTPUT_FILENAME = config.roster.outputFileName;
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
        
        if (now - stats.mtimeMs > retentionMs) {
            fs.unlinkSync(filePath);
            console.log(`Deleted old file: ${file}`);
            deletedCount++;
        }
    }
    
    console.log(`Cleanup complete. Deleted ${deletedCount} old file(s).`);
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
}

// Run the download
downloadExcelFile()
    .then(() => {
        console.log('Script completed successfully');
        process.exit(0);
    })
    .catch((error) => {
        console.error('Script failed:', error);
        process.exit(1);
    });
