/** Google Drive cache helpers: create/read/write/list/delete JSON files */

function getOrCreateCacheFolder() {
  const folders = DriveApp.getFoldersByName(CACHE_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(CACHE_FOLDER_NAME);
}

function writeJSONToDrive(ticker, exchange, jsonData) {
  const folder = getOrCreateCacheFolder();
  const fileName = `${ticker}_${exchange}.json`;
  const jsonString = JSON.stringify(jsonData, null, 2);
  const file = folder.createFile(fileName, jsonString, MimeType.PLAIN_TEXT);
  logInfo(`Created ${fileName} (${(jsonString.length/1024).toFixed(2)} KB)`);
  return { success: true, fileId: file.getId(), name: fileName, size: jsonString.length };
}

function readJSONFromDrive(ticker, exchange) {
  const folder = getOrCreateCacheFolder();
  const fileName = `${ticker}_${exchange}.json`;
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) return null;
  const file = files.next();
  const text = file.getBlob().getDataAsString();
  try {
    return JSON.parse(text);
  } catch (e) {
    throw new Error(`Invalid JSON in ${fileName}: ${e.message}`);
  }
}

function fileExists(ticker, exchange) {
  const folder = getOrCreateCacheFolder();
  const files = folder.getFilesByName(`${ticker}_${exchange}.json`);
  return files.hasNext();
}

function overwriteJSONInDrive(ticker, exchange, jsonData) {
  const folder = getOrCreateCacheFolder();
  const fileName = `${ticker}_${exchange}.json`;
  const files = folder.getFilesByName(fileName);
  const jsonString = JSON.stringify(jsonData, null, 2);

  if (files.hasNext()) {
    const file = files.next();
    file.setContent(jsonString);
    logInfo(`Updated ${fileName} (${(jsonString.length/1024).toFixed(2)} KB)`);
    return { success: true, name: fileName, size: jsonString.length };
  } else {
    const file = folder.createFile(fileName, jsonString, MimeType.PLAIN_TEXT);
    logInfo(`Created ${fileName} (${(jsonString.length/1024).toFixed(2)} KB)`);
    return { success: true, name: fileName, size: jsonString.length };
  }
}

function deleteJSONFromDrive(ticker, exchange) {
  const folder = getOrCreateCacheFolder();
  const fileName = `${ticker}_${exchange}.json`;
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) return false;
  const file = files.next();
  file.setTrashed(true);
  logWarn(`Deleted ${fileName}`);
  return true;
}

function listCachedFiles() {
  const folder = getOrCreateCacheFolder();
  const files = folder.getFiles();
  const list = [];
  while (files.hasNext()) {
    const f = files.next();
    list.push({
      name: f.getName(),
      sizeKB: +(f.getSize()/1024).toFixed(2),
      lastUpdated: f.getLastUpdated(),
      url: f.getUrl()
    });
  }
  logInfo(`Cached files: ${list.length}`);
  return list;
}

/** Quick sanity test */
function testFolderAccess() {
  const folder = getOrCreateCacheFolder();
  logInfo(`Folder: ${folder.getName()} — ${folder.getUrl()}`);
}
