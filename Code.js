// ******************************************************************************************
// ****************************** CODE.GS (BACKEND) *****************************************
// Version 1.15.0 - v272: Tooltip hover items, fix editor save, vehicle check VLI, SAC ISP content
// ******************************************************************************************

// --- CONFIGURATION ---
const SCRIPT_PROP = PropertiesService.getScriptProperties();
// Version code utilisée pour invalider le snapshot cache lors d'un déploiement
const CODE_VERSION = 'v273';
const BOOTSTRAP_SNAPSHOT_KEY = "BOOTSTRAP_SNAPSHOT_V1";
const PHOTO_PRESENCE_KEY = "PHOTO_PRESENCE_JSON";
const SHEET_NAMES = {
  INVENTORY: "Inventaire",
  HISTORY: "Historique",
  CONFIG: "Config",
  FORMS: "Structure_Forms"
};

// Spreadsheet externe "Carte Prévisionnelle Opérationnelle" pour la liste des ISP / centres
const EXTERNAL_SS_ID = "1bwKVm5Bto9u8lPopGxbwA3PfStJeTd159_cO-eKV4FY";
const EXTERNAL_ISP_SHEET = "ISP";

// Centres virtuels (hors ISP) avec emails fixes
var VIRTUAL_CENTERS = {
  "Astreinte Départementale Médicale": "brice.dubrey@sdis66.fr;florian.bois@sdis66.fr",
  "Garde PSud": "brice.dubrey@sdis66.fr;florian.bois@sdis66.fr"
};

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.bagParam = e.parameter.bag || null;
  return template.evaluate()
      .setTitle('Vérifications Matériel')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

// --- BOOTSTRAP (data + photos + mileages) with chunked cache ---
// PropertiesService: 9KB/value limit → chunk across multiple keys
// CacheService: 100KB/value limit → chunk if needed, 300s TTL

var CHUNK_SIZE_PROP = 8000;   // ~8KB per property chunk (safe under 9KB)
var CHUNK_SIZE_CACHE = 90000; // ~90KB per cache chunk (safe under 100KB)
var CACHE_TTL = 300;          // 5 minutes

function setChunkedProp_(baseKey, jsonStr) {
  try {
    var chunks = [];
    for (var i = 0; i < jsonStr.length; i += CHUNK_SIZE_PROP) {
      chunks.push(jsonStr.substring(i, i + CHUNK_SIZE_PROP));
    }
    // Delete old chunks first (up to 50 possible old ones)
    for (var c = 0; c < 50; c++) {
      var k = baseKey + '_' + c;
      if (!SCRIPT_PROP.getProperty(k)) break;
      SCRIPT_PROP.deleteProperty(k);
    }
    SCRIPT_PROP.setProperty(baseKey + '_N', String(chunks.length));
    for (var c = 0; c < chunks.length; c++) {
      SCRIPT_PROP.setProperty(baseKey + '_' + c, chunks[c]);
    }
  } catch(e) {
    Logger.log('setChunkedProp_ error (non-fatal): ' + e);
  }
}

function getChunkedProp_(baseKey) {
  try {
    var n = parseInt(SCRIPT_PROP.getProperty(baseKey + '_N') || '0');
    if (!n) return null;
    var str = '';
    for (var c = 0; c < n; c++) {
      var part = SCRIPT_PROP.getProperty(baseKey + '_' + c);
      if (part === null) return null;
      str += part;
    }
    return str;
  } catch(e) {
    Logger.log('getChunkedProp_ error: ' + e);
    return null;
  }
}

function deleteChunkedProp_(baseKey) {
  try {
    var n = parseInt(SCRIPT_PROP.getProperty(baseKey + '_N') || '0');
    for (var c = 0; c < Math.max(n, 50); c++) {
      var k = baseKey + '_' + c;
      if (!SCRIPT_PROP.getProperty(k) && c >= n) break;
      try { SCRIPT_PROP.deleteProperty(k); } catch(e) {}
    }
    try { SCRIPT_PROP.deleteProperty(baseKey + '_N'); } catch(e) {}
  } catch(e) {}
}

function setChunkedCache_(cache, baseKey, jsonStr, ttl) {
  try {
    if (jsonStr.length <= CHUNK_SIZE_CACHE) {
      cache.put(baseKey, jsonStr, ttl);
      cache.put(baseKey + '_N', '1', ttl);
    } else {
      var chunks = [];
      for (var i = 0; i < jsonStr.length; i += CHUNK_SIZE_CACHE) {
        chunks.push(jsonStr.substring(i, i + CHUNK_SIZE_CACHE));
      }
      cache.put(baseKey + '_N', String(chunks.length), ttl);
      for (var c = 0; c < chunks.length; c++) {
        cache.put(baseKey + '_C' + c, chunks[c], ttl);
      }
      // Also put a marker so single get fails gracefully
      cache.remove(baseKey);
    }
  } catch(e) {
    Logger.log('setChunkedCache_ error (non-fatal): ' + e);
  }
}

function getChunkedCache_(cache, baseKey) {
  try {
    var nStr = cache.get(baseKey + '_N');
    if (nStr === '1') {
      return cache.get(baseKey);
    }
    var n = parseInt(nStr || '0');
    if (!n) {
      // Fallback: try single value (old format)
      return cache.get(baseKey);
    }
    var str = '';
    for (var c = 0; c < n; c++) {
      var part = cache.get(baseKey + '_C' + c);
      if (part === null) return null; // partial miss
      str += part;
    }
    return str;
  } catch(e) {
    return null;
  }
}

function getBootstrapData() {
  const cache = CacheService.getScriptCache();
  // If code version changed since last snapshot, force rebuild
  try {
    const propVersion = SCRIPT_PROP.getProperty('CODE_VERSION') || '';
    if (propVersion !== CODE_VERSION) {
      cache.remove("BOOTSTRAP_V1");
      cache.remove("BOOTSTRAP_V1_N");
      deleteChunkedProp_(BOOTSTRAP_SNAPSHOT_KEY);
      // Also delete old non-chunked key if it exists
      try { SCRIPT_PROP.deleteProperty(BOOTSTRAP_SNAPSHOT_KEY); } catch(e) {}
      SCRIPT_PROP.setProperty('CODE_VERSION', CODE_VERSION);
      Logger.log('CODE_VERSION mismatch — forcing bootstrap snapshot rebuild.');
    }
  } catch(e) { Logger.log('Error checking CODE_VERSION: ' + e); }

  const cached = getChunkedCache_(cache, "BOOTSTRAP_V1");
  if (cached) return JSON.parse(cached);

  const snap = getChunkedProp_(BOOTSTRAP_SNAPSHOT_KEY);
  if (snap) {
    setChunkedCache_(cache, "BOOTSTRAP_V1", snap, CACHE_TTL);
    return JSON.parse(snap);
  }

  const payload = rebuildBootstrapSnapshot_();
  if (payload) {
    var jsonStr = JSON.stringify(payload);
    setChunkedCache_(cache, "BOOTSTRAP_V1", jsonStr, CACHE_TTL);
  }
  return payload;
}

function rebuildBootstrapSnapshot_() {
  const base = getData();
  if (!base || !base.success) return base;
  const payload = {
    success: true,
    data: base.data,
    photoPresence: getPhotoPresenceMap(),
    vliMileages: getAllVliMileages()
  };
  setChunkedProp_(BOOTSTRAP_SNAPSHOT_KEY, JSON.stringify(payload));
  return payload;
}

// --- INITIALISATION ---
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!ss.getSheetByName(SHEET_NAMES.INVENTORY)) {
    const s = ss.insertSheet(SHEET_NAMES.INVENTORY);
    s.appendRow(["Catégorie", "Nom", "Dernier_Controle", "Prochain_Controle", "Statut", "Dernier_Verificateur", "Prochain_Item_Nom", "Prochain_Item_Date", "Mail_Orange", "Mail_Red", "Etat", "Localisation", "Ordre", "SousType"]);
  } else {
    const s = ss.getSheetByName(SHEET_NAMES.INVENTORY);
    const lastCol = Math.max(1, s.getLastColumn());
    const header = s.getRange(1, 1, 1, lastCol).getValues()[0];
    if (header.length < 12 || header[11] !== "Localisation") s.getRange(1, 12).setValue("Localisation");
    if (header.length < 13 || header[12] !== "Ordre") s.getRange(1, 13).setValue("Ordre");
    if (header.length < 14 || header[13] !== "SousType") s.getRange(1, 14).setValue("SousType");
  }
  
  if (!ss.getSheetByName(SHEET_NAMES.HISTORY)) {
    const s = ss.insertSheet(SHEET_NAMES.HISTORY);
    s.appendRow(["Date", "Nom", "Verificateur", "Details_JSON"]);
  }
  
  if (!ss.getSheetByName(SHEET_NAMES.CONFIG)) {
    const s = ss.insertSheet(SHEET_NAMES.CONFIG);
    s.appendRow(["Categorie", "Frequence_Jours"]);
  }
  
  // Stockage des options globales par défaut
  if (!SCRIPT_PROP.getProperty("GLOBAL_OPTS")) {
    SCRIPT_PROP.setProperty("GLOBAL_OPTS", JSON.stringify({
      enableExpiry: true,
      enableQR: true,
      enableVerifier: true,
      enablePhotos: true
    }));
  }
}

// --- DATA FETCHING (Chargement des données) ---
function getData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    setup(); // S'assure que tout est prêt
    if (!SCRIPT_PROP.getProperty("INIT_V3_CLEANUP")) { cleanupCategories_(ss); SCRIPT_PROP.setProperty("INIT_V3_CLEANUP", "1"); }
    if (!SCRIPT_PROP.getProperty("INIT_V4_REMOVE_DEFAULTS")) { removeAutoDefaultBags_(ss); SCRIPT_PROP.setProperty("INIT_V4_REMOVE_DEFAULTS", "1"); }
    if (!SCRIPT_PROP.getProperty("INIT_V5_ORDER")) { initializeInventoryOrder_(ss); SCRIPT_PROP.setProperty("INIT_V5_ORDER", "1"); }
    if (!SCRIPT_PROP.getProperty("INIT_V6_VSSO")) { initVSSOContent_(); SCRIPT_PROP.setProperty("INIT_V6_VSSO", "1"); }
    if (!SCRIPT_PROP.getProperty("INIT_V7_GLOBAL_RED")) { addGlobalRedRecipients_(); SCRIPT_PROP.setProperty("INIT_V7_GLOBAL_RED", "1"); }
    if (!SCRIPT_PROP.getProperty("INIT_V8_VSSO_CAT")) { migrateVssoCategory_(ss); SCRIPT_PROP.setProperty("INIT_V8_VSSO_CAT", "1"); }
    if (!SCRIPT_PROP.getProperty("INIT_V9_VLI_CONTENT")) { initVLIContent_(); SCRIPT_PROP.setProperty("INIT_V9_VLI_CONTENT", "1"); }
    // Force VLI and SAC ISP content refresh for v272+
    if (!SCRIPT_PROP.getProperty("INIT_V10_VLI_VEHICLE_CHECK")) {
      initVLIContent_();
      // Also refresh SAC ISP in FORMS_JSON
      let f_ = {};
      const sv_ = SCRIPT_PROP.getProperty("FORMS_JSON");
      if (sv_) try { f_ = JSON.parse(sv_); } catch(e) { f_ = {}; }
      f_["SAC ISP"] = getSacISPContent_();
      SCRIPT_PROP.setProperty("FORMS_JSON", JSON.stringify(f_));
      SCRIPT_PROP.setProperty("INIT_V10_VLI_VEHICLE_CHECK", "1");
    }
    // v273: Remove DIO from SAC ISP, remove Tensiometre from both
    if (!SCRIPT_PROP.getProperty("INIT_V11_SAC_ISP_DIO")) {
      initVLIContent_();
      let f11_ = {};
      const sv11_ = SCRIPT_PROP.getProperty("FORMS_JSON");
      if (sv11_) try { f11_ = JSON.parse(sv11_); } catch(e) { f11_ = {}; }
      f11_["SAC ISP"] = getSacISPContent_();
      SCRIPT_PROP.setProperty("FORMS_JSON", JSON.stringify(f11_));
      SCRIPT_PROP.setProperty("INIT_V11_SAC_ISP_DIO", "1");
    }
    // Charger les formulaires depuis les feuilles Contenu_* (si la fonction existe)
    if (typeof initializeForms === 'function') {
      initializeForms();
    } else if (typeof loadFormStructures === 'function') {
      loadFormStructures();
    } else {
      Logger.log("initializeForms introuvable: formulaires non rechargés.");
    }
    
    // 1. Config
    const confSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
    const confData = confSheet.getDataRange().getValues();
    let frequencies = {};
    let categoriesOrder = [];
    
    for (let i = 1; i < confData.length; i++) {
      if(confData[i][0]) {
        frequencies[confData[i][0]] = confData[i][1];
        categoriesOrder.push(confData[i][0]);
      }
    }
    
    // 2. Inventaire
    const invSheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
    const invData = invSheet.getDataRange().getValues();
    let inventory = [];
    let dashboard = {};
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    for (let i = 1; i < invData.length; i++) {
      const row = invData[i];
      if (!row[0]) continue;
      
      // Recalcul dynamique du statut basé sur la date de prochain contrôle
      let calculatedStatus = row[4] || "green";
      if (row[4] !== 'purple' && row[10] !== 'HS' && row[3]) {
        const nextD = new Date(row[3]);
        nextD.setHours(0, 0, 0, 0);
        const daysLeft = Math.floor((nextD - today) / (1000 * 60 * 60 * 24));
        if (daysLeft < 0) calculatedStatus = "red";
        else if (daysLeft < 7) calculatedStatus = "orange";
        else calculatedStatus = "green";
      }
      
      const item = {
        category: row[0],
        name: row[1],
        lastDate: formatDate(row[2]),
        nextDate: formatDate(row[3]),
        status: calculatedStatus,
        lastVerifier: row[5],
        nextItemName: row[6],
        nextItemDate: formatDate(row[7]),
        mailOrange: row[8],
        mailRed: row[9],
        state: row[10],
        location: row[11] || "",
        order: row[12] || "",
        subType: row[13] || ""
      };
      
      inventory.push(item);
      
      if (!dashboard[item.category]) dashboard[item.category] = [];
      dashboard[item.category].push(item);
    }
    
    // 3. Forms (Checklists)
    let forms = {};
    const savedForms = SCRIPT_PROP.getProperty("FORMS_JSON");
    if (savedForms) forms = JSON.parse(savedForms);
    // Toujours injecter le contenu VSSO (loadFormStructures peut l'écraser)
    forms["VSSO"] = getVSSOContent_();
    
    // 4. Historique
    const histSheet = ss.getSheetByName(SHEET_NAMES.HISTORY);
    const lastRow = histSheet.getLastRow();
    let history = [];
    if (lastRow > 1) {
      const startRow = Math.max(2, lastRow - 500); 
      const histData = histSheet.getRange(startRow, 1, lastRow - startRow + 1, 4).getValues();
      
      for (let i = histData.length - 1; i >= 0; i--) {
        history.push({
          dateStr: formatDate(histData[i][0], true),
          name: histData[i][1],
          verifier: histData[i][2],
          details: histData[i][3]
        });
      }
    }
    
    // 5. Options & Stats
    let options = JSON.parse(SCRIPT_PROP.getProperty("GLOBAL_OPTS") || "{}");
    let mailConfig = JSON.parse(SCRIPT_PROP.getProperty("MAIL_CONF") || "{}");
    
    let stats = { ok:0, orange:0, red:0, expiredItems:0 };
    inventory.forEach(i => {
      if(i.state !== 'HS') {
        if(i.status === 'green') stats.ok++;
        if(i.status === 'orange') stats.orange++;
        if(i.status === 'red') stats.red++;
        if(i.status === 'purple') stats.expiredItems++;
      }
    });

    // 6. Centres externes (pour dropdown localisation)
    let centers = [];
    try { centers = getExternalAgentsData().centers || []; } catch(e) { Logger.log("Erreur chargement centres: " + e); }
    // Toujours inclure les centres virtuels, même si le spreadsheet externe échoue
    const vcNames = Object.keys(VIRTUAL_CENTERS);
    vcNames.forEach(vc => { if (centers.indexOf(vc) === -1) centers.push(vc); });
    centers.sort();

    return {
      success: true,
      data: {
        inventory: inventory,
        dashboard: dashboard,
        categoriesOrder: categoriesOrder,
        frequencies: frequencies,
        forms: forms,
        history: history,
        options: options,
        mailConfig: mailConfig,
        stats: stats,
        centers: centers
      }
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Fonction pour recalculer les statuts basés sur les dates
function recalculateStatuses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = invSheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  for (let i = 1; i < data.length; i++) {
    const nextDate = data[i][3]; // Colonne Prochain_Controle
    if (!nextDate) continue;
    
    let status = "green";
    const nextD = new Date(nextDate);
    nextD.setHours(0, 0, 0, 0);
    const daysLeft = Math.floor((nextD - today) / (1000 * 60 * 60 * 24));
    
    if (daysLeft < 0) status = "red";
    else if (daysLeft < 7) status = "orange";
    
    invSheet.getRange(i + 1, 5).setValue(status); // Colonne Statut
  }
  
  Logger.log("Statuts recalculés");
}

// --- ACTIONS PRINCIPALES ---

function saveCheck(bagName, formData, nextItemName, nextItemDate, verifierName, verificationTime, observations) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = invSheet.getDataRange().getValues();
  
  let bagRowIndex = -1;
  let category = "";
  let currentFreq = 30;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == bagName) {
      bagRowIndex = i + 1;
      category = data[i][0];
      break;
    }
  }
  
  if (bagRowIndex === -1) return { success: false, error: "Sac non trouvé" };
  
  const confSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  const confData = confSheet.getDataRange().getValues();
  for(let i=1; i<confData.length; i++) {
    if(confData[i][0] == category) {
      currentFreq = parseInt(confData[i][1]) || 30;
      break;
    }
  }
  
  // Garde PSud: fréquence forcée à 1 jour
  const bagLocation = (data[bagRowIndex - 1][11] || "").trim();
  const isGardePSud_ = bagName === "VLI 08" 
    || normalizeCenter_(bagLocation) === "garde psud"
    || (category === "VSSO" && normalizeCenter_(bagLocation) === "perpignan sud");
  if (isGardePSud_) currentFreq = 1;
  
  const now = new Date();
  const next = new Date();
  next.setDate(now.getDate() + currentFreq);
  
  let status = "green";
  let itemAlert = "";
  
  if (nextItemDate) {
    const itemD = new Date(nextItemDate);
    const diffTime = itemD - now;
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
    
    if (diffDays < 0) {
      status = "purple";
      itemAlert = "OBJET PÉRIMÉ : " + nextItemName;
    }
  }
  
  invSheet.getRange(bagRowIndex, 3).setValue(now);
  invSheet.getRange(bagRowIndex, 4).setValue(next);
  invSheet.getRange(bagRowIndex, 5).setValue(status);
  invSheet.getRange(bagRowIndex, 6).setValue(verifierName);
  invSheet.getRange(bagRowIndex, 7).setValue(nextItemName);
  invSheet.getRange(bagRowIndex, 8).setValue(nextItemDate);
  
  const histSheet = ss.getSheetByName(SHEET_NAMES.HISTORY);
  let detailString = JSON.stringify(formData);
  if (observations) detailString += " || OBSERVATIONS: " + observations;
  if(itemAlert) detailString += " || " + itemAlert;
  
  // Ajouter le temps de vérification dans le détail
  const timeInfo = (verificationTime !== undefined && verificationTime !== null && verificationTime !== "") ? ` [⏱️ ${verificationTime}]` : "";
  
  histSheet.appendRow([now, bagName, verifierName, detailString + timeInfo]);

  invalidateCache_();
  
  return { success: true };
}

// --- GESTION DES PHOTOS (GOOGLE DRIVE) ---

function getPhotoFolder() {
  const folders = DriveApp.getFoldersByName("APP_PHOTOS_VERIF");
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder("APP_PHOTOS_VERIF");
  }
}

function saveBagPhoto(category, bagName, section, base64Data) {
  try {
    const folder = getPhotoFolder();
    const photoKey = makePhotoKey_(category, bagName, section);
    const sanitized = sanitizeBagName_(photoKey);
    const timestamp = new Date().getTime();
    const fileName = "PHOTO_" + sanitized + "_" + timestamp + ".jpg";
    
    Logger.log("Enregistrement photo pour: " + photoKey + " -> " + fileName);

    const existing = getBagPhotos(category, bagName, section);
    const action = existing && existing.length > 0 ? "modify" : "add";
    
    // Création de la nouvelle photo avec timestamp
    const data = base64Data.split(",")[1]; 
    const blob = Utilities.newBlob(Utilities.base64Decode(data), "image/jpeg", fileName);
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Stocker métadonnées dans la description
    const desc = "DESC:Photo de vérification | CAT:" + category + " | BAG:" + bagName + " | SEC:" + section + " | KEY:" + photoKey;
    file.setDescription(desc);
    
    logPhotoEvent(action, bagName, file.getId(), fileName);

    updatePhotoPresence_(photoKey, true);

    invalidateCache_();

    Logger.log("Photo sauvée avec succès: " + file.getId());
    return { success: true, fileId: file.getId(), fileName: fileName, timestamp: timestamp, url: file.getUrl() };
  } catch (e) {
    Logger.log("ERREUR sauvegarde photo: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function getBagPhotos(category, bagName, section) {
  try {
    const folder = getPhotoFolder();
    const photoKey = makePhotoKey_(category, bagName, section);
    const sanitized = sanitizeBagName_(photoKey);
    const prefix = "PHOTO_" + sanitized + "_";
    const photos = [];
    
    // Chercher tous les fichiers et les filtrer par prefix
    const allFiles = folder.getFiles();
    while (allFiles.hasNext()) {
      const file = allFiles.next();
      if (file.getName().startsWith(prefix)) {
        const blob = file.getBlob();
        const base64 = "data:image/jpeg;base64," + Utilities.base64Encode(blob.getBytes());
        const desc = file.getDescription() || "";
        const descMatch = desc.match(/DESC:([^|]*)/);
        const description = descMatch ? descMatch[1].trim() : "";
        
        const timestamp = parseInt(file.getName().split("_").pop().replace(".jpg", "")) || 0;
        photos.push({
          fileId: file.getId(),
          fileName: file.getName(),
          timestamp: timestamp,
          base64: base64,
          description: description,
          dateStr: timestamp > 0 ? new Date(timestamp).toLocaleString() : "Sans date"
        });
      }
    }
    
    // Trier par timestamp décroissant (photos récentes en premier)
    photos.sort((a, b) => b.timestamp - a.timestamp);
    Logger.log("getBagPhotos(" + photoKey + "): trouvé " + photos.length + " photos");
    return photos;
  } catch (e) {
    Logger.log("ERREUR getBagPhotos: " + e.toString());
    return [];
  }
}

function getBagPhoto(category, bagName, section) {
  // Compatibilité - retourne la photo la plus récente
  const photos = getBagPhotos(category, bagName, section);
  return photos.length > 0 ? photos[0].base64 : null;
}

function getBagLatestPhotoMeta(category, bagName, section) {
  try {
    const photos = getBagPhotos(category, bagName, section);
    if (photos.length > 0) {
      return {
        hasPhoto: true,
        base64: photos[0].base64,
        fileId: photos[0].fileId,
        timestamp: photos[0].timestamp || null
      };
    }
    return { hasPhoto: false };
  } catch (e) {
    Logger.log("ERREUR getBagLatestPhotoMeta: " + e.toString());
    return { hasPhoto: false, error: e.toString() };
  }
}

function deletePhotoFile(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const desc = file.getDescription() || "";
    const catMatch = desc.match(/CAT:([^|]*)/);
    const bagMatch = desc.match(/BAG:([^|]*)/);
    const secMatch = desc.match(/SEC:([^|]*)/);
    const category = catMatch ? catMatch[1].trim() : "";
    const bagName = bagMatch ? bagMatch[1].trim() : "Unknown";
    const section = secMatch ? secMatch[1].trim() : "Unknown";
    
    file.setTrashed(true);
    logPhotoEvent("delete", bagName, fileId, file.getName());

    if (category && bagName && section) {
      const photoKey = makePhotoKey_(category, bagName, section);
      updatePhotoPresence_(photoKey, false);
    } else {
      rebuildPhotoPresenceMap_();
    }

    invalidateCache_();
    
    Logger.log("Photo supprimée: " + fileId);
    return { success: true };
  } catch (e) {
    Logger.log("ERREUR suppression photo: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function logPhotoEvent(action, bagName, fileId, fileName) {
  try {
    const prop = PropertiesService.getScriptProperties();
    const histStr = prop.getProperty("PHOTO_HISTORY") || "[]";
    let history = JSON.parse(histStr);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
    const invData = invSheet.getDataRange().getValues();
    let bagCategory = "";
    for (let i = 1; i < invData.length; i++) {
      if (invData[i][1] === bagName) {
        bagCategory = invData[i][0];
        break;
      }
    }
    
    history.push({
      action: action,
      bagName: bagName,
      category: bagCategory,
      fileId: fileId,
      fileName: fileName,
      timestamp: new Date().getTime(),
      dateStr: new Date().toLocaleString()
    });
    
    prop.setProperty("PHOTO_HISTORY", JSON.stringify(history));
  } catch (e) {
    Logger.log("ERREUR logPhotoEvent: " + e.toString());
  }
}

function getPhotoHistory() {
  try {
    const prop = PropertiesService.getScriptProperties();
    const histStr = prop.getProperty("PHOTO_HISTORY") || "[]";
    let history = JSON.parse(histStr);
    
    // Charger les photos pour celles encore existantes
    const folder = getPhotoFolder();
    const existingFiles = {};
    const allFiles = folder.getFiles();
    while (allFiles.hasNext()) {
      const f = allFiles.next();
      existingFiles[f.getId()] = f;
    }
    
    history.forEach(h => {
      if (existingFiles[h.fileId] && (h.action === "add" || h.action === "modify")) {
        const f = existingFiles[h.fileId];
        const blob = f.getBlob();
        h.base64 = "data:image/jpeg;base64," + Utilities.base64Encode(blob.getBytes());
      }
    });
    
    history.sort((a, b) => b.timestamp - a.timestamp);
    return history;
  } catch (e) {
    Logger.log("ERREUR getPhotoHistory: " + e.toString());
    return [];
  }
}

function getPhotoPresenceMap() {
  try {
    const prop = SCRIPT_PROP.getProperty(PHOTO_PRESENCE_KEY);
    if (prop) return JSON.parse(prop);
  } catch (e) {
    Logger.log("ERREUR getPhotoPresenceMap: " + e.toString());
  }
  return rebuildPhotoPresenceMap_();
}

function rebuildPhotoPresenceMap_() {
  const map = {};
  try {
    const folder = getPhotoFolder();
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      const name = f.getName() || "";
      const match = name.match(/^PHOTO_(.+)_\d+\.jpg$/);
      if (match && match[1]) map[match[1]] = true;
    }
    SCRIPT_PROP.setProperty(PHOTO_PRESENCE_KEY, JSON.stringify(map));
  } catch (e) {
    Logger.log("ERREUR rebuildPhotoPresenceMap_: " + e.toString());
  }
  return map;
}

function updatePhotoPresence_(photoKey, hasPhoto) {
  try {
    const map = getPhotoPresenceMap() || {};
    const sanitized = sanitizeBagName_(photoKey);
    if (hasPhoto) map[sanitized] = true; else delete map[sanitized];
    SCRIPT_PROP.setProperty(PHOTO_PRESENCE_KEY, JSON.stringify(map));
  } catch (e) {
    Logger.log("ERREUR updatePhotoPresence_: " + e.toString());
  }
}

function makePhotoKey_(category, bagName, section) {
  return (category || "") + "||" + (bagName || "") + "||" + (section || "");
}

function sanitizeBagName_(str) {
  return (str || "").replace(/[^a-zA-Z0-9]/g, "_");
}

// Fonction de test pour créer le dossier et tester une photo
function testPhotoSystem() {
  try {
    const folder = getPhotoFolder();
    Logger.log("Dossier créé/trouvé: " + folder.getName() + " (ID: " + folder.getId() + ")");
    Logger.log("URL du dossier: " + folder.getUrl());
    
    // Test de création d'une photo simple
    const testData = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="; // 1x1 pixel rouge
    const blob = Utilities.newBlob(Utilities.base64Decode(testData), "image/png", "TEST.png");
    const testFile = folder.createFile(blob);
    testFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    Logger.log("Fichier test créé: " + testFile.getName());
    Logger.log("URL test: " + testFile.getUrl());
    
    return {
      success: true,
      folderId: folder.getId(),
      folderUrl: folder.getUrl(),
      testFileUrl: testFile.getUrl()
    };
  } catch (e) {
    Logger.log("ERREUR testPhotoSystem: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

// --- FONCTIONS ADMIN ---

function getNextOrder_(sheet, cat) {
  const data = sheet.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === cat) {
      const v = parseInt(data[i][12], 10);
      if (!isNaN(v) && v > max) max = v;
    }
  }
  return max + 1;
}

function initializeInventoryOrder_(ss) {
  const s = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = s.getDataRange().getValues();
  const counters = {};
  for (let i = 1; i < data.length; i++) {
    const cat = String(data[i][0]).trim();
    if (!cat) continue;
    if (!counters[cat]) counters[cat] = 1;
    const v = parseInt(data[i][12], 10);
    if (isNaN(v) || v <= 0) {
      s.getRange(i + 1, 13).setValue(counters[cat]);
    }
    counters[cat]++;
  }
}

function addBag(cat, name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const nextOrder = getNextOrder_(s, cat);
  s.appendRow([cat, name, "", "", "green", "", "", "", "", "", "Actif", "", nextOrder, ""]);
  invalidateCache_();
}

function updateBagSubType(name, subType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == name) {
      s.getRange(i + 1, 14).setValue(subType || "");
      break;
    }
  }
  invalidateCache_();
}

function deleteBag(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = s.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][1] == name) {
      s.deleteRow(i+1);
      break;
    }
  }
  invalidateCache_();
}

function updateBagStatus(name, state) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = s.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][1] == name) {
      s.getRange(i+1, 11).setValue(state);
      break;
    }
  }
  invalidateCache_();
}

function createNewCategory(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEET_NAMES.CONFIG);
  s.appendRow([name, 30]);
  invalidateCache_();
}

function renameBag(oldName, newName) {
  if(!oldName || !newName) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let bagCategory = "";
  
  // Renommer dans l'inventaire
  const invSheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const invData = invSheet.getDataRange().getValues();
  for(let i = 1; i < invData.length; i++) {
    if(invData[i][1] === oldName) {
      bagCategory = invData[i][0];
      invSheet.getRange(i + 1, 2).setValue(newName);
    }
  }
  
  // Renommer dans l'historique
  const histSheet = ss.getSheetByName(SHEET_NAMES.HISTORY);
  const histData = histSheet.getDataRange().getValues();
  for(let i = 1; i < histData.length; i++) {
    if(histData[i][1] === oldName) {
      histSheet.getRange(i + 1, 2).setValue(newName);
    }
  }

  // Renommer les photos liées (standard + impact)
  renameBagPhotos_(oldName, newName, bagCategory);
  invalidateCache_();
}

function renameBagPhotos_(oldName, newName, bagCategory) {
  try {
    const folder = getPhotoFolder();
    const files = folder.getFiles();
    const sanitizedOld = sanitizeBagName_(oldName);
    const sanitizedNew = sanitizeBagName_(newName);
    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();
      const desc = file.getDescription() || "";
      // IMPACT photos
      if (name.startsWith("IMPACT_" + sanitizedOld + "_")) {
        const newNameFile = name.replace("IMPACT_" + sanitizedOld + "_", "IMPACT_" + sanitizedNew + "_");
        const newDesc = desc.replace(/BAG:([^|]*)/, "BAG:" + newName);
        file.setName(newNameFile);
        file.setDescription(newDesc);
        continue;
      }
      // Standard photos
      if (desc.indexOf("BAG:" + oldName) !== -1 || name.indexOf("PHOTO_" + sanitizedOld) === 0) {
        const catMatch = desc.match(/CAT:([^|]*)/);
        const secMatch = desc.match(/SEC:([^|]*)/);
        const cat = catMatch ? catMatch[1].trim() : bagCategory;
        const sec = secMatch ? secMatch[1].trim() : "";
        if (!cat || !sec) continue;
        const newKey = makePhotoKey_(cat, newName, sec);
        const sanitizedKey = sanitizeBagName_(newKey);
        const timestamp = name.split("_").pop().replace(".jpg", "");
        const newFileName = "PHOTO_" + sanitizedKey + "_" + timestamp + ".jpg";
        const newDesc = desc
          .replace(/BAG:([^|]*)/, "BAG:" + newName)
          .replace(/KEY:([^|]*)/, "KEY:" + newKey);
        file.setName(newFileName);
        file.setDescription(newDesc);
      }
    }
    rebuildPhotoPresenceMap_();
  } catch (e) {
    Logger.log("ERREUR renameBagPhotos_: " + e.toString());
  }
}

function renameCategory(oldCat, newCat) {
  if(!oldCat || !newCat) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Renommer dans la config
  const confSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  const confData = confSheet.getDataRange().getValues();
  for(let i = 1; i < confData.length; i++) {
    if(confData[i][0] === oldCat) {
      confSheet.getRange(i + 1, 1).setValue(newCat);
    }
  }
  
  // Renommer dans l'inventaire
  const invSheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const invData = invSheet.getDataRange().getValues();
  for(let i = 1; i < invData.length; i++) {
    if(invData[i][0] === oldCat) {
      invSheet.getRange(i + 1, 1).setValue(newCat);
    }
  }
  
  // Renommer dans les formulaires
  const forms = SCRIPT_PROP.getProperty("FORMS_JSON");
  if(forms) {
    const formsObj = JSON.parse(forms);
    if(formsObj[oldCat]) {
      formsObj[newCat] = formsObj[oldCat];
      delete formsObj[oldCat];
      SCRIPT_PROP.setProperty("FORMS_JSON", JSON.stringify(formsObj));
    }
  }
  invalidateCache_();
}

function deleteCategory(categoryName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Supprimer de la config
  const confSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  const confData = confSheet.getDataRange().getValues();
  for(let i = confData.length - 1; i >= 1; i--) {
    if(confData[i][0] === categoryName) {
      confSheet.deleteRow(i + 1);
      break;
    }
  }
  
  // 2. Supprimer tous les items de cette catégorie dans l'inventaire
  const invSheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const invData = invSheet.getDataRange().getValues();
  for(let i = invData.length - 1; i >= 1; i--) {
    if(invData[i][0] === categoryName) {
      invSheet.deleteRow(i + 1);
    }
  }
  
  // 3. Supprimer les formulaires de cette catégorie
  const forms = SCRIPT_PROP.getProperty("FORMS_JSON");
  if(forms) {
    const formsObj = JSON.parse(forms);
    delete formsObj[categoryName];
    SCRIPT_PROP.setProperty("FORMS_JSON", JSON.stringify(formsObj));
  }

  invalidateCache_();
  
  return { success: true };
}

function deleteHistoryEntry(historyIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const histSheet = ss.getSheetByName(SHEET_NAMES.HISTORY);
  const lastRow = histSheet.getLastRow();
  
  // L'historique est affiché en ordre inversé (plus récent en premier)
  // Donc l'index 0 = dernière ligne, index 1 = avant-dernière, etc.
  const rowToDelete = lastRow - historyIndex;
  
  if(rowToDelete > 1 && rowToDelete <= lastRow) {
    histSheet.deleteRow(rowToDelete);
    invalidateCache_();
    return { success: true };
  }
  
  return { success: false, error: "Ligne introuvable" };
}

function updateCategoriesConfig(confArray) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEET_NAMES.CONFIG);
  s.clearContents();
  s.appendRow(["Categorie", "Frequence_Jours"]);
  confArray.forEach(c => {
    s.appendRow([c.name, c.freq]);
  });
  invalidateCache_();
}

function updateCategoryContent(catName, dataJson) {
  let forms = {};
  const saved = SCRIPT_PROP.getProperty("FORMS_JSON");
  if(saved) forms = JSON.parse(saved);
  
  let groups = {};
  dataJson.forEach(row => {
    if(!groups[row.section]) {
      groups[row.section] = { section: row.section, position: row.position, items: [] };
    }
    if(row.position) groups[row.section].position = row.position;
    
    if(row.item) {
      groups[row.section].items.push({
        name: row.item,
        type: row.type,
        def: row.def,
        subsection: row.subsection || ""
      });
    }
  });
  
  let structured = [];
  for (let key in groups) {
    structured.push(groups[key]);
  }
  
  forms[catName] = structured;
  SCRIPT_PROP.setProperty("FORMS_JSON", JSON.stringify(forms));
  invalidateCache_();
}

function updateBagMails(bag, type, val) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = s.getDataRange().getValues();
  const col = type === 'orange' ? 9 : 10;
  
  for(let i=1; i<data.length; i++) {
    if(data[i][1] == bag) {
      s.getRange(i+1, col).setValue(val);
      break;
    }
  }
  invalidateCache_();
}

function updateVliLocation(bag, location) {
  // Délègue vers la version avec mise à jour mails
  return updateBagLocationWithMails(bag, location);
}

function updateVliLocationsBatch(list) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = s.getDataRange().getValues();
  const map = {};
  (list || []).forEach(it => { if (it && it.name) map[it.name] = it.location || ""; });
  
  let extAgents = null;
  try { extAgents = getExternalAgentsData(); } catch(e) { Logger.log("Erreur lecture agents externes: " + e); }
  
  for (let i = 1; i < data.length; i++) {
    const name = data[i][1];
    const category = String(data[i][0]).trim();
    if (map.hasOwnProperty(name)) {
      const loc = map[name];
      s.getRange(i + 1, 12).setValue(loc);
      // Auto-maj mails si pas SAC IADE
      if (category !== "SAC IADE" && extAgents) {
        applyMailsForLocation_(s, i + 1, loc, extAgents);
      }
    }
  }
  invalidateCache_();
  return { success: true };
}

function updateBagOrders(list) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = s.getDataRange().getValues();
  const map = {};
  (list || []).forEach(it => { if (it && it.name) map[it.name] = parseInt(it.order, 10) || ""; });
  for (let i = 1; i < data.length; i++) {
    const name = data[i][1];
    if (map.hasOwnProperty(name)) {
      s.getRange(i + 1, 13).setValue(map[name]);
    }
  }
  invalidateCache_();
  return { success: true };
}

function saveGlobalOptions(opts) {
  SCRIPT_PROP.setProperty("GLOBAL_OPTS", JSON.stringify(opts));
  invalidateCache_();
}

function saveMailConfig(conf) {
  SCRIPT_PROP.setProperty("MAIL_CONF", JSON.stringify(conf));
  invalidateCache_();
}

function formatDate(dateObj, withTime) {
  if (!dateObj || dateObj === "") return "";
  const d = new Date(dateObj);
  if (isNaN(d.getTime())) return "";
  
  let day = ("0" + d.getDate()).slice(-2);
  let month = ("0" + (d.getMonth() + 1)).slice(-2);
  let year = d.getFullYear();
  
  let res = `${day}/${month}/${year}`;
  if(withTime) {
    let h = ("0" + d.getHours()).slice(-2);
    let m = ("0" + d.getMinutes()).slice(-2);
    res += ` ${h}:${m}`;
  }
  return res;
}

// --- TRIGGER (AUTOMATISATION) ---
function installTrigger(hour) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkDailyAlerts') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger('checkDailyAlerts')
      .timeBased()
      .everyDays(1)
      .atHour(parseInt(hour))
      .create();
      
  return "Automatisation activée à " + hour + "h00 tous les jours.";
}

function checkDailyAlerts() {
  const data = getData();
  if(!data.success) return;
  
  const inv = data.data.inventory;
  const conf = data.data.mailConfig;
  
  inv.forEach(item => {
    if(item.state === 'HS') return;
    
    let sendMail = false;
    let subject = "";
    let body = "";
    let recipient = "";
    
    if(item.status === 'red' || item.status === 'purple') {
      const GLOBAL_RED = "brice.dubrey@sdis66.fr,florian.bois@sdis66.fr";
      let allRecipients = item.mailRed ? item.mailRed.replace(/;/g, ",").trim() : "";
      if (allRecipients) allRecipients += "," + GLOBAL_RED;
      else allRecipients = GLOBAL_RED;
      recipient = allRecipients;
      subject = conf.redSub || "ALERTE ROUGE";
      body = conf.redBody || "Matériel périmé.";
      sendMail = true;
    }
    else if(item.status === 'orange') {
      if(item.mailOrange) {
        recipient = item.mailOrange;
        subject = conf.orangeSub || "ALERTE ORANGE";
        body = conf.orangeBody || "Matériel bientot périmé.";
        sendMail = true;
      }
    }
    
    if(sendMail && recipient) {
      body = body.replace(/{nom}/g, item.name)
                 .replace(/{categorie}/g, item.category)
                 .replace(/{date}/g, item.lastDate)
                 .replace(/{echeance}/g, item.nextDate);
      
      subject = subject.replace(/{nom}/g, item.name);
      
      // Convertir les ; en , pour MailApp et dédupliquer
      const cleanRecipient = [...new Set(recipient.replace(/;/g, ",").split(",").map(e => e.trim()).filter(e => e))].join(",");
      
      try {
        MailApp.sendEmail(cleanRecipient, subject, body);
      } catch(e) {
        console.log("Erreur envoi mail: " + e);
      }
    }
  });
}

// ===================================================================
// === SYNC MAILS FROM AGENT SPREADSHEET ===
// ===================================================================

/**
 * Lit le spreadsheet agents et retourne un objet { centerName: [email1, email2, ...] }
 */
function getAgentEmailsByCenter_() {
  const agentSS = SpreadsheetApp.openById(EXTERNAL_SS_ID);
  const sheet = agentSS.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < data.length; i++) {
    const email = (data[i][3] || "").trim();
    if (!email) continue;
    const main = (data[i][1] || "").trim();
    const secondary = (data[i][2] || "").trim();

    [main, secondary].forEach(center => {
      if (!center) return;
      const key = normalizeCenter_(center);
      if (!map[key]) map[key] = [];
      if (!map[key].includes(email)) map[key].push(email);
    });
  }
  return map;
}

/**
 * Retourne un objet { "NOM_DE_FAMILLE_NORMALISE": email } pour lookup IADE
 */
function getAgentEmailByLastName_() {
  const agentSS = SpreadsheetApp.openById(EXTERNAL_SS_ID);
  const sheet = agentSS.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < data.length; i++) {
    const fullName = (data[i][0] || "").trim();
    const email = (data[i][3] || "").trim();
    if (!fullName || !email) continue;
    const lastName = fullName.split(/\s+/)[0].toUpperCase();
    map[lastName] = email;
  }
  return map;
}

/**
 * Normalise un nom de centre pour le matching
 */
function normalizeCenter_(name) {
  if (!name) return "";
  let n = name.trim();
  const ALIASES = {
    "le boulou": "boulou",
    "pnord": "perpignan nord",
    "p nord": "perpignan nord",
    "psud": "perpignan sud",
    "p sud": "perpignan sud",
    "pouest": "perpignan ouest",
    "p ouest": "perpignan ouest",
    "garde psud": "garde psud",
    "st cyprien": "saint cyprien",
    "st paul de fenouillet": "saint paul de fenouillet",
    "vinça": "vinca",
    "ille": "ille sur tet",
    "côte vermeille": "cote vermeille"
  };
  const lower = n.toLowerCase();
  if (ALIASES[lower]) return ALIASES[lower];
  return lower.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

/**
 * Extrait le nom du centre depuis le nom du sac/VLI
 */
function extractCenterFromBag_(name, category, location) {
  if (category === "VLI") {
    return normalizeCenter_(location || "");
  }
  if (category === "SAC ISP" || category === "SAC RESERVE") {
    let clean = name;
    clean = clean.replace(/^sac\s+(isp|reserve|réserve)\s+/i, "");
    clean = clean.replace(/\s+\d+$/, "");
    return normalizeCenter_(clean.trim());
  }
  return "";
}

/**
 * Synchronise les mails orange et rouge de tout l'inventaire
 * depuis le spreadsheet des agents (batch).
 */
function syncMailsFromAgentSheet() {
  try {
    const centerEmails = getAgentEmailsByCenter_();
    const lastNameEmails = getAgentEmailByLastName_();
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const inv = ss.getSheetByName(SHEET_NAMES.INVENTORY);
    const data = inv.getDataRange().getValues();
    
    let updated = 0;
    
    for (let i = 1; i < data.length; i++) {
      const category = (data[i][0] || "").trim();
      const name = (data[i][1] || "").trim();
      const location = (data[i][11] || "").trim();
      
      let mails = "";
      
      if (name === "VLI 05") {
        mails = "florian.bois@sdis66.fr; brice.dubrey@sdis66.fr";
      }
      else if (name === "VLI 08") {
        mails = "florian.bois@sdis66.fr; brice.dubrey@sdis66.fr";
      }
      else if (category === "SAC IADE") {
        const iadeMatch = name.match(/sac\s+iade\s+(.+)/i);
        if (iadeMatch) {
          const lastName = iadeMatch[1].trim().toUpperCase();
          for (let key in lastNameEmails) {
            if (key === lastName || key.startsWith(lastName) || lastName.startsWith(key)) {
              mails = lastNameEmails[key];
              break;
            }
          }
        }
      }
      else {
        const center = extractCenterFromBag_(name, category, location);
        if (center && centerEmails[center]) {
          mails = centerEmails[center].join("; ");
        }
      }
      
      if (mails) {
        inv.getRange(i + 1, 9).setValue(mails);
        inv.getRange(i + 1, 10).setValue(mails);
        updated++;
      }
    }
    
    invalidateCache_();
    return { success: true, message: "✅ " + updated + " sacs/VLI mis à jour avec les mails des agents." };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ===================================================================
// === RENAME IADE BAGS (one-time migration) ===
// ===================================================================

function renameIadeBags() {
  const IADE_RENAME = {
    "Iade 1": "Sac IADE Bedu",
    "Iade 2": "Sac IADE Comas",
    "Iade 3": "Sac IADE Le Roy",
    "Iade 4": "Sac IADE Petipre",
    "Iade 5": "Sac IADE Py",
    "Iade 6": "Sac IADE Spilemont"
  };
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = inv.getDataRange().getValues();
  let renamed = 0;
  
  for (let i = 1; i < data.length; i++) {
    const name = (data[i][1] || "").trim();
    if (IADE_RENAME[name]) {
      inv.getRange(i + 1, 2).setValue(IADE_RENAME[name]);
      renamed++;
    }
  }
  
  const hist = ss.getSheetByName(SHEET_NAMES.HISTORY);
  if (hist && hist.getLastRow() > 1) {
    const histData = hist.getDataRange().getValues();
    for (let i = 1; i < histData.length; i++) {
      const hName = (histData[i][1] || "").trim();
      if (IADE_RENAME[hName]) {
        hist.getRange(i + 1, 2).setValue(IADE_RENAME[hName]);
      }
    }
  }
  
  invalidateCache_();
  return "✅ " + renamed + " sacs IADE renommés.";
}

// ===================================================================
// === GARDE PSUD — Vérification quotidienne 16h ===
// ===================================================================

/**
 * Vérifie les véhicules Garde PSud (ex: VLI 08).
 * Si pas vérifié aujourd'hui à 16h → rouge + mail.
 */
function checkGardePSudAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = inv.getDataRange().getValues();
  const today = new Date();
  const todayStr = Utilities.formatDate(today, "Europe/Paris", "dd/MM/yyyy");
  
  const GARDE_PSUD_MAILS = "florian.bois@sdis66.fr; brice.dubrey@sdis66.fr";
  
  for (let i = 1; i < data.length; i++) {
    const name = (data[i][1] || "").trim();
    const location = (data[i][11] || "").trim();
    const state = (data[i][10] || "").trim();
    const lastDate = data[i][2];
    
    const bagCategory = (data[i][0] || "").trim();
    const isGardePSud = normalizeCenter_(location) === "garde psud" 
                     || name === "VLI 08"
                     || (bagCategory === "VSSO" && normalizeCenter_(location) === "perpignan sud");
    
    if (!isGardePSud || state === "HS") continue;
    
    let checkedToday = false;
    if (lastDate) {
      const ld = new Date(lastDate);
      const ldStr = Utilities.formatDate(ld, "Europe/Paris", "dd/MM/yyyy");
      checkedToday = (ldStr === todayStr);
    }
    
    if (!checkedToday) {
      inv.getRange(i + 1, 5).setValue("red");
      try {
        MailApp.sendEmail(
          GARDE_PSUD_MAILS,
          "🔴 ALERTE GARDE PSUD — " + name + " non vérifié aujourd'hui!",
          "ATTENTION!\n\nLe véhicule " + name + " (Garde PSud) n'a PAS été vérifié aujourd'hui.\n\nLa vérification quotidienne est OBLIGATOIRE pour les véhicules de Garde PSud.\n\nMerci d'agir immédiatement."
        );
      } catch(e) {
        console.log("Erreur envoi mail Garde PSud: " + e);
      }
    } else {
      inv.getRange(i + 1, 5).setValue("green");
      const tomorrow = new Date(today);
      tomorrow.setDate(tomorrow.getDate() + 1);
      inv.getRange(i + 1, 4).setValue(tomorrow);
    }
  }
  
  invalidateCache_();
}

/**
 * Installe le trigger Garde PSud à 16h
 */
function installGardePSudTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkGardePSudAlerts') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger('checkGardePSudAlerts')
      .timeBased()
      .everyDays(1)
      .atHour(16)
      .create();
  
  const fixResult = fixGardePSudDates();
  return "✅ Trigger Garde PSud activé — vérification quotidienne à 16h.\n" + fixResult;
}

/**
 * Corrige immédiatement les dates des véhicules Garde PSud
 */
function fixGardePSudDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = inv.getDataRange().getValues();
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  let fixed = 0;
  
  for (let i = 1; i < data.length; i++) {
    const name = (data[i][1] || "").trim();
    const location = (data[i][11] || "").trim();
    const state = (data[i][10] || "").trim();
    const bagCategory = (data[i][0] || "").trim();
    
    const isGP = normalizeCenter_(location) === "garde psud" 
              || name === "VLI 08"
              || (bagCategory === "VSSO" && normalizeCenter_(location) === "perpignan sud");
    
    if (!isGP || state === "HS") continue;
    
    inv.getRange(i + 1, 4).setValue(tomorrow);
    fixed++;
  }
  
  invalidateCache_();
  return "✅ " + fixed + " véhicules Garde PSud mis à jour (prochain contrôle = demain).";
}

// ===================================================================
// === EXTERNAL AGENTS / CENTRES (Carte Prévisionnelle) — Dropdown ===
// ===================================================================

/**
 * Lit les ISP depuis le spreadsheet externe "Carte Prévisionnelle Opérationnelle"
 * Colonnes : A=Nom, B=Centre Principal, C=Centre Secondaire, D=Email
 * Résultat mis en cache 24h dans ScriptProperties
 */
function refreshExternalAgents() {
  try {
    const ss = SpreadsheetApp.openById(EXTERNAL_SS_ID);
    const sheet = ss.getSheetByName(EXTERNAL_ISP_SHEET);
    if (!sheet) { Logger.log("Onglet ISP introuvable dans le spreadsheet externe"); return { centers: [], agents: [] }; }

    const data = sheet.getDataRange().getValues();
    const agents = [];
    const centersSet = {};

    for (let i = 1; i < data.length; i++) {
      const nom = String(data[i][0] || "").trim();
      const principal = String(data[i][1] || "").trim();
      const secondaire = String(data[i][2] || "").trim();
      const email = String(data[i][3] || "").trim();
      if (!nom) continue;
      agents.push({ nom: nom, principal: principal, secondaire: secondaire, email: email });
      if (principal) centersSet[principal] = true;
      if (secondaire) centersSet[secondaire] = true;
    }

    // Injecter les centres virtuels
    Object.keys(VIRTUAL_CENTERS).forEach(vc => { centersSet[vc] = true; });
    const centers = Object.keys(centersSet).sort();
    const result = { centers: centers, agents: agents, timestamp: Date.now() };
    SCRIPT_PROP.setProperty("EXT_AGENTS_CACHE", JSON.stringify(result));
    Logger.log("Agents externes rafraîchis: " + agents.length + " agents, " + centers.length + " centres");
    return { centers: centers, agents: agents };
  } catch (e) {
    Logger.log("ERREUR refreshExternalAgents: " + e.toString());
    return { centers: Object.keys(VIRTUAL_CENTERS).sort(), agents: [] };
  }
}

/**
 * Retourne les données agents/centres depuis le cache (ou rafraîchit si >24h)
 */
function getExternalAgentsData() {
  const cached = SCRIPT_PROP.getProperty("EXT_AGENTS_CACHE");
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (parsed.timestamp && (Date.now() - parsed.timestamp < 86400000)) {
        return { centers: parsed.centers || [], agents: parsed.agents || [] };
      }
    } catch(e) {}
  }
  return refreshExternalAgents();
}

/**
 * Retourne les emails des ISP affectés (principal OU secondaire) à un centre donné
 */
function getEmailsForCenter_(centerName) {
  if (!centerName) return [];
  const data = getExternalAgentsData();
  const emails = [];
  (data.agents || []).forEach(a => {
    if ((a.principal === centerName || a.secondaire === centerName) && a.email) {
      emails.push(a.email);
    }
  });
  return [...new Set(emails)];
}

/**
 * Met à jour la localisation d'un sac/VLI et auto-renseigne les mails ISP
 * Exception : SAC IADE → mails inchangés (sac personnel)
 */
function updateBagLocationWithMails(bagName, location) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = s.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == bagName) {
      const category = String(data[i][0]).trim();
      s.getRange(i + 1, 12).setValue(location || "");

      if (category !== "SAC IADE") {
        let extAgents = null;
        try { extAgents = getExternalAgentsData(); } catch(e) {}
        applyMailsForLocation_(s, i + 1, location, extAgents);
      }
      break;
    }
  }
  invalidateCache_();
  return { success: true };
}

/**
 * Helper : écrit les mails orange/rouge sur une ligne d'inventaire
 * en se basant sur la localisation (centre) sélectionnée
 */
function applyMailsForLocation_(sheet, rowNum, location, extData) {
  const GLOBAL_RED = "brice.dubrey@sdis66.fr;florian.bois@sdis66.fr";
  if (!location) {
    sheet.getRange(rowNum, 9).setValue("");
    sheet.getRange(rowNum, 10).setValue(GLOBAL_RED);
    return;
  }
  // Centre virtuel (Astreinte Départementale Médicale, Garde PSud, etc.)
  if (VIRTUAL_CENTERS[location]) {
    const vcEmails = VIRTUAL_CENTERS[location];
    sheet.getRange(rowNum, 9).setValue(vcEmails);  // Mail_Orange
    sheet.getRange(rowNum, 10).setValue(vcEmails); // Mail_Red
    return;
  }
  // Centre ISP classique
  if (extData) {
    const emails = [];
    (extData.agents || []).forEach(a => {
      if ((a.principal === location || a.secondaire === location) && a.email) emails.push(a.email);
    });
    const uniqueEmails = [...new Set(emails)];
    const emailStr = uniqueEmails.join(";");
    sheet.getRange(rowNum, 9).setValue(emailStr);
    sheet.getRange(rowNum, 10).setValue(emailStr ? emailStr + ";" + GLOBAL_RED : GLOBAL_RED);
  } else {
    sheet.getRange(rowNum, 9).setValue("");
    sheet.getRange(rowNum, 10).setValue(GLOBAL_RED);
  }
}

// ===================================================================
// === VLI IMPACT PHOTOS SYSTEM ===
// ===================================================================

function saveVliImpact(bagName, base64Data, comment) {
  try {
    const folder = getPhotoFolder();
    const timestamp = new Date().getTime();
    const sanitized = sanitizeBagName_(bagName);
    const fileName = "IMPACT_" + sanitized + "_" + timestamp + ".jpg";
    const data = base64Data.split(",")[1];
    const blob = Utilities.newBlob(Utilities.base64Decode(data), "image/jpeg", fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    file.setDescription("IMPACT|BAG:" + bagName + "|COMMENT:" + (comment || ""));
    invalidateCache_();
    return { success: true, fileId: file.getId() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getVliImpacts(bagName) {
  try {
    const folder = getPhotoFolder();
    const sanitized = sanitizeBagName_(bagName);
    const prefix = "IMPACT_" + sanitized + "_";
    const impacts = [];
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().startsWith(prefix)) {
        const desc = file.getDescription() || "";
        const commentMatch = desc.match(/COMMENT:(.*)/);
        const comment = commentMatch ? commentMatch[1].trim() : "";
        const blob = file.getBlob();
        const base64 = "data:image/jpeg;base64," + Utilities.base64Encode(blob.getBytes());
        const timestamp = parseInt(file.getName().split("_").pop().replace(".jpg", "")) || 0;
        impacts.push({
          fileId: file.getId(),
          base64: base64,
          comment: comment,
          timestamp: timestamp,
          dateStr: timestamp > 0 ? new Date(timestamp).toLocaleString('fr-FR') : "Sans date"
        });
      }
    }
    impacts.sort((a, b) => b.timestamp - a.timestamp);
    return impacts;
  } catch (e) {
    Logger.log("ERREUR getVliImpacts: " + e.toString());
    return [];
  }
}

function deleteVliImpact(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    invalidateCache_();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function updateVliImpactComment(fileId, newComment) {
  try {
    const file = DriveApp.getFileById(fileId);
    let desc = file.getDescription() || "";
    desc = desc.replace(/COMMENT:.*/, "COMMENT:" + newComment);
    file.setDescription(desc);
    invalidateCache_();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ===================================================================
// === VLI MILEAGE SYSTEM ===
// ===================================================================

function saveVliMileage(bagName, km, dateStr) {
  const key = "VLI_KM_" + sanitizeBagName_(bagName);
  const data = { km: km, date: dateStr, timestamp: new Date().getTime() };
  SCRIPT_PROP.setProperty(key, JSON.stringify(data));
  invalidateCache_();
  return { success: true };
}

function invalidateCache_() {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove("BOOTSTRAP_V1");
    cache.remove("BOOTSTRAP_V1_N");
    // Remove any chunked cache entries
    for (var c = 0; c < 20; c++) { try { cache.remove("BOOTSTRAP_V1_C" + c); } catch(e) {} }
    rebuildBootstrapSnapshot_();
  } catch (e) {
    Logger.log("Cache invalidate error: " + e.toString());
  }
}

function getAllVliMileages() {
  const props = SCRIPT_PROP.getProperties();
  const result = {};
  for (let key in props) {
    if (key.startsWith("VLI_KM_")) {
      try {
        result[key.replace("VLI_KM_", "")] = JSON.parse(props[key]);
      } catch(e) {}
    }
  }
  return result;
}

// ===================================================================
// === DEFAULT CONTENT INITIALIZATION ===
// ===================================================================

function initializeDefaultContent() {
  let forms = {};
  const saved = SCRIPT_PROP.getProperty("FORMS_JSON");
  if (saved) { try { forms = JSON.parse(saved); } catch(e) { forms = {}; } }

  forms["SAC ISP"] = getSacISPContent_();

  forms["SAC RESERVE"] = getSacReserveContent_();
  if (!forms["SAC IADE"] || forms["SAC IADE"].length === 0) {
    forms["SAC IADE"] = [{ section: "Contenu général", position: "", items: [{ name: "À définir", type: "case", def: "true" }] }];
  }

  SCRIPT_PROP.setProperty("FORMS_JSON", JSON.stringify(forms));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const confSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (confSheet) {
    const confData = confSheet.getDataRange().getValues();
    const existingCats = confData.slice(1).map(r => String(r[0]).trim());
    ["VLI", "SAC ISP", "SAC RESERVE", "SAC IADE"].forEach(cat => {
      if (!existingCats.includes(cat)) {
        confSheet.appendRow([cat, 30]);
      }
    });
  }
  return "Contenu initialisé! SAC ISP: " + forms["SAC ISP"].length + " sections, SAC RESERVE: " + forms["SAC RESERVE"].length + " sections";
}

function cleanupCategories_(ss) {
  // === STANDARD CATEGORY NAMES ===
  const STANDARD = { "VLI": "VLI", "SAC ISP": "SAC ISP", "Sac ISP": "SAC ISP", "sac isp": "SAC ISP",
    "SAC RESERVE": "SAC RESERVE", "Sac RESERVE": "SAC RESERVE", "SAC IADE": "SAC IADE", "Sac IADE": "SAC IADE", "Sac Iade": "SAC IADE", "sac iade": "SAC IADE" };
  function norm(n) { const t = String(n).trim(); return STANDARD[t] || t.toUpperCase(); }
  
  // 1. DEDUPLICATE CONFIG - garder une seule ligne par catégorie normalisée
  const confSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  const confData = confSheet.getDataRange().getValues();
  const seen = {}; const rowsToDelete = [];
  for (let i = 1; i < confData.length; i++) {
    const raw = String(confData[i][0]).trim();
    if (!raw) continue;
    const std = norm(raw);
    if (seen[std]) { rowsToDelete.push(i + 1); } // doublon
    else { seen[std] = true; if (raw !== std) confSheet.getRange(i + 1, 1).setValue(std); }
  }
  for (let i = rowsToDelete.length - 1; i >= 0; i--) confSheet.deleteRow(rowsToDelete[i]);
  // Ajouter SAC RESERVE si absent
  const finalConf = confSheet.getDataRange().getValues();
  const finalCats = finalConf.slice(1).map(r => String(r[0]).trim());
  if (!finalCats.includes("SAC RESERVE")) confSheet.appendRow(["SAC RESERVE", 30]);
  
  // 2. MIGRATE FORMS_JSON keys to standard names + inject missing content
  let forms = {};
  const saved = SCRIPT_PROP.getProperty("FORMS_JSON");
  if (saved) { try { forms = JSON.parse(saved); } catch(e) { forms = {}; } }
  const newForms = {};
  for (let key in forms) { newForms[norm(key)] = forms[key]; }
  if (!newForms["SAC ISP"] || newForms["SAC ISP"].length === 0) newForms["SAC ISP"] = getSacISPContent_();
  if (!newForms["SAC RESERVE"] || newForms["SAC RESERVE"].length === 0) newForms["SAC RESERVE"] = getSacReserveContent_();
  SCRIPT_PROP.setProperty("FORMS_JSON", JSON.stringify(newForms));
  
  // 3. NORMALIZE Inventaire categories (no auto-add items)
  const invSheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const invData = invSheet.getDataRange().getValues();
  for (let i = 1; i < invData.length; i++) {
    const raw = String(invData[i][0]).trim();
    if (raw && norm(raw) !== raw) invSheet.getRange(i + 1, 1).setValue(norm(raw));
  }
}

function removeAutoDefaultBags_(ss) {
  const invSheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const invData = invSheet.getDataRange().getValues();
  const autoNames = new Set(["Sac ISP 1", "Sac IADE 1", "Sac Réserve 1", "Sac Réserve 2"]);
  const rowsToDelete = [];
  for (let i = 1; i < invData.length; i++) {
    const name = String(invData[i][1] || "").trim();
    if (!name || autoNames.has(name)) rowsToDelete.push(i + 1);
  }
  for (let i = rowsToDelete.length - 1; i >= 0; i--) invSheet.deleteRow(rowsToDelete[i]);
}

function removeAutoReserveBags() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const invData = invSheet.getDataRange().getValues();
  const targets = new Set(["Sac Réserve 1", "Sac Réserve 2"]);
  const rowsToDelete = [];
  for (let i = 1; i < invData.length; i++) {
    const cat = String(invData[i][0]).trim();
    const name = String(invData[i][1]).trim();
    if (cat === "SAC RESERVE" && targets.has(name)) rowsToDelete.push(i + 1);
  }
  for (let i = rowsToDelete.length - 1; i >= 0; i--) invSheet.deleteRow(rowsToDelete[i]);
  return "Auto Sac Réserve supprimés";
}

function runCleanupNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  cleanupCategories_(ss);
  return "Cleanup OK";
}

function getSacISPContent_() {
  return [
    { section: "Sac ISP", position: "Sac ISP", items: [
      { name: "Ampoulier Médicaments", type: "nombre", def: "1", subsection: "" },
      { name: "Drap UU", type: "nombre", def: "1", subsection: "" },
      { name: "Perfalgan", type: "nombre", def: "1", subsection: "Intérieur du sac" },
      { name: "NaCl 100ml", type: "nombre", def: "1", subsection: "Intérieur du sac" },
      { name: "NaCl 500ml", type: "nombre", def: "1", subsection: "Intérieur du sac" },
      { name: "Ringer Lactate", type: "nombre", def: "1", subsection: "Intérieur du sac" },
      { name: "Glucosé 10% 250ml", type: "nombre", def: "1", subsection: "Intérieur du sac" },
      { name: "Glucosé 5% 500ml", type: "nombre", def: "1", subsection: "Intérieur du sac" },
      { name: "Garrot", type: "nombre", def: "1", subsection: "Pochette jaune" },
      { name: "Sparadrap", type: "nombre", def: "1", subsection: "Pochette jaune" },
      { name: "Kit à perfusion", type: "nombre", def: "1", subsection: "Pochette jaune" },
      { name: "Paquet de compresses stériles", type: "nombre", def: "1", subsection: "Pochette jaune" },
      { name: "Bouchons obturateurs", type: "nombre", def: "2", subsection: "Pochette jaune" },
      { name: "Seringue 10ml", type: "nombre", def: "1", subsection: "Pochette jaune" },
      { name: "Bétadine alcoolique", type: "nombre", def: "1", subsection: "Pochette jaune" },
      { name: "Valve anti-retour", type: "nombre", def: "4", subsection: "Pochette jaune" },
      { name: "Trocards", type: "nombre", def: "2", subsection: "Pochette jaune" },
      { name: "Gel hydroalcoolique", type: "nombre", def: "1", subsection: "Pochette violette" },
      { name: "Medinette", type: "nombre", def: "1", subsection: "Pochette violette" },
      { name: "DASRI", type: "nombre", def: "1", subsection: "Pochette violette" },
      { name: "Gants UU", type: "nombre", def: "2", subsection: "Pochette violette" },
      { name: "Sacs poubelles", type: "nombre", def: "2", subsection: "Pochette violette" },
      { name: "Manche laryngo monté et lame n°4", type: "nombre", def: "1", subsection: "Sacoche bleue" },
      { name: "Tube laryngé pédiatrique", type: "nombre", def: "1", subsection: "Sacoche bleue" },
      { name: "Tube laryngé T4", type: "nombre", def: "1", subsection: "Sacoche bleue" },
      { name: "Piles de rechange", type: "nombre", def: "2", subsection: "Sacoche bleue" },
      { name: "Cale dents", type: "nombre", def: "1", subsection: "Sacoche bleue" },
      { name: "Kit allégé", type: "nombre", def: "1", subsection: "Sacoche bleue" },
      { name: "Pince Magyll", type: "nombre", def: "1", subsection: "Sacoche bleue" },
      { name: "Perfuseur", type: "nombre", def: "1", subsection: "Pochette rouge" },
      { name: "Opsite", type: "nombre", def: "1", subsection: "Pochette rouge" },
      { name: "Compresses stériles", type: "nombre", def: "2", subsection: "Pochette rouge" },
      { name: "Kit perfusion", type: "nombre", def: "1", subsection: "Pochette rouge" },
      { name: "Penthrox", type: "nombre", def: "1", subsection: "Pochette verte" },
      { name: "Aérosol adulte", type: "nombre", def: "1", subsection: "Pochette verte" },
      { name: "Aérosol enfant", type: "nombre", def: "1", subsection: "Pochette verte" },
      { name: "Perceuse", type: "nombre", def: "1", subsection: "Sacoche jaune DIO" },
      { name: "Aiguille rose 15mm", type: "nombre", def: "1", subsection: "Sacoche jaune DIO" },
      { name: "Aiguille bleue 25mm", type: "nombre", def: "1", subsection: "Sacoche jaune DIO" },
      { name: "Aiguille jaune 45mm", type: "nombre", def: "1", subsection: "Sacoche jaune DIO" },
      { name: "NaCl 10ml", type: "nombre", def: "2", subsection: "Sacoche jaune DIO" },
      { name: "Seringue 20cc", type: "nombre", def: "1", subsection: "Sacoche jaune DIO" },
      { name: "Trocard", type: "nombre", def: "1", subsection: "Sacoche jaune DIO" },
      { name: "Manchon compression", type: "nombre", def: "1", subsection: "Sacoche jaune DIO" },
      { name: "Thermomètre", type: "nombre", def: "1", subsection: "Pochette droite" },
      { name: "Lecteur glycémie", type: "nombre", def: "1", subsection: "Pochette droite" },
      { name: "Sucres sachet", type: "nombre", def: "2", subsection: "Pochette droite" },
      { name: "Lancettes", type: "nombre", def: "10", subsection: "Pochette droite" },
      { name: "Bandelettes dextro", type: "nombre", def: "10", subsection: "Pochette droite" },
      { name: "Compresses", type: "nombre", def: "10", subsection: "Pochette droite" },
      { name: "Pansement hémostatique", type: "nombre", def: "1", subsection: "Pochette droite" },
      { name: "Garrot tourniquet", type: "nombre", def: "1", subsection: "Pochette droite" },
      { name: "Ciseaux Gesko", type: "nombre", def: "1", subsection: "Pochette droite" },
      { name: "Stéthoscope simple pavillon", type: "nombre", def: "1", subsection: "Pochette gauche" },
      { name: "Sonde gastrique n°14", type: "nombre", def: "1", subsection: "Pochette gauche" },
      { name: "Sonde gastrique n°18", type: "nombre", def: "1", subsection: "Pochette gauche" },
      { name: "Seringue gavage 60ml", type: "nombre", def: "1", subsection: "Pochette gauche" },
      { name: "Poche urine", type: "nombre", def: "1", subsection: "Pochette gauche" }
    ]}
  ];
}

function getSacReserveContent_() {
  return [
    { section: "Solutés", position: "Poche principale", items: [
      { name: "Kit Perfusion (4)", type: "nombre", def: "4" },
      { name: "NaCl 0.9% 500ml (4)", type: "nombre", def: "4" },
      { name: "Kit Perfalgan (4)", type: "nombre", def: "4" },
      { name: "Kit NaCl 100ml (4)", type: "nombre", def: "4" },
      { name: "Kétoprofène 100ml (2)", type: "nombre", def: "2" },
      { name: "Glucose 10% 250ml (2)", type: "nombre", def: "2" },
      { name: "Ringer Lactate 500ml (1)", type: "case", def: "true" },
      { name: "Glucose 5% 500ml (1)", type: "case", def: "true" },
      { name: "Penthrox (1)", type: "case", def: "true" }
    ]},
    { section: "Pochette rouge longue — Intubation", position: "Poche principale", items: [
      { name: "Tube laryngé adulte taille 4 (1)", type: "case", def: "true" },
      { name: "Seringue étalonnée (1)", type: "case", def: "true" },
      { name: "Cale dents (1)", type: "case", def: "true" },
      { name: "Lame laryngoscope UU n°3 (1)", type: "case", def: "true" }
    ]},
    { section: "Poche latérale droite", position: "Latéral droit", items: [
      { name: "Perfuseur 3 voies (1)", type: "case", def: "true" },
      { name: "Bouchons (2)", type: "nombre", def: "2" },
      { name: "Seringues 10ml (2)", type: "nombre", def: "2" },
      { name: "Trocards (2)", type: "nombre", def: "2" },
      { name: "Valves anti-retour (2)", type: "nombre", def: "2" },
      { name: "Opsite (1)", type: "case", def: "true" },
      { name: "DASRI Médinette (1)", type: "case", def: "true" }
    ]},
    { section: "Poche latérale gauche — Aérosols", position: "Latéral gauche", items: [
      { name: "Kit aérosol adulte (2)", type: "nombre", def: "2" },
      { name: "Kit aérosol enfant (1)", type: "case", def: "true" }
    ]}
  ];
}

function getVLIContent_() {
  return [
    { section: "Réserve médicament", position: "Coffre", items: [
      { name: "Ampoulier réserve (dans coffre)", type: "nombre", def: "1", subsection: "" }
    ]},
    { section: "Arrière du véhicule", position: "Arrière du véhicule", items: [
      { name: "Gants UU S", type: "nombre", def: "1", subsection: "" },
      { name: "Gants UU M", type: "nombre", def: "1", subsection: "" },
      { name: "Gants UU L", type: "nombre", def: "1", subsection: "" },
      { name: "Gants UU XL", type: "nombre", def: "1", subsection: "" },
      { name: "Gel hydroalcoolique", type: "nombre", def: "1", subsection: "" }
    ]},
    { section: "Antares", position: "Antares", items: [
      { name: "Poste antares et chargeur", type: "nombre", def: "1", subsection: "" },
      { name: "Batterie réserve", type: "nombre", def: "1", subsection: "" }
    ]},
    { section: "Scope Schiller", position: "Scope Schiller", items: [
      { name: "Scope", type: "nombre", def: "1", subsection: "" },
      { name: "Rasoir", type: "nombre", def: "1", subsection: "" },
      { name: "Sachet électrodes", type: "nombre", def: "1", subsection: "" },
      { name: "Électrode DSA", type: "nombre", def: "1", subsection: "" },
      { name: "Câble ECG", type: "nombre", def: "1", subsection: "" },
      { name: "Câble saturation adulte", type: "nombre", def: "1", subsection: "" },
      { name: "Câble saturation enfant", type: "nombre", def: "1", subsection: "" },
      { name: "Brassard PNI", type: "nombre", def: "3", subsection: "" },
      { name: "Chargeur", type: "nombre", def: "1", subsection: "" },
      { name: "Imprimante", type: "nombre", def: "1", subsection: "" },
      { name: "Dispositif capno", type: "nombre", def: "3", subsection: "" }
    ]},
    { section: "PSE", position: "PSE", items: [
      { name: "PSE", type: "nombre", def: "2", subsection: "" },
      { name: "Seringue 50cc + prolongateur", type: "nombre", def: "1", subsection: "Kit PSE Sacoche bleue" },
      { name: "Robinet 3 voies", type: "nombre", def: "2", subsection: "Kit PSE Sacoche bleue" }
    ]},
    { section: "Aspirateur de mucosité", position: "Aspirateur de mucosité", items: [
      { name: "Aspirateur", type: "nombre", def: "1", subsection: "" },
      { name: "Tuyau aspiration monté", type: "nombre", def: "1", subsection: "" },
      { name: "Poche aspiration montée", type: "nombre", def: "1", subsection: "" },
      { name: "Peters", type: "nombre", def: "1", subsection: "Sur le côté" },
      { name: "Yankauer", type: "nombre", def: "1", subsection: "Sur le côté" },
      { name: "Sonde aspi CH10", type: "nombre", def: "1", subsection: "Sur le côté" },
      { name: "Sonde aspi CH14", type: "nombre", def: "1", subsection: "Sur le côté" },
      { name: "Sonde aspi CH16", type: "nombre", def: "1", subsection: "Sur le côté" }
    ]},
    { section: "Kit Décontamination", position: "Kit Décontamination", items: [
      { name: "Masques FFP2", type: "nombre", def: "20", subsection: "" },
      { name: "Kit protection allégé", type: "nombre", def: "2", subsection: "" },
      { name: "Bactinul", type: "nombre", def: "1", subsection: "" },
      { name: "Aniosgel", type: "nombre", def: "1", subsection: "" },
      { name: "Sacs poubelles", type: "nombre", def: "10", subsection: "" },
      { name: "Masques chirurgicaux", type: "nombre", def: "20", subsection: "" },
      { name: "Lingettes antibactériennes (sachet de 50)", type: "nombre", def: "1", subsection: "" },
      { name: "Medinette grand modèle", type: "nombre", def: "1", subsection: "" }
    ]},
    { section: "Frigo", position: "Frigo", items: [
      { name: "Si présent, contrôle température 2 à 8", type: "case", def: "false", subsection: "" }
    ]},
    { section: "Sac ISP", position: "Sac ISP", items: [
      { name: "Ampoulier Médicaments", type: "nombre", def: "1", subsection: "" },
      { name: "Drap UU", type: "nombre", def: "1", subsection: "" },
      { name: "Perfalgan", type: "nombre", def: "1", subsection: "Intérieur du sac" },
      { name: "NaCl 100ml", type: "nombre", def: "1", subsection: "Intérieur du sac" },
      { name: "NaCl 500ml", type: "nombre", def: "1", subsection: "Intérieur du sac" },
      { name: "Ringer Lactate", type: "nombre", def: "1", subsection: "Intérieur du sac" },
      { name: "Glucosé 10% 250ml", type: "nombre", def: "1", subsection: "Intérieur du sac" },
      { name: "Glucosé 5% 500ml", type: "nombre", def: "1", subsection: "Intérieur du sac" },
      { name: "Garrot", type: "nombre", def: "1", subsection: "Pochette jaune" },
      { name: "Sparadrap", type: "nombre", def: "1", subsection: "Pochette jaune" },
      { name: "Kit à perfusion", type: "nombre", def: "1", subsection: "Pochette jaune" },
      { name: "Paquet de compresses stériles", type: "nombre", def: "1", subsection: "Pochette jaune" },
      { name: "Bouchons obturateurs", type: "nombre", def: "2", subsection: "Pochette jaune" },
      { name: "Seringue 10ml", type: "nombre", def: "1", subsection: "Pochette jaune" },
      { name: "Bétadine alcoolique", type: "nombre", def: "1", subsection: "Pochette jaune" },
      { name: "Valve anti-retour", type: "nombre", def: "4", subsection: "Pochette jaune" },
      { name: "Trocards", type: "nombre", def: "2", subsection: "Pochette jaune" },
      { name: "Gel hydroalcoolique", type: "nombre", def: "1", subsection: "Pochette violette" },
      { name: "Medinette", type: "nombre", def: "1", subsection: "Pochette violette" },
      { name: "DASRI", type: "nombre", def: "1", subsection: "Pochette violette" },
      { name: "Gants UU", type: "nombre", def: "2", subsection: "Pochette violette" },
      { name: "Sacs poubelles", type: "nombre", def: "2", subsection: "Pochette violette" },
      { name: "Manche laryngo monté et lame n°4", type: "nombre", def: "1", subsection: "Sacoche bleue" },
      { name: "Tube laryngé pédiatrique", type: "nombre", def: "1", subsection: "Sacoche bleue" },
      { name: "Tube laryngé T4", type: "nombre", def: "1", subsection: "Sacoche bleue" },
      { name: "Piles de rechange", type: "nombre", def: "2", subsection: "Sacoche bleue" },
      { name: "Cale dents", type: "nombre", def: "1", subsection: "Sacoche bleue" },
      { name: "Kit allégé", type: "nombre", def: "1", subsection: "Sacoche bleue" },
      { name: "Pince Magyll", type: "nombre", def: "1", subsection: "Sacoche bleue" },
      { name: "Perfuseur", type: "nombre", def: "1", subsection: "Pochette rouge" },
      { name: "Opsite", type: "nombre", def: "1", subsection: "Pochette rouge" },
      { name: "Compresses stériles", type: "nombre", def: "2", subsection: "Pochette rouge" },
      { name: "Kit perfusion", type: "nombre", def: "1", subsection: "Pochette rouge" },
      { name: "Penthrox", type: "nombre", def: "1", subsection: "Pochette verte" },
      { name: "Aérosol adulte", type: "nombre", def: "1", subsection: "Pochette verte" },
      { name: "Aérosol enfant", type: "nombre", def: "1", subsection: "Pochette verte" },
      { name: "Perceuse", type: "nombre", def: "1", subsection: "Sacoche jaune DIO" },
      { name: "Aiguille rose 15mm", type: "nombre", def: "1", subsection: "Sacoche jaune DIO" },
      { name: "Aiguille bleue 25mm", type: "nombre", def: "1", subsection: "Sacoche jaune DIO" },
      { name: "Aiguille jaune 45mm", type: "nombre", def: "1", subsection: "Sacoche jaune DIO" },
      { name: "NaCl 10ml", type: "nombre", def: "2", subsection: "Sacoche jaune DIO" },
      { name: "Seringue 20cc", type: "nombre", def: "1", subsection: "Sacoche jaune DIO" },
      { name: "Trocard", type: "nombre", def: "1", subsection: "Sacoche jaune DIO" },
      { name: "Manchon compression", type: "nombre", def: "1", subsection: "Sacoche jaune DIO" },
      { name: "Thermomètre", type: "nombre", def: "1", subsection: "Pochette droite" },
      { name: "Lecteur glycémie", type: "nombre", def: "1", subsection: "Pochette droite" },
      { name: "Sucres sachet", type: "nombre", def: "2", subsection: "Pochette droite" },
      { name: "Lancettes", type: "nombre", def: "10", subsection: "Pochette droite" },
      { name: "Bandelettes dextro", type: "nombre", def: "10", subsection: "Pochette droite" },
      { name: "Compresses", type: "nombre", def: "10", subsection: "Pochette droite" },
      { name: "Pansement hémostatique", type: "nombre", def: "1", subsection: "Pochette droite" },
      { name: "Garrot tourniquet", type: "nombre", def: "1", subsection: "Pochette droite" },
      { name: "Ciseaux Gesko", type: "nombre", def: "1", subsection: "Pochette droite" },
      { name: "Stéthoscope simple pavillon", type: "nombre", def: "1", subsection: "Pochette gauche" },
      { name: "Sonde gastrique n°14", type: "nombre", def: "1", subsection: "Pochette gauche" },
      { name: "Sonde gastrique n°18", type: "nombre", def: "1", subsection: "Pochette gauche" },
      { name: "Seringue gavage 60ml", type: "nombre", def: "1", subsection: "Pochette gauche" },
      { name: "Poche urine", type: "nombre", def: "1", subsection: "Pochette gauche" }
    ]},
    { section: "Sac Réserve", position: "Sac Réserve", items: [
      { name: "Kit perfusion", type: "nombre", def: "5", subsection: "" },
      { name: "Kit perfalgan", type: "nombre", def: "5", subsection: "" },
      { name: "NaCl 100ml", type: "nombre", def: "3", subsection: "" },
      { name: "Kit aérosol adulte", type: "nombre", def: "2", subsection: "" },
      { name: "Kit aérosol enfant", type: "nombre", def: "1", subsection: "" },
      { name: "Électrode DSA", type: "nombre", def: "1", subsection: "" },
      { name: "Rasoirs", type: "nombre", def: "1", subsection: "" },
      { name: "Rouleur papier ECG", type: "nombre", def: "1", subsection: "" },
      { name: "Électrodes ECG", type: "nombre", def: "2", subsection: "" },
      { name: "Penthrox", type: "nombre", def: "2", subsection: "" },
      { name: "Drape UU", type: "nombre", def: "1", subsection: "" },
      { name: "NaCl 500ml", type: "nombre", def: "5", subsection: "" },
      { name: "Ringer Lactate", type: "nombre", def: "2", subsection: "" },
      { name: "Glucosé 10% 250ml", type: "nombre", def: "3", subsection: "" },
      { name: "Glucosé 5% 500ml", type: "nombre", def: "2", subsection: "" },
      { name: "Medinette", type: "nombre", def: "1", subsection: "" },
      { name: "Sac aspiration réserve avec tuyau aspi", type: "nombre", def: "1", subsection: "" },
      { name: "Pochette petit matériel (seringue 10, trocards, valve anti-retour)", type: "nombre", def: "1", subsection: "" }
    ]},
    { section: "Sac Meopa", position: "Sac Meopa", items: [
      { name: "Obus", type: "nombre", def: "1", subsection: "" },
      { name: "Circuit Meopa usage unique", type: "nombre", def: "5", subsection: "" },
      { name: "Filtres", type: "nombre", def: "5", subsection: "" },
      { name: "Masques taille 1 gris", type: "nombre", def: "2", subsection: "" },
      { name: "Masques taille 2", type: "nombre", def: "2", subsection: "" },
      { name: "Masques taille 3", type: "nombre", def: "3", subsection: "" },
      { name: "Masques taille 4", type: "nombre", def: "3", subsection: "" }
    ]},
    { section: "Sac oxygénothérapie", position: "Sac oxygénothérapie", items: [
      { name: "Sac scellé", type: "nombre", def: "1", subsection: "" }
    ]},
    { section: "Sac sousan", position: "Sac sousan", items: [
      { name: "Présent scellé", type: "nombre", def: "1", subsection: "" }
    ]},
    { section: "Petit Matériel", position: "Petit Matériel", items: [
      { name: "Carton DASRI", type: "nombre", def: "1", subsection: "" },
      { name: "Carton récupération consommable PUI", type: "case", def: "true", subsection: "" },
      { name: "Kit tire-tiques", type: "nombre", def: "1", subsection: "" },
      { name: "Appareil", type: "nombre", def: "1", subsection: "Analyseur Hb (Hemocue)" },
      { name: "Compresses non stériles", type: "nombre", def: "10", subsection: "Analyseur Hb (Hemocue)" },
      { name: "Lancettes", type: "nombre", def: "20", subsection: "Analyseur Hb (Hemocue)" },
      { name: "Cuvettes", type: "nombre", def: "20", subsection: "Analyseur Hb (Hemocue)" },
      { name: "Piles de rechange", type: "nombre", def: "4", subsection: "Analyseur Hb (Hemocue)" },
      { name: "Appareil", type: "nombre", def: "1", subsection: "Rad 57" },
      { name: "Piles de rechange", type: "nombre", def: "4", subsection: "Rad 57" },
      { name: "Flacon Cyanokit", type: "nombre", def: "1", subsection: "Kit cyanokit" },
      { name: "Perfuseur", type: "nombre", def: "1", subsection: "Kit cyanokit" },
      { name: "NaCl 250ml", type: "nombre", def: "1", subsection: "Kit cyanokit" },
      { name: "Seringue 50ml", type: "nombre", def: "1", subsection: "Kit cyanokit" },
      { name: "Trocard", type: "nombre", def: "1", subsection: "Kit cyanokit" },
      { name: "Kit accouchement", type: "nombre", def: "1", subsection: "Kit Accouchement" },
      { name: "Gants stériles 6,5/7,5/8,5", type: "nombre", def: "1", subsection: "Kit Accouchement" }
    ]},
    { section: "EPI", position: "EPI", items: [
      { name: "Casque F2", type: "nombre", def: "2", subsection: "" },
      { name: "Vestes de feu", type: "nombre", def: "2", subsection: "" },
      { name: "Chasubles", type: "nombre", def: "2", subsection: "" },
      { name: "Cônes signalisation", type: "nombre", def: "3", subsection: "" },
      { name: "Clé tricoise", type: "case", def: "true", subsection: "" },
      { name: "Jeu de chaînes à neige", type: "case", def: "true", subsection: "" },
      { name: "Badge télépéage", type: "case", def: "true", subsection: "" }
    ]},
    { section: "SSO", position: "SSO", items: [
      { name: "Caisses Orange", type: "case", def: "true", subsection: "" },
      { name: "Caisses Jaune", type: "case", def: "true", subsection: "" },
      { name: "Pack Eau", type: "case", def: "true", subsection: "" },
      { name: "Sac isotherme bleu métro", type: "case", def: "true", subsection: "" },
      { name: "Rallonge Maréchal", type: "case", def: "true", subsection: "" }
    ]},
    { section: "Vérification du véhicule", position: "Véhicule", items: [
      { name: "État général vérifié (pneus, carburant, essuie-glaces, ...)", type: "case", def: "false", subsection: "" },
      { name: "Commentaire état véhicule", type: "texte", def: "", subsection: "" }
    ]}
  ];
}

function initVLIContent_() {
  let forms = {};
  const saved = SCRIPT_PROP.getProperty("FORMS_JSON");
  if (saved) { try { forms = JSON.parse(saved); } catch(e) { forms = {}; } }
  forms["VLI"] = getVLIContent_();
  SCRIPT_PROP.setProperty("FORMS_JSON", JSON.stringify(forms));
  Logger.log("Contenu VLI initialisé/mis à jour");
}

// ===================================================================
// === VSSO CONTENT ===
// ===================================================================

function getVSSOContent_() {
  return [
    { section: "Capucine", position: "Capucine", items: [
      { name: "Papiers toilettes", type: "nombre", def: "2", subsection: "" },
      { name: "Chaises pliables", type: "nombre", def: "2", subsection: "" },
      { name: "Drapes jetables", type: "nombre", def: "2", subsection: "" },
      { name: "Rouleau tableau PMA", type: "nombre", def: "1", subsection: "" }
    ]},
    { section: "Tiroir 1", position: "Tiroir 1", items: [
      { name: "Gobelets prédosés en soupe", type: "nombre", def: "12", subsection: "" }
    ]},
    { section: "Tiroir 3", position: "Tiroir 3", items: [
      { name: "Rad-57", type: "nombre", def: "1", subsection: "" },
      { name: "Piles de rechange", type: "nombre", def: "4", subsection: "" },
      { name: "Fonctionnement vérifié", type: "case", def: "false", subsection: "" },
      { name: "Boite masques chirurgicaux", type: "nombre", def: "1", subsection: "" },
      { name: "Rouleur petits sacs poubelle", type: "nombre", def: "1", subsection: "" },
      { name: "Rouleur grands sacs poubelle", type: "nombre", def: "1", subsection: "" },
      { name: "Produit désinfection", type: "nombre", def: "1", subsection: "" },
      { name: "Paquet lingettes désinfection", type: "nombre", def: "1", subsection: "" },
      { name: "Rouleur petits sacs jaunes DASRI", type: "nombre", def: "1", subsection: "" }
    ]},
    { section: "Tiroir 4", position: "Tiroir 4", items: [
      { name: "Kit Tire tique", type: "nombre", def: "1", subsection: "" },
      { name: "Unidoses 10ml NaCl", type: "nombre", def: "35", subsection: "" },
      { name: "Solution hydroalcoolique", type: "nombre", def: "1", subsection: "" },
      { name: "Couverture de survie", type: "nombre", def: "5", subsection: "" },
      { name: "Boite gants d'hygiène", type: "nombre", def: "1", subsection: "" },
      { name: "Serviettes en papier", type: "nombre", def: "100", subsection: "" }
    ]},
    { section: "Tiroir 5", position: "Tiroir 5", items: [
      { name: "Sticks de sucre", type: "nombre", def: "300", subsection: "" },
      { name: "Boites allumettes", type: "nombre", def: "10", subsection: "" },
      { name: "Agitateurs", type: "nombre", def: "250", subsection: "" }
    ]},
    { section: "Au sol", position: "Au sol", items: [
      { name: "Sac de transport type sac de l'avant", type: "nombre", def: "2", subsection: "" },
      { name: "Frigo branché et thermostat position 2", type: "case", def: "false", subsection: "" },
      { name: "Bouteilles eau dans frigo", type: "nombre", def: "24", subsection: "" },
      { name: "Pompe grise XGloo (pour kit dalle Led)", type: "nombre", def: "1", subsection: "" },
      { name: "Groupe électrogène", type: "nombre", def: "1", subsection: "" },
      { name: "Rallonge / enrouleur électrique", type: "nombre", def: "1", subsection: "" }
    ]},
    { section: "Rangement bas paroie latérale droite", position: "Paroie latérale droite", items: [
      { name: "Chaises de réhabilitation", type: "nombre", def: "2", subsection: "" },
      { name: "Sac sousan", type: "nombre", def: "1", subsection: "" },
      { name: "Lits de camps", type: "nombre", def: "2", subsection: "" },
      { name: "Poubelle noire pliable ronde", type: "nombre", def: "1", subsection: "" },
      { name: "Bouteille O2 5l", type: "nombre", def: "1", subsection: "" }
    ]},
    { section: "Coffre paroie latérale droite", position: "Paroie latérale droite", items: [
      { name: "Boite sacs de réhabilitation", type: "nombre", def: "1", subsection: "" }
    ]},
    { section: "Paroie latérale haut gauche", position: "Paroie latérale gauche", items: [
      { name: "Gobelets prédosés en soupe (réserve)", type: "case", def: "true", subsection: "Coffre 1" },
      { name: "Gobelets prédosés en chocolat", type: "case", def: "true", subsection: "Coffre 1" },
      { name: "Gobelets jetables", type: "nombre", def: "50", subsection: "Coffre 2" },
      { name: "Petite bouilloire électrique", type: "nombre", def: "1", subsection: "Coffre 2" },
      { name: "Gobelets non jetables bleus", type: "nombre", def: "3", subsection: "Coffre 2" },
      { name: "Gobelets prédosés en café", type: "nombre", def: "75", subsection: "Coffre 2" }
    ]},
    { section: "Paroie gauche cellule", position: "Paroie gauche cellule", items: [
      { name: "Bouteille O2 15l", type: "nombre", def: "1", subsection: "" },
      { name: "Rampe O2 linde + flexible", type: "nombre", def: "1", subsection: "" },
      { name: "Masques haute concentration (dans 2 sacs rouges par 10)", type: "nombre", def: "20", subsection: "" },
      { name: "Chaise de portage", type: "nombre", def: "1", subsection: "" },
      { name: "Percolateur", type: "nombre", def: "1", subsection: "" }
    ]},
    { section: "Porte latérale extérieur côté conducteur", position: "Extérieur côté conducteur", items: [
      { name: "Kit dalle LED dans 2 sacs noirs", type: "nombre", def: "1", subsection: "" },
      { name: "Bidons essence pleins", type: "nombre", def: "2", subsection: "" },
      { name: "Rallonge prise maréchal", type: "nombre", def: "1", subsection: "" },
      { name: "Cônes de Lübeck pliables", type: "nombre", def: "3", subsection: "" },
      { name: "Petite poubelle accessible par trappe depuis cellule", type: "nombre", def: "1", subsection: "" },
      { name: "Extincteur à poudre", type: "nombre", def: "1", subsection: "" },
      { name: "Triangle de signalisation", type: "nombre", def: "1", subsection: "" }
    ]},
    { section: "Brancard", position: "Brancard", items: [
      { name: "Plan dur", type: "nombre", def: "1", subsection: "" },
      { name: "Fonctionnement brancard (montée / descente)", type: "case", def: "false", subsection: "" }
    ]},
    { section: "Poste de conduite", position: "Poste de conduite", items: [
      { name: "Antares + chargeur", type: "nombre", def: "1", subsection: "" },
      { name: "Classeur noir procédure VSSO", type: "nombre", def: "1", subsection: "" },
      { name: "Chasuble haute visibilité", type: "nombre", def: "2", subsection: "" },
      { name: "Badge télépéage", type: "nombre", def: "1", subsection: "" },
      { name: "Constat amiable + photocopie carte grise", type: "nombre", def: "1", subsection: "" }
    ]},
    { section: "Caisses extérieures", position: "Caisses extérieures", items: [
      { name: "Caisses oranges (dans frigo)", type: "nombre", def: "3", subsection: "Dans frigo" },
      { name: "Pack bouteille 24/pack (dans frigo)", type: "nombre", def: "3", subsection: "Dans frigo" },
      { name: "Caisses oranges (à côté frigo)", type: "nombre", def: "6", subsection: "À côté frigo" },
      { name: "Pack bouteille 24/pack (à côté frigo)", type: "nombre", def: "6", subsection: "À côté frigo" },
      { name: "Caisse isotherme de transport bleue", type: "nombre", def: "1", subsection: "À côté frigo" }
    ]}
  ];
}

function initVSSOContent_() {
  let forms = {};
  const saved = SCRIPT_PROP.getProperty("FORMS_JSON");
  if (saved) { try { forms = JSON.parse(saved); } catch(e) { forms = {}; } }
  forms["VSSO"] = getVSSOContent_();
  SCRIPT_PROP.setProperty("FORMS_JSON", JSON.stringify(forms));
  Logger.log("Contenu VSSO initialisé");
}

function addGlobalRedRecipients_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = s.getDataRange().getValues();
  const newEmails = ["brice.dubrey@sdis66.fr", "florian.bois@sdis66.fr"];
  
  for (let i = 1; i < data.length; i++) {
    const currentRed = data[i][9] ? data[i][9].toString().trim() : "";
    const existingEmails = currentRed ? currentRed.split(/[;,]/).map(e => e.trim()).filter(e => e) : [];
    let changed = false;
    newEmails.forEach(email => {
      if (!existingEmails.includes(email)) {
        existingEmails.push(email);
        changed = true;
      }
    });
    if (changed) {
      s.getRange(i + 1, 10).setValue(existingEmails.join(";"));
    }
  }
  Logger.log("Emails globaux rouges ajoutés à tous les items de l'inventaire");
}

/**
 * Migration V8 : Convertit les VLI ayant subType=VSSO en catégorie VSSO
 * et ajoute VSSO dans la feuille Config si absente.
 */
function migrateVssoCategory_(ss) {
  // 1. Ajouter VSSO dans Config si pas déjà présente
  const confSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  const confData = confSheet.getDataRange().getValues();
  let hasVSSO = false;
  for (let i = 1; i < confData.length; i++) {
    if (confData[i][0] === "VSSO") { hasVSSO = true; break; }
  }
  if (!hasVSSO) confSheet.appendRow(["VSSO", 1]);

  // 2. Convertir les VLI ayant subType=VSSO en catégorie VSSO
  const invSheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const invData = invSheet.getDataRange().getValues();
  let migrated = 0;
  for (let i = 1; i < invData.length; i++) {
    const cat = (invData[i][0] || "").trim();
    const sub = (invData[i][13] || "").trim();
    if (cat === "VLI" && sub === "VSSO") {
      invSheet.getRange(i + 1, 1).setValue("VSSO");   // Catégorie = VSSO
      invSheet.getRange(i + 1, 14).setValue("");       // Effacer subType
      migrated++;
    }
  }
  Logger.log("Migration VSSO: " + migrated + " items convertis de VLI+VSSO vers catégorie VSSO");
}
