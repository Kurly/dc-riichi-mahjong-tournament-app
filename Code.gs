/* ==================================================
   1. ROUTER & INITIALIZATION
   ================================================== */

function doGet(e) {
  const portal = e.parameter.portal;
  let template, title;

  if (portal === 'admin') {
    template = HtmlService.createTemplateFromFile('admin');
    title = '🏆 Tournament Admin';
  } else if (portal === 'player') {
    template = HtmlService.createTemplateFromFile('player');
    title = 'Player Portal';
  } else {
    template = HtmlService.createTemplateFromFile('index');
    title = 'Mahjong Portal';
  }

  return template.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getScriptUrl() { return ScriptApp.getService().getUrl(); }

/* ==================================================
   2. DATA AGGREGATION & CACHING
   ================================================== */

// Cache objects to prevent multiple calls to the spreadsheet in a single execution
let _cachedDataSS = null;
let _settingsMap = null;
let _cachedSettings = null;

function getDataSS() {
  if (_cachedDataSS) return _cachedDataSS;
  const master = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = master.getSheetByName("Settings");
  if (!settingsSheet) return master; 
  
  const data = settingsSheet.getDataRange().getValues();
  let targetId = "";
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == "Active_Tournament_ID") { 
      targetId = data[i][1]; 
      break; 
    }
  }
  
  if (targetId) {
    try { 
      _cachedDataSS = SpreadsheetApp.openById(targetId);
      return _cachedDataSS;
    } catch (e) { return master; }
  }
  return master;
}

function getInitialAdminData() {
  const settings = getFullSettings();
  return {
    settings: settings,
    tournaments: getTournamentList(),
    rulesets: getUniqueRulesets(), 
    url: getSpreadsheetUrl(),
    players: getPlayers(),
    schedule: getScheduleTables(),
    pairingState: getPairingState(),
    penalties: getRecentPenalties(),
    penaltyList: getPenaltyDefinitions(), 
    scoreLog: getScoreLog(),
    allGames: getAllGamesData()
  };
}

function getScoringUpdateData() {
  SpreadsheetApp.flush();
  return {
    scoreLog: getScoreLog(),
    allGames: getAllGamesData()
  };
}

function getUniqueRulesets() {
  const master = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = master.getSheetByName("Penalties_List");
  if (!sheet) return ["Default"];
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return ["Default"];

  const headers = data[0].map(h => String(h).toLowerCase());
  const ruleIdx = headers.indexOf("ruleset");
  if (ruleIdx === -1) return ["Default"];

  let rulesets = new Set();
  for(let i = 1; i < data.length; i++) {
    if(data[i][ruleIdx]) rulesets.add(data[i][ruleIdx].toString());
  }
  return Array.from(rulesets).sort();
}

function getPenaltyDefinitions() {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Penalties_List");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  const headers = data[0].map(h => String(h).toLowerCase());
  let typeIdx = headers.indexOf("type");
  let foulIdx = headers.indexOf("foul");
  let penIdx = headers.indexOf("penalty");
  let ptIdx = headers.indexOf("point deduction");

  if (typeIdx === -1) typeIdx = 0;
  if (foulIdx === -1) foulIdx = 1;
  if (penIdx === -1) penIdx = 2;
  if (ptIdx === -1) ptIdx = 3;

  let defs = [];
  for(let i=1; i<data.length; i++) {
    if(data[i][typeIdx]) {
        defs.push({
          type: data[i][typeIdx],
          foul: data[i][foulIdx],
          penalty: data[i][penIdx],
          pointDeduction: data[i][ptIdx]
        });
    }
  }
  return defs;
}

function getAllGamesData() {
  try {
    const ss = getDataSS();
    const pairSheet = ss.getSheetByName("Pairings");
    const scoreSheet = ss.getSheetByName("Scores"); 
    if(!pairSheet) return {};
    
    let scoredMap = {};
    if (scoreSheet && scoreSheet.getLastRow() > 1) {
        const sData = scoreSheet.getDataRange().getValues();
        for (let i = 1; i < sData.length; i++) {
            let r = String(sData[i][1]);
            let t = String(sData[i][2]);
            if (!scoredMap[r]) scoredMap[r] = new Set();
            scoredMap[r].add(t);
        }
    }

    const data = pairSheet.getDataRange().getValues();
    const pMap = getPlayerMap();
    let gamesByRound = {};
    let currentRound = 0;

    for(let row of data) {
      if(!row[0]) continue; 
      let cell = row[0].toString().toUpperCase();
      if(cell.includes("ROUND")) {
        let match = cell.match(/\d+/);
        currentRound = match ? parseInt(match[0]) : 0;
        continue;
      }
      
      if(currentRound > 0) {
        let tableId = parseInt(row[0]);
        if (!isNaN(tableId)) {
          if (!gamesByRound[currentRound]) gamesByRound[currentRound] = [];
          const getP = (id) => ({ id: id, name: pMap[id] || id });
          let isScored = (scoredMap[String(currentRound)] && scoredMap[String(currentRound)].has(String(tableId)));
          gamesByRound[currentRound].push({ 
              id: tableId, 
              p1: row[1] ? getP(row[1]) : { id: "?", name: "?" }, 
              p2: row[2] ? getP(row[2]) : { id: "?", name: "?" }, 
              p3: row[3] ? getP(row[3]) : { id: "?", name: "?" }, 
              p4: row[4] ? getP(row[4]) : { id: "?", name: "?" },
              isScored: isScored
          });
        }
      }
    }
    return gamesByRound;
  } catch (e) {
    console.error(e);
    return {};
  }
}

/* ==================================================
   3. DATABASE & SETTINGS
   ================================================== */

function getSpreadsheetUrl() { return getDataSS().getUrl(); }

function getTournamentList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const parents = DriveApp.getFileById(ss.getId()).getParents();
    if (!parents.hasNext()) return [];
    const folder = parents.next();
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    let list = [];
    while (files.hasNext()) {
      let f = files.next();
      if (f.getId() !== ss.getId()) list.push({ name: f.getName(), id: f.getId() });
    }
    return list.sort((a,b) => a.name.localeCompare(b.name));
  } catch (e) { return []; }
}

function switchTournament(fileId) {
  const master = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(fileId);
  updateSheetSetting(master, "Active_Tournament_ID", fileId);
  updateSheetSetting(master, "Active_Tournament_Name", file.getName());
  return "Switched to: " + file.getName();
}

function startNewTournament(tournamentName, rulesetName) {
  const master = SpreadsheetApp.getActiveSpreadsheet();
  const cleanName = tournamentName || "New Tournament " + new Date().toLocaleDateString();
  const newSS = SpreadsheetApp.create(cleanName);
  const newId = newSS.getId();
  const masterFile = DriveApp.getFileById(master.getId());
  
  if (masterFile.getParents().hasNext()) {
    DriveApp.getFileById(newId).moveTo(masterFile.getParents().next());
  }

  try {
    newSS.insertSheet("Players").appendRow(["Player ID", "Name"]);
    newSS.insertSheet("Settings").appendRow(["Key", "Value"]);
    
    // Penalty list generation
    const pList = newSS.insertSheet("Penalties_List");
    pList.appendRow(["Type", "Foul", "Penalty", "Point Deduction"]);
    const masterPList = master.getSheetByName("Penalties_List");
    if (masterPList) {
        const data = masterPList.getDataRange().getValues();
        const h = data[0].map(x => String(x).toLowerCase());
        const rIdx = h.indexOf("ruleset");
        const tIdx = h.indexOf("type");
        const fIdx = h.indexOf("foul");
        const pIdx = h.indexOf("penalty");
        const ptIdx = h.indexOf("point deduction");
        
        if (rIdx > -1 && tIdx > -1 && fIdx > -1 && pIdx > -1) {
            let rowsToAdd = [];
            for (let i = 1; i < data.length; i++) {
                if (String(data[i][rIdx]) === String(rulesetName)) {
                    let ptVal = (ptIdx > -1) ? data[i][ptIdx] : "0";
                    rowsToAdd.push([data[i][tIdx], data[i][fIdx], data[i][pIdx], ptVal]);
                }
            }
            if (rowsToAdd.length > 0) pList.getRange(2, 1, rowsToAdd.length, 4).setValues(rowsToAdd);
        }
    } else {
        pList.appendRow(["Major", "Example Penalty", "Chombo", "-20"]);
    }

    const sc = newSS.insertSheet("Scores");
    sc.appendRow(["Timestamp", "Round", "Game ID", "P1 ID", "Raw P1", "Formatted P1", "P2 ID", "Raw P2", "Formatted P2", "P3 ID", "Raw P3", "Formatted P3", "P4 ID", "Raw P4", "Formatted P4", "Leftover"]);
    
    const pen = newSS.insertSheet("Penalties");
    pen.appendRow(["Timestamp", "Player ID", "Points Deducted", "Reason", "Round", "Table", "Notes"]);
    
    newSS.insertSheet("Pairings");
    newSS.insertSheet("Leaderboard");
    const def = newSS.getSheetByName("Sheet1");
    if (def) newSS.deleteSheet(def);

    updateSheetSetting(master, "Active_Tournament_ID", newId);
    updateSheetSetting(master, "Active_Tournament_Name", cleanName);
    updateSheetSetting(newSS, "Tournament Name", cleanName);
    updateSheetSetting(newSS, "Rotation Seed", 0);
    updateSheetSetting(newSS, "Round Count", 4);
    updateSheetSetting(newSS, "Starting Points", 30000);
    updateSheetSetting(newSS, "Uma 1st", 15);
    updateSheetSetting(newSS, "Uma 2nd", 5);
    updateSheetSetting(newSS, "Tiebreaker_Rule", "split");
    
    return "Success: Created '" + cleanName + "'";
  } catch (e) { throw new Error("Setup failed: " + e.message); }
}

function updateSheetSetting(ss, key, val) {
  let sheet = ss.getSheetByName("Settings");
  if (!sheet) sheet = ss.insertSheet("Settings");
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == key) { 
      sheet.getRange(i + 1, 2).setValue(val);
      return; 
    }
  }
  sheet.appendRow([key, val]);
}

function getFullSettings() {
  if (_cachedSettings) return _cachedSettings;
  const master = SpreadsheetApp.getActiveSpreadsheet();
  const dataSS = getDataSS();
  
  const read = (ss, k, d) => {
    let sheet = ss.getSheetByName("Settings");
    if(!sheet) return d;
    const v = sheet.getDataRange().getValues();
    for(let i=0; i<v.length; i++) if(v[i][0] == k) return v[i][1] === "" ? d : v[i][1];
    return d;
  };

  _cachedSettings = {
    activeName: read(master, "Active_Tournament_Name", "No Active Tournament"),
    activeId: read(master, "Active_Tournament_ID", ""),
    uma1: read(dataSS, "Uma 1st", 15),
    uma2: read(dataSS, "Uma 2nd", 5),
    startPoints: read(dataSS, "Starting Points", 30000),
    roundCount: read(dataSS, "Round Count", 4),
    pairingMode: read(dataSS, "Pairing_Mode", "scramble"),
    topCutEnabled: read(dataSS, "Top_Cut_Enabled", "false"),
    topCutSize: read(dataSS, "Top_Cut_Size", 0),
    topCutRound: read(dataSS, "Top_Cut_Round", 0),
    tiebreakerRule: read(dataSS, "Tiebreaker_Rule", "split")
  };
  return _cachedSettings;
}

function saveTournamentSettings(form) {
  const ss = getDataSS();
  updateSheetSetting(ss, "Uma 1st", form.uma1);
  updateSheetSetting(ss, "Uma 2nd", form.uma2);
  updateSheetSetting(ss, "Starting Points", form.startPoints);
  updateSheetSetting(ss, "Round Count", form.roundCount);
  updateSheetSetting(ss, "Pairing_Mode", form.pairingMode);
  updateSheetSetting(ss, "Top_Cut_Enabled", form.topCutEnabled);
  updateSheetSetting(ss, "Top_Cut_Size", form.topCutSize);
  updateSheetSetting(ss, "Top_Cut_Round", form.topCutRound);
  updateSheetSetting(ss, "Tiebreaker_Rule", form.tiebreakerRule);
  return "Settings Saved.";
}

function readSetting(ss, key, def) {
  if (!_settingsMap) {
    _settingsMap = new Map();
    let sheet = ss.getSheetByName("Settings");
    if(sheet) {
      const data = sheet.getDataRange().getValues();
      for(let i=0; i<data.length; i++) {
        _settingsMap.set(data[i][0], data[i][1]);
      }
    }
  }
  let val = _settingsMap.get(key);
  return (val !== undefined && val !== "") ? val : def;
}

/* ==================================================
   4. PLAYER MANAGEMENT
   ================================================== */

function getNextSafeId(sheet) {
  const data = sheet.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const match = data[i][0].toString().match(/\d+/);
    if (match) { 
        let n = parseInt(match[0], 10); 
        if (n > max) max = n;
    }
  }
  return max + 1;
}

function addPlayer(name, manualId) {
  const ss = getDataSS();
  let sheet = ss.getSheetByName("Players");
  if (!sheet) { sheet = ss.insertSheet("Players"); sheet.appendRow(["Player ID", "Name"]); }
  let id = manualId || "P" + getNextSafeId(sheet);
  sheet.appendRow([id, name]);
  return getPlayers();
}

function addPlayersBulk(names) {
  const ss = getDataSS();
  let sheet = ss.getSheetByName("Players");
  if (!sheet) { sheet = ss.insertSheet("Players"); sheet.appendRow(["Player ID", "Name"]); }
  let nextNum = getNextSafeId(sheet);
  const rows = [];
  names.forEach(n => { if(n.trim()) { rows.push(["P" + nextNum, n.trim()]); nextNum++; } });
  if (rows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 2).setValues(rows);
  return getPlayers();
}

function deletePlayer(playerId) {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Players");
  if (!sheet) return getPlayers();
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0].toString() == playerId.toString()) { 
        sheet.deleteRow(i + 1);
    }
  }
  return getPlayers();
}

function togglePlayerDNF(playerId) {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Players");
  if (!sheet) return getPlayers();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() == playerId.toString()) {
      let cur = data[i][1].toString();
      let neu = cur.startsWith("[DNF] ") ? cur.replace("[DNF] ", "") : "[DNF] " + cur;
      sheet.getRange(i + 1, 2).setValue(neu);
      break;
    }
  }
  return getPlayers();
}

function getPlayers() {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Players");
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(r => ({ id: r[0], name: r[1] }));
}

function clearAllPlayers() {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Players");
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  return [];
}

function getPlayerMap() {
  const list = getPlayers();
  let map = {};
  list.forEach(p => map[p.id.toString()] = p.name);
  return map;
}

/* ==================================================
   5. PAIRING LOGIC
   ================================================== */

function getPairingState() {
  const ss = getDataSS();
  const pairSheet = ss.getSheetByName("Pairings");
  const sett = getFullSettings();
  const players = getPlayers();
  if (players.length < 4) return { error: "Need at least 4 players." };

  let maxRound = 0;
  if (pairSheet && pairSheet.getLastRow() > 1) {
    const data = pairSheet.getDataRange().getValues();
    for(let row of data) {
      if(!row[0]) continue;
      let cell = row[0].toString().toUpperCase();
      if (cell.includes("ROUND")) {
        let match = cell.match(/\d+/);
        if(match && parseInt(match[0]) > maxRound) maxRound = parseInt(match[0]);
      }
    }
  }

  let validOptions = [];
  let pCount = players.length;
  if (pCount % 4 !== 0) pCount += (4 - (pCount % 4));
  for (let b = 1; b <= pCount / 4; b++) {
    let baseSize = Math.floor((pCount / b) / 4) * 4;
    if (baseSize < 4 && b > 1) break; 
    let lastBucketSize = pCount - (baseSize * (b - 1));
    if (lastBucketSize >= 4 && lastBucketSize % 4 === 0) {
        let label = (b === 1) ?
                    `1 Bucket (${pCount})` : 
                    (baseSize === lastBucketSize) ?
                    `${b} Buckets (${baseSize} each)` : 
                    `${b} Buckets (${b-1}x${baseSize}, 1x${lastBucketSize})`;
        validOptions.push({ val: b, label: label });
    }
  }

  return {
    nextRound: maxRound + 1,
    totalRounds: sett.roundCount,
    playerCount: players.length,
    validBuckets: validOptions,
    lastBuckets: Number(readSetting(ss, "Last_Bucket_Count", 1))
  };
}

function generateNextRound(bucketCount, addSubs) {
  const ss = getDataSS();
  let pairSheet = ss.getSheetByName("Pairings");
  if (!pairSheet) pairSheet = ss.insertSheet("Pairings");
  const state = getPairingState();
  const round = state.nextRound;

  const allGames = getAllGamesData();
  const currentRound = round - 1;
  if (currentRound > 0 && allGames[currentRound]) {
    const unscored = allGames[currentRound].filter(g => !g.isScored);
    if (unscored.length > 0) {
      return { success: false, message: `Cannot generate round. ${unscored.length} table(s) are missing scores in Round ${currentRound}.` };
    }
  }

  let players = getPlayers();
  if (addSubs && players.length % 4 !== 0) {
    const subsNeeded = 4 - (players.length % 4);
    let subCount = players.filter(p => p.name.toUpperCase().startsWith("SUB")).length;
    for (let i = 0; i < subsNeeded; i++) {
      subCount++;
      addPlayer(`SUB ${subCount}`);
    }
    players = getPlayers();
  }

  if (players.length % 4 !== 0) {
    return { success: false, message: "Cannot generate round. The number of players must be a multiple of 4." };
  }

  // Optimize pairing lookup using Map of Sets
  let historyMap = new Map();
  players.forEach(p => historyMap.set(p.id, new Set()));
  if (pairSheet.getLastRow() > 1) {
    const data = pairSheet.getDataRange().getValues();
    let inData = false;
    for(let row of data) {
      if(!row[0]) continue;
      if(row[0].toString().includes("ROUND")) { inData = true; continue; }
      if(inData && row[1]) {
        let pIds = [row[1], row[2], row[3], row[4]];
        for(let i=0; i<4; i++) {
          for(let j=0; j<4; j++) {
            if(i !== j && historyMap.has(pIds[i])) {
                historyMap.get(pIds[i]).add(pIds[j]);
            }
          }
        }
      }
    }
  }

  const mode = readSetting(ss, "Pairing_Mode", "scramble");
  const cutEnabled = readSetting(ss, "Top_Cut_Enabled", "false") === "true";
  const cutSize = parseInt(readSetting(ss, "Top_Cut_Size", 0));
  const cutRound = parseInt(readSetting(ss, "Top_Cut_Round", 0));
  let isCutActive = readSetting(ss, "Top_Cut_Active", "false") === "true";
  let savedCutIDs = readSetting(ss, "Top_Cut_Player_IDs", "").split(",").filter(x => x);
  let buckets = [];
  let shouldTrigger = (cutEnabled && cutSize > 0 && !isCutActive);
  
  if (shouldTrigger && cutRound > 0) {
      if (round <= cutRound) shouldTrigger = false;
  }

  if (shouldTrigger) {
      isCutActive = true;
      const standings = getStandingsData();
      let ranked = players.map(p => {
          let s = standings.find(x => x.id === p.id);
          return { ...p, pts: s ? s.totalScore : -9999 };
      });
      ranked.sort((a,b) => b.pts - a.pts);

      let topPool = ranked.slice(0, cutSize);
      let restPool = ranked.slice(cutSize);
      
      updateSheetSetting(ss, "Top_Cut_Start_Round", round);
      updateSheetSetting(ss, "Top_Cut_Player_IDs", topPool.map(p=>p.id).join(","));
      updateSheetSetting(ss, "Top_Cut_Active", "true");
      
      if (mode === 'swiss') {
          buckets.push(topPool.sort(() => Math.random() - 0.5));
          let remBuckets = Math.max(1, bucketCount - 1);
          let total = restPool.length;
          let baseSize = Math.floor((total / remBuckets) / 4) * 4;
          let currentIdx = 0;
          for (let i = 0; i < remBuckets; i++) {
              let size = baseSize;
              if (i === remBuckets - 1) size = total - currentIdx;
              if (size > 0) {
                  buckets.push(restPool.slice(currentIdx, currentIdx + size).sort(() => Math.random() - 0.5));
                  currentIdx += size;
              }
          }
      } else {
          buckets.push(topPool.sort(() => Math.random() - 0.5));
          buckets.push(restPool.sort(() => Math.random() - 0.5));
      }
  }
  else if (isCutActive && savedCutIDs.length > 0) {
      let topPool = players.filter(p => savedCutIDs.includes(p.id));
      let restPool = players.filter(p => !savedCutIDs.includes(p.id));
      
      if (mode === 'swiss') {
          const standings = getStandingsData();
          const getStats = (pid) => standings.find(x => x.id === pid);
          topPool.sort((a,b) => {
              let sa = getStats(a.id); let sb = getStats(b.id);
              return (sb ? sb.auxScore : -9999) - (sa ? sa.auxScore : -9999);
          });
          buckets.push(topPool); 
          
          restPool.sort((a,b) => {
              let sa = getStats(a.id); let sb = getStats(b.id);
              return (sb ? sb.totalScore : -9999) - (sa ? sa.totalScore : -9999);
          });
          let remBuckets = Math.max(1, bucketCount - 1);
          let total = restPool.length;
          let baseSize = Math.floor((total / remBuckets) / 4) * 4;
          let currentIdx = 0;
          for (let i = 0; i < remBuckets; i++) {
              let size = baseSize;
              if (i === remBuckets - 1) size = total - currentIdx;
              if (size > 0) {
                  buckets.push(restPool.slice(currentIdx, currentIdx + size));
                  currentIdx += size;
              }
          }
      } else {
          buckets.push(topPool.sort(() => Math.random() - 0.5));
          buckets.push(restPool.sort(() => Math.random() - 0.5));
      }
  }
  else {
      if (mode === 'swiss') buckets = recalculateSwissBuckets(players, bucketCount);
      else buckets.push(players.sort(() => Math.random() - 0.5));
  }
  
  updateSheetSetting(ss, "Last_Bucket_Count", bucketCount);

  let roundTables = [];
  let tableCounter = 1;
  buckets.forEach((bucket, bIdx) => {
    let pool = [...bucket];
    let bucketChar = String.fromCharCode(65 + bIdx); 
    let allowRepeats = (isCutActive && bIdx === 0);

    while(pool.length >= 4) {
      let bestTable = null;
      let minRepeats = 999;
      
      if (allowRepeats) {
          bestTable = pool.slice(0, 4);
      } else {
          for(let attempt=0; attempt<1000; attempt++) {
            for (let i = pool.length - 1; i > 0; i--) {
                const j = Math.floor(Math.random() * (i + 1));
                [pool[i], pool[j]] = [pool[j], pool[i]];
            }
            let candidates = pool.slice(0, 4);
            let repeats = countRepeats(candidates, historyMap);
            if (repeats === 0) { bestTable = candidates; minRepeats = 0; break; }
            if (repeats < minRepeats) { minRepeats = repeats; bestTable = candidates; }
          }
      }
      
      pool = pool.filter(p => !bestTable.includes(p));
      roundTables.push([tableCounter++, bestTable[0].id, bestTable[1].id, bestTable[2].id, bestTable[3].id, bucketChar]);
    }
  });

  let output = [[`--- ROUND ${round} (${mode.toUpperCase()}) ---`, "", "", "", "", ""]];
  roundTables.forEach(row => output.push(row));
  output.push(["", "", "", "", "", ""]); 
  pairSheet.getRange(pairSheet.getLastRow() + 1, 1, output.length, 6).setValues(output);
  return { success: true, message: `Generated Round ${round} pairings!` };
}

function recalculateSwissBuckets(players, bucketCount) {
    const standings = getStandingsData();
    let ranked = players.map(p => {
      let s = standings.find(x => x.id === p.id);
      return { ...p, pts: s ? s.totalScore : -9999 };
    });
    ranked.sort((a,b) => b.pts - a.pts);
    let buckets = [];
    let total = ranked.length;
    let baseSize = Math.floor((total / bucketCount) / 4) * 4;
    let currentIdx = 0;
    for (let i = 0; i < bucketCount; i++) {
      let size = (i === bucketCount - 1) ? total - currentIdx : baseSize;
      let slice = ranked.slice(currentIdx, currentIdx + size);
      currentIdx += size;
      buckets.push(slice.sort(() => Math.random() - 0.5));
    }
    return buckets;
}

function countRepeats(players, historyMap) {
  let repeats = 0;
  for(let i=0; i<players.length; i++) {
    let pid = players[i].id;
    if(pid === "BYE" || pid.startsWith("SUB")) continue;
    let past = historyMap.get(pid);
    if (!past) continue;
    
    for(let j=i+1; j<players.length; j++) {
      let pid2 = players[j].id;
      if(pid2 === "BYE" || pid2.startsWith("SUB")) continue;
      if(past.has(pid2)) repeats++;
    }
  }
  return repeats;
}

function getScheduleTables() {
  const ss = getDataSS();
  const pairSheet = ss.getSheetByName("Pairings");
  const pMap = getPlayerMap();
  if (!pairSheet) return [];
  const data = pairSheet.getDataRange().getValues();
  let schedule = [];
  let currentRoundObj = null;

  for (let row of data) {
    if (!row[0]) continue;
    let cell = row[0].toString().toUpperCase();
    if (cell.includes("ROUND")) {
      let match = cell.match(/\d+/);
      if (match) {
        currentRoundObj = { round: parseInt(match[0]), tables: [] };
        schedule.push(currentRoundObj);
      }
      continue;
    }
    if (currentRoundObj && !isNaN(parseInt(row[0]))) {
      let tId = parseInt(row[0]);
      let bucket = (row.length > 5 && row[5]) ? row[5] : "";
      let table = {
        id: tId, bucket: bucket,
        p1: pMap[row[1]] || row[1], p2: pMap[row[2]] || row[2],
        p3: pMap[row[3]] || row[3], p4: pMap[row[4]] || row[4]
      };
      currentRoundObj.tables.push(table);
    }
  }
  return schedule.reverse();
}

/* ==================================================
   6. SCORING & PENALTIES
   ================================================== */

function checkIfScored(round, gameId, p1Id, p2Id, p3Id, p4Id) {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Scores");
  if (!sheet) return { scored: false };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == round && data[i][2] == gameId) {
      const rowP1 = data[i][3]; const rowP2 = data[i][6]; const rowP3 = data[i][9]; const rowP4 = data[i][12];
      if (p1Id && (rowP1 != p1Id || rowP2 != p2Id || rowP3 != p3Id || rowP4 != p4Id)) {
        sheet.deleteRow(i + 1);
        return { scored: false, mismatchDeleted: true };
      }
      return { scored: true, rowIndex: i + 1, scores: { p1: data[i][4], p2: data[i][7], p3: data[i][10], p4: data[i][13], leftover: data[i][15] } };
    }
  }
  return { scored: false };
}

function saveScores(form) {
  const ss = getDataSS();
  let sheet = ss.getSheetByName("Scores");
  if (!sheet) sheet = ss.insertSheet("Scores");
  const settings = getFullSettings();
  const start = Number(settings.startPoints);
  let u1 = Number(settings.uma1);
  let u2 = Number(settings.uma2);

  if (Math.abs(u1) < 1000 && u1 !== 0) u1 *= 1000;
  if (Math.abs(u2) < 1000 && u2 !== 0) u2 *= 1000;

  const tieRule = settings.tiebreakerRule || 'split';
  const g = form.game;
  let pData = [ { id: g.p1Id, s: Number(g.p1Score), k: 'p1' }, { id: g.p2Id, s: Number(g.p2Score), k: 'p2' }, { id: g.p3Id, s: Number(g.p3Score), k: 'p3' }, { id: g.p4Id, s: Number(g.p4Score), k: 'p4' } ];
  const bonuses = [u1, u2, -u2, -u1];
  let res = {};
  
  // Clean grouping for Split Uma logic
  if (tieRule === 'head_bump' && form.rankedIds) {
      pData.sort((a, b) => {
         if (b.s !== a.s) return b.s - a.s; 
         return form.rankedIds.indexOf(a.id) - form.rankedIds.indexOf(b.id);
      });
      pData.forEach((p, idx) => { res[p.k] = { raw: p.s, final: ((p.s - start) + bonuses[idx]) / 1000 }; });
  } else {
      pData.sort((a, b) => b.s - a.s);
      let scoreGroups = new Map();
      pData.forEach((p, i) => {
          if (!scoreGroups.has(p.s)) scoreGroups.set(p.s, { players: [], totalBonus: 0 });
          let group = scoreGroups.get(p.s);
          group.players.push(p);
          group.totalBonus += bonuses[i];
      });
      
      scoreGroups.forEach((group, score) => {
          let avgBonus = group.totalBonus / group.players.length;
          group.players.forEach(p => {
              res[p.k] = { raw: p.s, final: ((p.s - start) + avgBonus) / 1000 };
          });
      });
  }

  const check = checkIfScored(form.round, g.gameId);
  const rowData = [ new Date(), form.round, g.gameId, g.p1Id, res.p1.raw, res.p1.final, g.p2Id, res.p2.raw, res.p2.final, g.p3Id, res.p3.raw, res.p3.final, g.p4Id, res.p4.raw, res.p4.final, g.leftoverScore ];
  
  if (check.scored && check.rowIndex) { sheet.getRange(check.rowIndex, 1, 1, rowData.length).setValues([rowData]); }
  else { sheet.appendRow(rowData); }
  
  SpreadsheetApp.flush(); 
  updateLeaderboardSheet();
  return { success: true, message: check.scored ? "Updated existing score." : "Saved new score." };
}

function addPenalty(round, table, playerId, points, reason, notes) {
  const ss = getDataSS();
  let sheet = ss.getSheetByName("Penalties");
  if (!sheet) { 
      sheet = ss.insertSheet("Penalties");
      sheet.appendRow(["Timestamp", "Player ID", "Points Deducted", "Reason", "Round", "Table", "Notes"]); 
  }
  
  let pts = Number(points);
  if (isNaN(pts)) pts = 0; 
  if (Math.abs(pts) < 1000 && pts !== 0) pts *= 1000;

  sheet.appendRow([new Date(), playerId, pts, reason, round, table, notes]);
  updateLeaderboardSheet();
  return { success: true, message: "Penalty Added." };
}

function getRecentPenalties() {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Penalties");
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const pMap = getPlayerMap();
  return data.slice(1).reverse().map(r => {
    let dateStr = "N/A";
    try { if (r[0]) dateStr = Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), "MM/dd HH:mm"); } catch(e) { dateStr = r[0].toString(); }
    return { 
        date: dateStr, 
        name: pMap[r[1]] || r[1], 
        points: r[2], 
        reason: r[3], 
        round: (r[4] || "-"), 
        table: (r[5] || "-"),
        notes: (r[6] || "") 
    };
  });
}

function getScoreLog() {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Scores");
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const pMap = getPlayerMap();
  return data.slice(1).reverse().map(r => {
    let dateStr = "N/A";
    try { if (r[0]) dateStr = Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), "MM/dd HH:mm"); } catch(e) { dateStr = r[0].toString(); }
    return { date: dateStr, round: r[1], game: r[2], p1: `${pMap[r[3]] || r[3]} (${r[4]})`, p2: `${pMap[r[6]] || r[6]} (${r[7]})`, p3: `${pMap[r[9]] || r[9]} (${r[10]})`, p4: `${pMap[r[12]] || r[12]} (${r[13]})`, leftover: r[15] };
  });
}

/* ==================================================
   7. LEADERBOARD
   ================================================== */

function getStandingsData() {
  const ss = getDataSS();
  const sSheet = ss.getSheetByName("Scores");
  const pSheet = ss.getSheetByName("Penalties");
  const pMap = getPlayerMap();
  
  let startRound = parseInt(readSetting(ss, "Top_Cut_Start_Round", "0"));
  let topIDsRaw = String(readSetting(ss, "Top_Cut_Player_IDs", ""));
  let topIDs = topIDsRaw ? topIDsRaw.split(",").filter(x => x) : [];
  let topSet = new Set(topIDs);
  let hasCut = (startRound > 0 && topIDs.length > 0);

  let stats = {};
  Object.keys(pMap).forEach(id => { 
      stats[id] = { 
          id: id, name: pMap[id], 
          totalPts: 0, postCutPts: 0, preCutPts: 0,
          played: 0, pen: 0, 
          isDNF: pMap[id].startsWith("[DNF]"),
          isTopCut: topSet.has(id)
      }; 
  });
  if(sSheet && sSheet.getLastRow() > 1) {
    const data = sSheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
      if (!data[i][1]) continue;
      let rNum = parseInt(data[i][1]);
      [3, 6, 9, 12].forEach((idIdx) => {
        const pid = data[i][idIdx]; 
        const pts = Number(data[i][idIdx + 2]); 
        if(pid && stats[pid]) { 
            stats[pid].totalPts += pts;
            stats[pid].played++;
            if (hasCut && rNum >= startRound) stats[pid].postCutPts += pts;
        }
      });
    }
  }
  
  if(pSheet && pSheet.getLastRow() > 1) {
    const pData = pSheet.getDataRange().getValues();
    for(let i=1; i<pData.length; i++) {
      if (!pData[i][1]) continue;
      let rNum = parseInt(pData[i][4]);
      const pid = pData[i][1]; 
      const deductFmt = Number(pData[i][2]) / 1000;
      if(pid && stats[pid]) { 
          stats[pid].totalPts -= deductFmt; 
          stats[pid].pen += deductFmt;
          if (hasCut && rNum >= startRound) stats[pid].postCutPts -= deductFmt;
      }
    }
  }

  Object.values(stats).forEach(p => {
      p.preCutPts = p.totalPts - p.postCutPts;
      p.sortScore = p.isTopCut ? p.postCutPts : p.totalPts; 
  });
  
  let topGroup = Object.values(stats).filter(p => p.isTopCut).sort((a,b) => b.sortScore - a.sortScore);
  let restGroup = Object.values(stats).filter(p => !p.isTopCut).sort((a,b) => {
      if (a.isDNF !== b.isDNF) return a.isDNF ? 1 : -1;
      return b.sortScore - a.sortScore;
  });
  
  const formatP = (p, rank) => ({
      rank: p.isDNF ? "-" : rank,
      id: p.id, name: p.name,
      displayScore: p.totalPts, auxScore: p.postCutPts, totalScore: p.totalPts,
      played: p.played, penalties: p.pen,
      isDNF: p.isDNF, isTopCut: p.isTopCut
  });
  
  return [
      ...topGroup.map((p, i) => formatP(p, i+1)),
      ...restGroup.map((p, i) => formatP(p, topGroup.length + i + 1))
  ];
}

function updateLeaderboardSheet() {
  const ss = getDataSS();
  let sheet = ss.getSheetByName("Leaderboard");
  if (!sheet) sheet = ss.insertSheet("Leaderboard");
  const standings = getStandingsData();
  const rows = standings.map(p => [ p.rank, p.id, p.name, p.played, p.displayScore ]);
  
  sheet.clear();
  sheet.appendRow(["Rank", "Player ID", "Name", "Games Played", "Total Points"]);
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 5).setValues(rows);
}

/* ==================================================
   8. SWAP & EDITING TOOLS
   ================================================== */

function getHistoryMatrix(excludeRound) {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Pairings");
  let historyMap = new Map();
  if (!sheet) return historyMap;
  const data = sheet.getDataRange().getValues();
  let currentRound = 0;
  for (let row of data) {
    if(!row[0]) continue;
    let cell = row[0].toString().toUpperCase();
    if (cell.includes("ROUND")) {
      let match = cell.match(/\d+/);
      currentRound = match ? parseInt(match[0]) : 0;
      continue;
    }
    if (currentRound >= excludeRound) continue;
    if (currentRound > 0 && row[1]) {
      let pIds = [row[1], row[2], row[3], row[4]];
      for (let i = 0; i < 4; i++) {
        for (let j = 0; j < 4; j++) {
          if (i !== j) {
            let p1 = pIds[i];
            let p2 = pIds[j];
            if (!historyMap.has(p1)) historyMap.set(p1, new Set());
            historyMap.get(p1).add(p2);
          }
        }
      }
    }
  }
  return historyMap;
}

function swapPairings(round, t1Id, p1Id, t2Id, p2Id, force) {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Pairings");
  const data = sheet.getDataRange().getValues();
  const pMap = getPlayerMap();
  let rIdx1 = -1, rIdx2 = -1; let row1, row2;
  let inRound = false;

  for (let i = 0; i < data.length; i++) {
    if(!data[i][0]) continue;
    let cell = data[i][0].toString().toUpperCase();
    if (cell.includes("ROUND")) {
      let match = cell.match(/\d+/);
      inRound = (match && parseInt(match[0]) == round);
      continue;
    }
    if (inRound) {
      if (data[i][0] == t1Id) { rIdx1 = i; row1 = data[i]; }
      if (data[i][0] == t2Id) { rIdx2 = i; row2 = data[i]; }
    }
  }

  if (rIdx1 === -1 || rIdx2 === -1) return { success: false, message: "Tables not found." };
  let cIdx1 = row1.indexOf(p1Id); let cIdx2 = row2.indexOf(p2Id);
  if (cIdx1 < 1 || cIdx2 < 1) return { success: false, message: "Player positions changed. Refresh and try again." };

  if (!force) {
    const histMap = getHistoryMatrix(round);
    let conflicts = [];
    for (let k = 1; k <= 4; k++) {
      let opp = row2[k];
      if (opp !== p2Id && opp !== "" && histMap.has(p1Id) && histMap.get(p1Id).has(opp)) conflicts.push(`${pMap[p1Id] || p1Id} played ${pMap[opp] || opp}`);
    }
    for (let k = 1; k <= 4; k++) {
      let opp = row1[k];
      if (opp !== p1Id && opp !== "" && histMap.has(p2Id) && histMap.get(p2Id).has(opp)) conflicts.push(`${pMap[p2Id] || p2Id} played ${pMap[opp] || opp}`);
    }
    if (conflicts.length > 0) return { success: false, warning: true, message: "⚠️ Conflict Warning:\n" + conflicts.join("\n") + "\n\nSwap anyway?" };
  }

  sheet.getRange(rIdx1 + 1, cIdx1 + 1).setValue(p2Id);
  sheet.getRange(rIdx2 + 1, cIdx2 + 1).setValue(p1Id);
  return { success: true, message: "✅ Players swapped successfully." };
}

function getPlayerScheduleMatrix() {
  const ss = getDataSS();
  const pairSheet = ss.getSheetByName("Pairings");
  const scoreSheet = ss.getSheetByName("Scores");
  const pMap = getPlayerMap();
  
  let players = {};
  Object.keys(pMap).forEach(k => { 
      players[k] = { id: k, name: pMap[k], tables: {}, scores: {} }; 
  });
  let maxRound = 0;

  if (pairSheet && pairSheet.getLastRow() > 1) {
    const data = pairSheet.getDataRange().getValues();
    let curRound = 0;
    for (let row of data) {
      if(!row[0]) continue;
      let cell = row[0].toString().toUpperCase();
      if (cell.includes("ROUND")) {
        let match = cell.match(/\d+/);
        if (match) {
            curRound = parseInt(match[0]);
            if(curRound > maxRound) maxRound = curRound;
        }
        continue;
      }
      if (curRound > 0 && !isNaN(parseInt(row[0]))) {
        let tId = parseInt(row[0]);
        [1, 2, 3, 4].forEach(c => {
            let pid = row[c];
            if (pid && players[pid]) {
                players[pid].tables[curRound] = tId;
            }
        });
      }
    }
  }

  if (scoreSheet && scoreSheet.getLastRow() > 1) {
    const data = scoreSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      let r = data[i][1];
      [3, 6, 9, 12].forEach(idx => {
          let pid = data[i][idx];
          let val = data[i][idx + 2]; 
          if (pid && players[pid]) {
              players[pid].scores[r] = val;
          }
      });
    }
  }

  let list = Object.values(players).sort((a,b) => {
      let na = parseInt(a.id.replace(/\D/g, '')) || 0;
      let nb = parseInt(b.id.replace(/\D/g, '')) || 0;
      return na - nb;
  });
  return { maxRound: maxRound, players: list };
}