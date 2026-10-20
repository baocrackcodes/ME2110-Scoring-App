// --- 1. UI SETUP ---
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ ME 2110 Admin')
    .addItem('0. Install Pirate Scores Sheet (one time)', 'setupPirateScoresSheet')
    .addItem('1. Generate Round of 32 (From R1/R2 Scores)', 'generateRoundOf32')
    .addSeparator()
    .addItem('2. Advance Winners -> Sweet 16', 'advanceToSweet16')
    .addItem('3. Advance Winners -> Elite 8', 'advanceToElite8')
    .addItem('4. Advance Winners -> Final 4', 'advanceToFinal4')
    .addSeparator()
    .addItem('⚠️ RESET ALL COMPETITION DATA', 'clearAllData')
    .addToUi();
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ME 2110 Scoring App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1'); 
}

function shuffleArray(array) {
  for (var i = array.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = array[i]; array[i] = array[j]; array[j] = temp;
  }
  return array;
}

// New competition schema (zero-indexed columns in the Scores sheet).
var PIRATE_SCORE_HEADERS = [
  'Timestamp', 'Round', 'Track', 'Heat', 'Color', 'Section', 'Team ID', 'Team Name',
  'Activated', 'Left Dock', 'Gold On Mound', 'Gold Surrounding',
  'Jewels On Mound', 'Jewels Surrounding', 'Flag Rank', 'Black Cannonballs',
  'White Parcels', 'Mermaids Toppled', 'Mermaids Moved', 'Davy Jones Outcome',
  'Total Score', 'DQ?', 'DQ Reason'
];

function getColMap(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var index = {};
  headers.forEach(function(h, i) { index[String(h).trim().toLowerCase()] = i; });
  var cols = {};
  var keys = { round:'Round', track:'Track', heat:'Heat', teamId:'Team ID',
    score:'Total Score', isDq:'DQ?', goldTop:'Gold On Mound', goldZone:'Gold Surrounding',
    jewelTop:'Jewels On Mound', jewelZone:'Jewels Surrounding', flag:'Flag Rank',
    black:'Black Cannonballs', white:'White Parcels', merTop:'Mermaids Toppled',
    merMoved:'Mermaids Moved', davy:'Davy Jones Outcome' };
  Object.keys(keys).forEach(function(k) {
    var col = index[keys[k].toLowerCase()];
    if (col === undefined) throw new Error('Scores sheet missing header: ' + keys[k] + '. Run setupPirateScoresSheet() first.');
    cols[k] = col;
  });
  return cols;
}

// Run ONCE after backing up/archiving existing scores. Does not delete historical rows.
// Creates a fresh Scores sheet with the pirate schema, moving the old sheet to an archive.
function setupPirateScoresSheet() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var old = ss.getSheetByName('Scores');
  if (!old) throw new Error('Scores tab is missing.');
  var existing = old.getRange(1,1,1,Math.max(old.getLastColumn(),1)).getValues()[0];
  if (String(existing[10]).trim() === 'Gold On Mound') {
    ui.alert('Pirate scoring schema is already installed. No changes made.');
    return;
  }
  var answer = ui.alert('Install new scoring schema',
    'This will archive the old Scores tab, create a NEW empty Scores tab, and clear the Live Leaderboard. Historical scores will NOT be migrated to the new format. Continue?',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  var archiveName = 'Scores Archive ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd-HHmmss');
  old.setName(archiveName);
  try {
    var sheet = ss.insertSheet('Scores');
    sheet.getRange(1,1,1,PIRATE_SCORE_HEADERS.length).setValues([PIRATE_SCORE_HEADERS]);
    sheet.setFrozenRows(1);
    var lb = ss.getSheetByName('Live Leaderboard');
    if (lb && lb.getLastRow() > 1) lb.getRange(2,1,lb.getLastRow()-1,4).clearContent();
    updateLiveLeaderboard();
  } catch (error) {
    // Avoid leaving the primary sheet without its expected name.
    if (!ss.getSheetByName('Scores')) old.setName('Scores');
    throw error;
  }
  ui.alert('New pirate Scores tab created. Prior scores archived in: ' + archiveName);
}

function getSheetNameFromRound(roundStr) {
  if (roundStr.indexOf("Round 1") > -1 || roundStr.indexOf("Round 2") > -1) return "Matches";
  if (roundStr.indexOf("32") > -1) return "Round of 32";
  if (roundStr.indexOf("16") > -1) return "Sweet 16";
  if (roundStr.indexOf("8") > -1) return "Elite 8";
  if (roundStr.indexOf("4") > -1) return "Final 4";
  return "Matches"; 
}

function getRoundNumber(roundStr) {
  if (roundStr.indexOf("32") > -1) return 32;
  if (roundStr.indexOf("16") > -1) return 16;
  if (roundStr.indexOf("8") > -1) return 8;
  if (roundStr.indexOf("4") > -1) return 4;
  return 1; 
}

function isValidTeam(id) {
  if (!id) return false;
  var clean = String(id).trim().toUpperCase();
  return clean !== "" && clean !== "NONE" && clean !== "TBD" && clean !== "EMPTY" && clean !== "UNDEFINED";
}

function getPlayoffTiebreakerWeights(scoreRow, map) {
  var flagPoints = {None:0, First:40, Second:28, Third:17, Fourth:7};
  var davyPoints = {None:0, ChestShip:11, ChestDock:32, HeartShip:39, HeartDock:61};
  return {
    tier1: davyPoints[scoreRow[map.davy]] || 0,
    tier2: flagPoints[scoreRow[map.flag]] || 0,
    tier3: (Number(scoreRow[map.merTop]) * 12) + (Number(scoreRow[map.merMoved]) * 5),
    tier4: (Number(scoreRow[map.black]) * 9) - (Number(scoreRow[map.white]) * 10),
    tier5: (Number(scoreRow[map.goldTop]) * 10) + (Number(scoreRow[map.goldZone]) * 4) +
           (Number(scoreRow[map.jewelTop]) * 15) + (Number(scoreRow[map.jewelZone]) * 9)
  };
}

// ==========================================
// THE NEW SINGLE SOURCE OF TRUTH LEADERBOARD
// ==========================================
function updateLiveLeaderboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scoresSheet = ss.getSheetByName("Scores");
  var teamSheet = ss.getSheetByName("Teams");
  var lbSheet = ss.getSheetByName("Live Leaderboard");
  
  if (!scoresSheet || !teamSheet || !lbSheet) return;

  var scoreData = scoresSheet.getDataRange().getValues();
  var teamData = teamSheet.getDataRange().getValues();
  var map = getColMap(scoresSheet);
  
  var teamDict = {};
  var sprint3Scores = {};
  
  for(var i=1; i<teamData.length; i++) {
    if(teamData[i][0]) {
      var cId = String(teamData[i][0]).trim().toUpperCase();
      teamDict[cId] = String(teamData[i][1]).trim();
      sprint3Scores[cId] = parseInt(teamData[i][2]) || 0;
    }
  }
  
  var bestRuns = {};
  for(var j=1; j<scoreData.length; j++) {
    var roundStr = String(scoreData[j][map.round]).toUpperCase(); 
    if(roundStr.indexOf("ROUND 1") > -1 || roundStr.indexOf("ROUND 2") > -1 || roundStr.indexOf("FLEX") > -1) {
      var tId = String(scoreData[j][map.teamId]).trim().toUpperCase(); 
      if(isValidTeam(tId)) {
        var rawTotalPts = parseInt(scoreData[j][map.score]); 
        if(isNaN(rawTotalPts)) rawTotalPts = 0; 
        
        var isDQ = scoreData[j][map.isDq] === true || String(scoreData[j][map.isDq]).toUpperCase() === "TRUE"; 
        var displayPts = isDQ ? -200 : rawTotalPts; 
          
        if(!bestRuns[tId] || displayPts > bestRuns[tId].max) {
           bestRuns[tId] = { max: displayPts, sprint3: sprint3Scores[tId] || 0 };
        } 
      }
    }
  }
  
  var leaderboard = [];
  for(var teamId in teamDict) {
     if(isValidTeam(teamId)) {
         if(bestRuns[teamId]) {
            leaderboard.push({ id: teamId, name: teamDict[teamId], data: bestRuns[teamId] });
         } else {
            leaderboard.push({ id: teamId, name: teamDict[teamId], data: { max: -1, sprint3: sprint3Scores[teamId] || 0 } });
         }
     }
  }
  
  // Valid qualifying runs first, never-scored teams next, DQ-only teams last.
  // -1 is the no-score sentinel and must not be compared as an ordinary score.
  leaderboard.sort(function(a, b) {
    function category(x) { return x === -200 ? 0 : x === -1 ? 1 : 2; }
    var ca = category(a.data.max), cb = category(b.data.max);
    if (ca !== cb) return cb - ca;
    if (ca === 2 && b.data.max !== a.data.max) return b.data.max - a.data.max;
    return b.data.sprint3 - a.data.sprint3;
  });
  
  // Wipe old data and write the new hard values
  if (lbSheet.getLastRow() > 1) {
      lbSheet.getRange(2, 1, lbSheet.getLastRow() - 1, 4).clearContent();
  }
  
  var outputData = [];
  for (var rank = 0; rank < leaderboard.length; rank++) {
      var t = leaderboard[rank];
      outputData.push([rank + 1, t.id, t.name, t.data.max]);
  }
  
  if (outputData.length > 0) {
      lbSheet.getRange(2, 1, outputData.length, 4).setValues(outputData);
  }
}

// ==========================================
// CORE SAVING & PLAYOFF ENGINE
// ==========================================
function saveScoresToSheet(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Scores');
  if (!sheet) throw new Error('Scores sheet missing.');
  var map = getColMap(sheet); // Validates new schema before writing.
  if (!payload || !payload.teams || payload.teams.length !== 4) throw new Error('Four team zones are required.');

  // Verify that standard qualifying submissions correspond to an actual scheduled
  // Round/Heat/Track and its assigned team IDs. FLEX and playoffs remain unchanged.
  var roundLabel = String(payload.round || '').trim().toUpperCase();
  if (roundLabel === 'ROUND 1' || roundLabel === 'ROUND 2') {
    var requestedHeat = parseInt(String(payload.heat).replace(/\D/g, ''), 10);
    var requestedTrack = parseInt(String(payload.track).replace(/\D/g, ''), 10);
    if (!(requestedHeat >= 1 && requestedHeat <= 6 && requestedTrack >= 1 && requestedTrack <= 4) ||
        (requestedHeat === 6 && requestedTrack !== 1)) {
      throw new Error('Invalid qualifying heat/track for 84-team schedule.');
    }
    var matchesSheet = ss.getSheetByName('Matches');
    if (!matchesSheet) throw new Error('Matches sheet missing.');
    var matches = matchesSheet.getDataRange().getValues();
    var scheduled = null;
    for (var mi = 1; mi < matches.length; mi++) {
      var mr = String(matches[mi][0]).trim().toUpperCase();
      var mh = parseInt(String(matches[mi][1]).replace(/\D/g, ''), 10);
      var mt = parseInt(String(matches[mi][2]).replace(/\D/g, ''), 10);
      if (mr === roundLabel && mh === requestedHeat && mt === requestedTrack) {
        if (scheduled) throw new Error('Duplicate Round/Heat/Track in Matches tab.');
        scheduled = matches[mi];
      }
    }
    if (!scheduled) throw new Error('Selected qualifying match is missing from Matches tab.');
    var expected = [3, 4, 5, 6].map(function(col) { return String(scheduled[col] || '').trim().toUpperCase(); });
    var receivedByColor = {};
    payload.teams.forEach(function(team) { receivedByColor[String(team.color).toUpperCase()] = String(team.teamId || '').trim().toUpperCase(); });
    ['YELLOW', 'BLACK', 'BLUE', 'WHITE'].forEach(function(color, i) {
      if ((receivedByColor[color] || '') !== expected[i]) {
        throw new Error('Roster mismatch on ' + color + ': reload the scheduled match before submitting.');
      }
    });
  }

  // Server-side validation is essential: browser-side validation can be bypassed.
  var count = {gold:0, jewel:0, black:0, white:0};
  var flags = {};
  var flagPoints = {None:0, First:40, Second:28, Third:17, Fourth:7};
  var davyPoints = {None:0, ChestShip:11, ChestDock:32, HeartShip:39, HeartDock:61};
  var cleanTeams = [];
  function amount(v, max, label) {
    var n = Number(v);
    if (!Number.isInteger(n) || n < 0 || n > max) throw new Error('Invalid ' + label + ': ' + v);
    return n;
  }
  payload.teams.forEach(function(t) {
    if (!isValidTeam(t.teamId)) return;
    var goldTop = amount(t.goldTop,10,'gold on mound');
    var goldZone = amount(t.goldZone,10,'gold surrounding');
    var jewelTop = amount(t.jewelTop,5,'jewels on mound');
    var jewelZone = amount(t.jewelZone,5,'jewels surrounding');
    var black = amount(t.blackCannonballs,16,'black cannonballs');
    var white = amount(t.whiteParcels,4,'white parcels');
    var merTop = amount(t.mermaidsToppled,2,'mermaids toppled');
    var merMoved = amount(t.mermaidsMoved,2,'mermaids moved');
    if (merTop + merMoved > 2) throw new Error('Each team has only 2 mermaids.');
    if (!(t.flagRank in flagPoints)) throw new Error('Invalid flag rank.');
    if (!(t.davyOutcome in davyPoints)) throw new Error('Invalid Davy Jones outcome.');
    if (t.flagRank !== 'None') {
      if (flags[t.flagRank]) throw new Error('A flag rank was assigned to multiple teams.');
      flags[t.flagRank] = true;
    }
    count.gold += goldTop + goldZone; count.jewel += jewelTop + jewelZone;
    count.black += black; count.white += white;
    var activated = t.activate === true || String(t.activate).toUpperCase() === 'TRUE';
    var leftDock = t.leaveDock === true || String(t.leaveDock).toUpperCase() === 'TRUE';
    if (leftDock && !activated) throw new Error('A team cannot leave dock without activating.');
    var dq = t.isDQ === true || String(t.isDQ).toUpperCase() === 'TRUE';
    var score = dq ? -200 : ((activated ? 1 : 0) + (leftDock ? 10 : 0) +
      goldTop*10 + goldZone*4 + jewelTop*15 + jewelZone*9 + flagPoints[t.flagRank] +
      black*9 - white*10 + merTop*12 + merMoved*5 + davyPoints[t.davyOutcome]);
    cleanTeams.push({color:String(t.color).toUpperCase(),teamId:String(t.teamId),section:String(t.section||''),
      name:String(t.name||''),activate:activated,leaveDock:leftDock,goldTop:goldTop,goldZone:goldZone,
      jewelTop:jewelTop,jewelZone:jewelZone,flagRank:t.flagRank,blackCannonballs:black,whiteParcels:white,
      mermaidsToppled:merTop,mermaidsMoved:merMoved,davyOutcome:t.davyOutcome,score:score,isDQ:dq,
      dqReason:String(t.dqReason||'')});
  });
  if (count.gold>10 || count.jewel>5 || count.black>16 || count.white>4) {
    throw new Error('Shared resource limit exceeded: gold 10, jewels 5, black cannonballs 16, white parcels 4.');
  }

  // Rejected submissions never create partial rows.
  var stamp = new Date();
  var rows = cleanTeams.map(function(t) { return [stamp,String(payload.round),String(payload.track),String(payload.heat),
    t.color,t.section,t.teamId,t.name,t.activate,t.leaveDock,t.goldTop,t.goldZone,t.jewelTop,t.jewelZone,
    t.flagRank,t.blackCannonballs,t.whiteParcels,t.mermaidsToppled,t.mermaidsMoved,t.davyOutcome,
    t.score,t.isDQ,t.dqReason]; });
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    if(rows.length) sheet.getRange(sheet.getLastRow()+1,1,rows.length,PIRATE_SCORE_HEADERS.length).setValues(rows);
    var targetRoundNum = getRoundNumber(payload.round);
    if (targetRoundNum === 1 || String(payload.round).toUpperCase().includes('FLEX')) updateLiveLeaderboard();
    if (targetRoundNum >= 4) {
      var bracketSheet = ss.getSheetByName(getSheetNameFromRound(payload.round));
      if (bracketSheet) {
        var data = bracketSheet.getDataRange().getValues();
        var track = parseInt(String(payload.track).replace(/\D/g,''),10);
        var heat = parseInt(String(payload.heat).replace(/\D/g,''),10);
        for (var r=1;r<data.length;r++) {
          if (Number(data[r][0])!==targetRoundNum || Number(data[r][1])!==heat || Number(data[r][2])!==track) continue;
          var ranking = cleanTeams.map(function(t) {
            var row = new Array(PIRATE_SCORE_HEADERS.length).fill(0);
            row[map.goldTop]=t.goldTop;row[map.goldZone]=t.goldZone;
            row[map.jewelTop]=t.jewelTop;row[map.jewelZone]=t.jewelZone;
            row[map.flag]=t.flagRank;row[map.black]=t.blackCannonballs;row[map.white]=t.whiteParcels;
            row[map.merTop]=t.mermaidsToppled;row[map.merMoved]=t.mermaidsMoved;row[map.davy]=t.davyOutcome;
            return {color:t.color,val:t.isDQ?Number.NEGATIVE_INFINITY:t.score,
              weights:getPlayoffTiebreakerWeights(row,map),display:t.isDQ?'DQ'+(t.dqReason?' - '+t.dqReason:''):String(t.score)};
          });
          ranking.sort(function(a,b) {
            if (b.val!==a.val) return b.val-a.val;
            for (var tier=1;tier<=5;tier++) {
              var k='tier'+tier;
              if(b.weights[k]!==a.weights[k]) return b.weights[k]-a.weights[k];
            }
            return 0;
          });
          var byColor = {YELLOW:'',BLACK:'',BLUE:'',WHITE:''};
          ranking.forEach(function(t,i) {
            var winner = i < (targetRoundNum===4?1:2) && t.val!==Number.NEGATIVE_INFINITY;
            byColor[t.color]=t.display+(winner?(targetRoundNum===4?' 🏆':' ✅'):'');
          });
          bracketSheet.getRange(r+1,8,1,5).setValues([[byColor.YELLOW,byColor.BLACK,byColor.BLUE,byColor.WHITE,'Complete']]);
          break;
        }
      }
    }
  } finally { lock.releaseLock(); }
  return true;
}

function getCompletedTracks(roundStr, heatStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scoresSheet = ss.getSheetByName("Scores");
  if (!scoresSheet) return [];
  
  var map = getColMap(scoresSheet);
  var data = scoresSheet.getDataRange().getValues();
  var completed = {};
  
  var targetRound = String(roundStr).trim().toUpperCase();
  var targetHeatNum = parseInt(String(heatStr).replace(/\D/g, ''));

  for (var i = 1; i < data.length; i++) {
    var rowRound = String(data[i][map.round]).trim().toUpperCase();
    var rowRawHeat = String(data[i][map.heat]).toUpperCase().replace(/HEAT/g, '').trim();
    var rowRawTrack = String(data[i][map.track]).toUpperCase().replace(/TRACK/g, '').replace(/\(COMPLETED\)/g, '').trim();
    var rowHeatNum = parseInt(rowRawHeat);
    var rowTrackNum = parseInt(rowRawTrack);

    var roundMatch = false;
    if(targetRound.includes("FLEX") && rowRound === targetRound) roundMatch = true;
    else if(!targetRound.includes("FLEX") && !rowRound.includes("FLEX")) {
       var r1 = parseInt(targetRound.replace(/\D/g, ''));
       var r2 = parseInt(rowRound.replace(/\D/g, ''));
       if(r1 === r2) roundMatch = true;
    }
    if (roundMatch && rowHeatNum === targetHeatNum && !isNaN(rowTrackNum)) completed[String(rowTrackNum)] = true; 
  }
  return Object.keys(completed); 
}

function getMatchData(round, heat, track) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheetName = getSheetNameFromRound(round);
  var matchSheet = ss.getSheetByName(targetSheetName);
  if(!matchSheet) return null; 
  
  var matchData = matchSheet.getDataRange().getValues();
  var teamSheet = ss.getSheetByName("Teams");
  var teamData = teamSheet.getDataRange().getValues();

  var teamDict = {};
  for(var i=1; i<teamData.length; i++) {
    if(teamData[i][0]) teamDict[String(teamData[i][0]).trim().toUpperCase()] = teamData[i][1];
  }

  var targetRoundNum = parseInt(String(round).replace(/\D/g, ''));
  var numTrack = parseInt(String(track).replace(/\D/g, ''));
  var numHeat = parseInt(String(heat).replace(/\D/g, ''));

  for(var j=1; j<matchData.length; j++) {
    var rawSheetRound = String(matchData[j][0]).trim().toUpperCase();
    var sheetRoundNum = parseInt(rawSheetRound.replace(/\D/g, ''));

    var matchRoundOk = false;
    if (targetSheetName === "Matches") {
       if(sheetRoundNum === targetRoundNum && !rawSheetRound.includes("FLEX")) matchRoundOk = true;
    } else {
       if(sheetRoundNum === targetRoundNum) matchRoundOk = true;
    }

    var sheetHeatNum = parseInt(String(matchData[j][1]).toUpperCase().replace(/HEAT/g, '').trim());
    var sheetTrackNum = parseInt(String(matchData[j][2]).toUpperCase().replace(/TRACK/g, '').trim());

    if(matchRoundOk && sheetHeatNum === numHeat && sheetTrackNum === numTrack) {
       var yId = String(matchData[j][3] || "").trim().toUpperCase();
       var bkId = String(matchData[j][4] || "").trim().toUpperCase();
       var blId = String(matchData[j][5] || "").trim().toUpperCase();
       var wId = String(matchData[j][6] || "").trim().toUpperCase();

       return {
         yellow: {id: matchData[j][3] || "None", name: teamDict[yId] || "Empty"},
         black:  {id: matchData[j][4] || "None", name: teamDict[bkId] || "Empty"},
         blue:   {id: matchData[j][5] || "None", name: teamDict[blId] || "Empty"},
         white:  {id: matchData[j][6] || "None", name: teamDict[wId] || "Empty"}
       };
    }
  }
  return null;
}

function generateRoundOf32() {
  updateLiveLeaderboard(); // Force a fresh sync before building
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lbSheet = ss.getSheetByName("Live Leaderboard");
  var r32Sheet = ss.getSheetByName("Round of 32");
  
  var lbData = lbSheet.getDataRange().getValues();
  var top32 = [];
  
  for(var m=1; m<=32; m++) {
      if (lbData[m] && lbData[m][1]) top32.push(String(lbData[m][1]));
      else top32.push("TBD");
  }
  
  if(r32Sheet.getLastRow() > 1) r32Sheet.getRange(2, 1, r32Sheet.getLastRow()-1, 12).clearContent();

  var r32Matches = [
    [top32[0], top32[15], top32[16], top32[31]], [top32[1], top32[14], top32[17], top32[30]], 
    [top32[2], top32[13], top32[18], top32[29]], [top32[3], top32[12], top32[19], top32[28]], 
    [top32[4], top32[11], top32[20], top32[27]], [top32[5], top32[10], top32[21], top32[26]], 
    [top32[6], top32[9],  top32[22], top32[25]], [top32[7], top32[8],  top32[23], top32[24]]
  ];
  
  var h32 = 1, t32 = 1;
  for(var k=0; k<r32Matches.length; k++) {
    var group = shuffleArray(r32Matches[k]); 
    r32Sheet.appendRow([32, h32, t32, group[0], group[1], group[2], group[3], "", "", "", "", "Pending"]);
    t32++; if(t32 > 4) { t32 = 1; h32++; }
  }
  SpreadsheetApp.getUi().alert("Round of 32 generated successfully directly from Live Leaderboard!");
}

function advanceToSweet16() { advanceRound("Round of 32", "Sweet 16", 16, 4); }
function advanceToElite8()  { advanceRound("Sweet 16", "Elite 8", 8, 2); }
function advanceToFinal4()  { advanceRound("Elite 8", "Final 4", 4, 1); }

function advanceRound(sourceSheetName, targetSheetName, targetRoundNum, numTargetMatches) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(sourceSheetName);
  var targetSheet = ss.getSheetByName(targetSheetName);

  var sourceData = sourceSheet.getDataRange().getValues();
  var advancingPool = [];

  for(var j=1; j<sourceData.length; j++) {
    // Indexes: 3=Yellow, 4=Black, 5=Blue, 6=White
    var teamIds = [sourceData[j][3], sourceData[j][4], sourceData[j][5], sourceData[j][6]];
    
    // THE FIX: Indexes perfectly aligned to 7=Yellow, 8=Black, 9=Blue, 10=White
    var teamResults = [String(sourceData[j][7]), String(sourceData[j][8]), String(sourceData[j][9]), String(sourceData[j][10])]; 

    for (var idx = 0; idx < 4; idx++) {
       if (teamResults[idx].includes("✅") || teamResults[idx].includes("🏆")) {
           advancingPool.push(teamIds[idx]);
       }
    }
  }

  var seededTeams = advancingPool;
  var expectedTeams = numTargetMatches * 4;
  while(seededTeams.length < expectedTeams) seededTeams.push("TBD");

  var snakeMatches = [];
  if (targetRoundNum === 16) {
    snakeMatches = [
      [seededTeams[0], seededTeams[7], seededTeams[8], seededTeams[15]],
      [seededTeams[1], seededTeams[6], seededTeams[9], seededTeams[14]],
      [seededTeams[2], seededTeams[5], seededTeams[10], seededTeams[13]],
      [seededTeams[3], seededTeams[4], seededTeams[11], seededTeams[12]]
    ];
  } else if (targetRoundNum === 8) {
    snakeMatches = [
      [seededTeams[0], seededTeams[3], seededTeams[4], seededTeams[7]],
      [seededTeams[1], seededTeams[2], seededTeams[5], seededTeams[6]]
    ];
  } else if (targetRoundNum === 4) {
    snakeMatches = [
      [seededTeams[0], seededTeams[1], seededTeams[2], seededTeams[3]]
    ];
  }

  if(targetSheet.getLastRow() > 1) targetSheet.getRange(2, 1, targetSheet.getLastRow()-1, 12).clearContent();
  var h = 1, t = 1;
  for(var m=0; m<numTargetMatches; m++) {
    var group = shuffleArray(snakeMatches[m]); 
    targetSheet.appendRow([targetRoundNum, h, t, group[0], group[1], group[2], group[3], "", "", "", "", "Pending"]);
    t++; if(t > 4) { t = 1; h++; }
  }
  SpreadsheetApp.getUi().alert("Winners Advanced successfully!");
}

function clearAllData() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('CRITICAL WARNING', 'This will wipe all submitted scores. Are you sure?', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var scoresSheet = ss.getSheetByName("Scores");
    if (scoresSheet.getLastRow() > 1) scoresSheet.getRange(2, 1, scoresSheet.getLastRow() - 1, PIRATE_SCORE_HEADERS.length).clearContent(); 

    var elimTabs = ["Round of 32", "Sweet 16", "Elite 8", "Final 4"];
    elimTabs.forEach(function(tabName) {
      var sheet = ss.getSheetByName(tabName);
      if (sheet && sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).clearContent();
    });

    var matchesSheet = ss.getSheetByName("Matches");
    if (matchesSheet.getLastRow() > 1) matchesSheet.getRange(2, 8, matchesSheet.getLastRow() - 1, 1).setValue("Pending");
    
    var lbSheet = ss.getSheetByName("Live Leaderboard");
    if (lbSheet && lbSheet.getLastRow() > 1) lbSheet.getRange(2, 1, lbSheet.getLastRow() - 1, 4).clearContent();
    
    updateLiveLeaderboard(); // Reset the board with initial Sprint 3 ranks

    ui.alert("System Reset Complete.");
  }
}