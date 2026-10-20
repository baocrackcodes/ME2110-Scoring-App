// --- 1. CONFIGURATION ---
var TARGET_SHEET_ID = '1t94_Wo5wyd4AQCOaGK5Y-nkcHhgMl7R9uLfjOwEHgUA'; 

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ME 2110 Competitor Hub')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1'); 
}

function isValidTeam(id) {
  if (!id) return false;
  var clean = String(id).trim().toUpperCase();
  return clean !== "" && clean !== "NONE" && clean !== "TBD" && clean !== "EMPTY" && clean !== "UNDEFINED";
}

// --- 2. FETCH COMPETITOR DATA ---
function getCompetitorData() {
  var payload = { sections: {}, leaderboard: [], matches: [], brackets: [] };

  try {
    var ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    
    // 1. Build Team Hierarchy for Dropdowns
    var teamSheet = ss.getSheetByName("Teams");
    var teamData = teamSheet ? teamSheet.getDataRange().getValues() : [];
    
    for(var i=1; i<teamData.length; i++) {
      if (teamData[i][0]) {
        var cleanId = String(teamData[i][0]).trim().toUpperCase();
        var tName = String(teamData[i][1]).trim();
        
        if (isValidTeam(cleanId)) {
            var section = cleanId.split('-')[0];
            if (!payload.sections[section]) payload.sections[section] = [];
            payload.sections[section].push({ id: cleanId, name: tName });
        }
      }
    }
    
    // Sort Sections and Teams alphabetically
    for (var sec in payload.sections) {
        payload.sections[sec].sort(function(a,b) { return a.id.localeCompare(b.id); });
    }

    // 2. READ SINGLE SOURCE OF TRUTH LEADERBOARD
    var lbSheet = ss.getSheetByName("Live Leaderboard");
    if (lbSheet && lbSheet.getLastRow() > 1) {
      var lbData = lbSheet.getDataRange().getValues();
      for (var L = 1; L < lbData.length; L++) {
         if (lbData[L][1]) {
             payload.leaderboard.push({
                rank: parseInt(lbData[L][0]),
                id: String(lbData[L][1]),
                name: String(lbData[L][2]),
                max: parseInt(lbData[L][3]) // -1 inherently handles "No Score" logic in the frontend
             });
         }
      }
    }

    // 3. READ SCORES TAB STRICTLY FOR QUALIFYING COMPLETION STATUS
    var scoresSheet = ss.getSheetByName("Scores");
    var completedTracks = {};
    
    if (scoresSheet && scoresSheet.getLastRow() > 1) {
      var fullData = scoresSheet.getDataRange().getValues();
      var headers = fullData[0]; 
      var findCol = function(names) {
         for(var j=0; j<names.length; j++) {
             for (var i=0; i<headers.length; i++) { if (String(headers[i]).toLowerCase().trim() === names[j].toLowerCase().trim()) return i; }
             for (var i=0; i<headers.length; i++) { if (String(headers[i]).toLowerCase().trim().indexOf(names[j].toLowerCase().trim()) > -1) return i; }
         }
         return -1;
      };
      
      var mapRound = findCol(["Round"]); 
      var mapTrack = findCol(["Track Number", "Track"]); 
      var mapHeat = findCol(["Heat"]);
      var mapTeamId = findCol(["Team ID", "ID"]); 

      for(var c=1; c<fullData.length; c++) {
        var rawRoundStr = String(fullData[c][mapRound]).trim().toUpperCase();
        var numStr = rawRoundStr.replace(/\D/g, ''); 
        var cRoundNum = -1;
        if (numStr === "1" && !rawRoundStr.includes("FLEX")) cRoundNum = 1;
        else if (numStr === "2" && !rawRoundStr.includes("FLEX")) cRoundNum = 2;

        if (cRoundNum === 1 || cRoundNum === 2) {
            var rawHeat = String(fullData[c][mapHeat]).toUpperCase().replace(/HEAT/g, '').trim();
            var rawTrack = String(fullData[c][mapTrack]).toUpperCase().replace(/TRACK/g, '').replace(/\(COMPLETED\)/g, '').trim();
            var cHeat = parseInt(rawHeat); var cTrack = parseInt(rawTrack); 
            var cTeamId = String(fullData[c][mapTeamId]).trim().toUpperCase();
            if (!isNaN(cHeat) && !isNaN(cTrack) && isValidTeam(cTeamId)) completedTracks[cRoundNum + "|" + cHeat + "|" + cTrack] = true;
        }
      }
    }

    // 4. Fetch Qualifying Matches
    var matchSheet = ss.getSheetByName("Matches");
    if (matchSheet && matchSheet.getLastRow() > 1) {
      var matchData = matchSheet.getDataRange().getValues();
      for(var m=1; m<matchData.length; m++) {
        var matchRoundStr = String(matchData[m][0]).trim().toUpperCase();
        var mNumStr = matchRoundStr.replace(/\D/g, '');
        if ((mNumStr === "1" || mNumStr === "2") && !matchRoundStr.includes("FLEX")) {
          var roundNum = mNumStr === "1" ? 1 : 2;
          var mRawHeat = String(matchData[m][1]).toUpperCase().replace(/HEAT/g, '').trim();
          var mRawTrack = String(matchData[m][2]).toUpperCase().replace(/TRACK/g, '').trim();
          var mHeatNum = parseInt(mRawHeat); var mTrackNum = parseInt(mRawTrack);

          if (!isNaN(mHeatNum) && !isNaN(mTrackNum)) {
              var isCompleteByTrack = completedTracks[roundNum + "|" + mHeatNum + "|" + mTrackNum] === true;
              payload.matches.push({
                round: "Round " + roundNum, heat: mHeatNum, track: mTrackNum, status: isCompleteByTrack ? "Complete" : "Pending",
                y: String(matchData[m][3] || "").trim().toUpperCase(), bk: String(matchData[m][4] || "").trim().toUpperCase(),
                w: String(matchData[m][6] || "").trim().toUpperCase(), bl: String(matchData[m][5] || "").trim().toUpperCase()
              });
          }
        }
      }
    }

    // 5. Fetch Playoff Brackets
    var bracketTabs = ["Round of 32", "Sweet 16", "Elite 8", "Final 4"];
    bracketTabs.forEach(function(tabName) {
      var bSheet = ss.getSheetByName(tabName);
      if (bSheet && bSheet.getLastRow() > 1) {
        var bData = bSheet.getDataRange().getValues();
        for(var b=1; b<bData.length; b++) {
          if (!bData[b][0]) continue;
          payload.brackets.push({
            round: tabName, heat: String(bData[b][1]), track: String(bData[b][2]), status: String(bData[b][11]),
            y: String(bData[b][3] || "").trim().toUpperCase(), bk: String(bData[b][4] || "").trim().toUpperCase(),
            w: String(bData[b][6] || "").trim().toUpperCase(), bl: String(bData[b][5] || "").trim().toUpperCase()
          });
        }
      }
    });

  } catch (error) { payload.error = String(error.message); }
  return JSON.stringify(payload);
}