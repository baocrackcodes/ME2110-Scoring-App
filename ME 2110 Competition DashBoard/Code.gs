// --- 1. CONFIGURATION ---
var TARGET_SHEET_ID = '1t94_Wo5wyd4AQCOaGK5Y-nkcHhgMl7R9uLfjOwEHgUA'; 

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ME 2110 Live Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1'); 
}

function isValidTeam(id) {
  if (!id) return false;
  var clean = String(id).trim().toUpperCase();
  return clean !== "" && clean !== "NONE" && clean !== "TBD" && clean !== "EMPTY" && clean !== "UNDEFINED";
}

function getDashboardData() {
  var payload = { schedule: [], leaderboard: [], brackets: { r32: [], s16: [], e8: [], f4: [] } };

  try {
    var ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    
    // 1. Fetch Team Names Dictionary
    var teamSheet = ss.getSheetByName("Teams");
    var teamData = teamSheet ? teamSheet.getDataRange().getValues() : [];
    var teamDict = { "TBD": "TBD", "NONE": "Empty", "EMPTY": "Empty", "UNDEFINED": "Empty" };
    for(var i=1; i<teamData.length; i++) {
      if (teamData[i][0]) teamDict[String(teamData[i][0]).trim().toUpperCase()] = teamData[i][1];
    }

    // 2. READ SINGLE SOURCE OF TRUTH LEADERBOARD
    var lbSheet = ss.getSheetByName("Live Leaderboard");
    if (lbSheet && lbSheet.getLastRow() > 1) {
      var lbData = lbSheet.getDataRange().getValues();
      for (var L = 1; L < lbData.length; L++) {
         if (lbData[L][1]) {
             payload.leaderboard.push({
                id: String(lbData[L][1]),
                name: String(lbData[L][2]),
                data: { max: lbData[L][3] }
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
            var cHeat = parseInt(rawHeat); 
            var cTrack = parseInt(rawTrack); 
            var cTeamId = String(fullData[c][mapTeamId]).trim().toUpperCase();
            
            if (!isNaN(cHeat) && !isNaN(cTrack) && isValidTeam(cTeamId)) {
               completedTracks[cRoundNum + "|" + cHeat + "|" + cTrack] = true;
            }
        }
      }
    }

    // 4. READ SCHEDULE
    var matchSheet = ss.getSheetByName("Matches");
    if (matchSheet && matchSheet.getLastRow() > 1) {
      var matchData = matchSheet.getDataRange().getValues();
      for(var m=1; m<matchData.length; m++) {
        var matchRoundStr = String(matchData[m][0]).trim().toUpperCase();
        var mNumStr = matchRoundStr.replace(/\D/g, '');
        
        if ((mNumStr === "1" || mNumStr === "2") && !matchRoundStr.includes("FLEX")) {
          var roundNum = mNumStr === "1" ? 1 : 2;
          var mRound = "Round " + roundNum;
          
          var mRawHeat = String(matchData[m][1]).toUpperCase().replace(/HEAT/g, '').trim();
          var mRawTrack = String(matchData[m][2]).toUpperCase().replace(/TRACK/g, '').trim();
          var mHeatNum = parseInt(mRawHeat);
          var mTrackNum = parseInt(mRawTrack);

          if (!isNaN(mHeatNum) && !isNaN(mTrackNum)) {
              var yId = String(matchData[m][3] || "").trim().toUpperCase();
              var bkId = String(matchData[m][4] || "").trim().toUpperCase();
              var blId = String(matchData[m][5] || "").trim().toUpperCase();
              var wId = String(matchData[m][6] || "").trim().toUpperCase(); 

              var isCompleteByTrack = completedTracks[roundNum + "|" + mHeatNum + "|" + mTrackNum] === true;

              payload.schedule.push({
                round: mRound,
                heat: mHeatNum, 
                track: mTrackNum, 
                status: isCompleteByTrack ? "Complete" : "Pending",
                teams: {
                  yellow: String(teamDict[yId] || matchData[m][3] || "Empty"),
                  black:  String(teamDict[bkId] || matchData[m][4] || "Empty"),
                  blue:   String(teamDict[blId] || matchData[m][5] || "Empty"),
                  white:  String(teamDict[wId] || matchData[m][6] || "Empty")
                }
              });
          }
        }
      }
    }

    // 5. READ PLAYOFF BRACKETS (WITH FAKE SCORE OVERRIDE TRICK)
    var bracketTabs = [
      { name: "Round of 32", key: "r32" }, { name: "Sweet 16", key: "s16" }, 
      { name: "Elite 8", key: "e8" }, { name: "Final 4", key: "f4" }
    ];

    bracketTabs.forEach(function(tab) {
      var bSheet = ss.getSheetByName(tab.name);

      if (bSheet && bSheet.getLastRow() > 1) {
        var bData = bSheet.getDataRange().getValues();
        for(var b=1; b<bData.length; b++) {
          if (!bData[b][0]) continue;
          
          var yId = String(bData[b][3] || "").trim().toUpperCase();
          var bkId = String(bData[b][4] || "").trim().toUpperCase();
          var blId = String(bData[b][5] || "").trim().toUpperCase();
          var wId = String(bData[b][6] || "").trim().toUpperCase();

          var yStr = String(bData[b][7] || "");
          var bkStr = String(bData[b][8] || "");
          var blStr = String(bData[b][9] || "");
          var wStr = String(bData[b][10] || "");

          // JEDI MIND TRICK: If spreadsheet has a checkmark, send a fake '1000' score 
          // to the untouched frontend so it automatically sorts them to the top and draws the UI.
          var yFakeScore = (yStr.includes("✅") || yStr.includes("🏆")) ? 1000 : 0;
          var bkFakeScore = (bkStr.includes("✅") || bkStr.includes("🏆")) ? 1000 : 0;
          var blFakeScore = (blStr.includes("✅") || blStr.includes("🏆")) ? 1000 : 0;
          var wFakeScore = (wStr.includes("✅") || wStr.includes("🏆")) ? 1000 : 0;

          var blankWeights = {tier1:0, tier2:0, tier3:0, tier4:0};

          payload.brackets[tab.key].push({
            heat: String(bData[b][1]), track: String(bData[b][2]), status: String(bData[b][11]),
            teams: [
              { name: String(teamDict[yId] || bData[b][3]), score: yFakeScore, color: "#FFD700", txt: "#000", w: blankWeights }, 
              { name: String(teamDict[bkId] || bData[b][4]), score: bkFakeScore, color: "#222222", txt: "#FFF", w: blankWeights }, 
              { name: String(teamDict[wId] || bData[b][6]), score: wFakeScore, color: "#FFFFFF", txt: "#000", w: blankWeights }, 
              { name: String(teamDict[blId] || bData[b][5]), score: blFakeScore, color: "#005A9C", txt: "#FFF", w: blankWeights }  
            ]
          });
        }
      }
    });

  } catch (error) {
    payload.error = String(error.message); 
  }

  return JSON.stringify(payload);
}