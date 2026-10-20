// ME 2110 Sprint 3 Pirate Scoring App — section-based scoring
var SPRINT_PIRATE_HEADERS = [
  'Round','Color Zone','Team ID','Team Name','Activated','Left Dock',
  'Gold On Mound','Gold Surrounding','Jewels On Mound','Jewels Surrounding',
  'Flag Rank','Black Cannonballs','White Parcels','Mermaids Toppled','Mermaids Moved',
  'Davy Jones Outcome','Total Score','DQ?','DQ Reason','Timestamp'
];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ Sprint 3 Admin')
    .addItem('0. Install Pirate Sprint 3 Sheets (one time)','setupSprintPirateSheets')
    .addSeparator()
    .addItem('⚠️ WIPE ALL SPRINT 3 SCORES','clearSprintData')
    .addToUi();
}
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ME 2110 Sprint 3 Pirate Scoring')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
function isValidSprintTeam(id) {
  return !!id && !['','NONE','TBD','EMPTY','UNDEFINED'].includes(String(id).trim().toUpperCase());
}
function getSprintTeams(section) {
  var sec=String(section||'').trim().toUpperCase();
  if(!/^A(0[1-9]|1[0-2])$/.test(sec)) throw new Error('Invalid section: '+sec);
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teams (DO NOT TOUCH)');
  if(!sheet) throw new Error('Missing Teams (DO NOT TOUCH) tab.');
  var data=sheet.getDataRange().getValues(), teams=[];
  for(var i=1;i<data.length;i++) {
    var id=String(data[i][0]||'').trim().toUpperCase();
    if(id.split('-')[0]===sec && isValidSprintTeam(id))
      teams.push({id:id,name:String(data[i][1]||'').trim()});
  }
  teams.sort(function(a,b){return a.id.localeCompare(b.id);});
  return teams;
}
function getSprintMatrix(n) {
  if (n !== 7) throw new Error('Expected exactly 7 teams in each section; found ' + n + '. Check Teams (DO NOT TOUCH).');
  return [
    [0,1,2,3], [4,5,6,0], [1,2,4,5], [3,6,0,1],
    [2,4,5,6], [0,3,1,2], [4,5,6,3]
  ];
}
function getAvailableRounds(section) {
  var matrix=getSprintMatrix(getSprintTeams(section).length);
  return matrix.map(function(_,i){return 'Round '+(i+1);});
}
function getMatchData(section,round) {
  var teams=getSprintTeams(section);
  var matrix=getSprintMatrix(teams.length);
  var m=/^Round ([1-7])$/.exec(String(round||''));
  if(!m || !matrix[Number(m[1])-1])return null;
  var indices=matrix[Number(m[1])-1];
  return {yellow:teams[indices[0]],black:teams[indices[1]],
    blue:teams[indices[2]],white:teams[indices[3]]};
}

// Archives all existing section scoring tabs individually, then creates new pirate schema tabs.
// No historical score is automatically converted or destroyed.
function setupSprintPirateSheets() {
  var ss=SpreadsheetApp.getActiveSpreadsheet(),ui=SpreadsheetApp.getUi();
  var answer=ui.alert('Install Sprint 3 pirate scoring schema',
    'Back up this spreadsheet first. Existing A01–A12 Scores tabs with old headers will be renamed as archives and new empty score tabs created. Continue?',
    ui.ButtonSet.YES_NO);
  if(answer!==ui.Button.YES)return;
  var updated=[];
  for(var i=1;i<=12;i++) {
    var sec='A'+('0'+i).slice(-2),name=sec+' Scores';
    var sheet=ss.getSheetByName(name);
    if(sheet) {
      var h=sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),1)).getValues()[0];
      if(String(h[6]).trim()==='Gold On Mound' && String(h[16]).trim()==='Total Score')continue;
      var oldName=sec+' Scores Archive '+Utilities.formatDate(new Date(),ss.getSpreadsheetTimeZone(),'yyyyMMdd-HHmmss');
      sheet.setName(oldName);
    }
    var fresh=ss.insertSheet(name);
    fresh.getRange(1,1,1,SPRINT_PIRATE_HEADERS.length).setValues([SPRINT_PIRATE_HEADERS]);
    fresh.setFrozenRows(1);
    updated.push(name);
  }
  ui.alert('Pirate scoring installed in '+updated.length+' section tabs. Historical data was archived.');
}

function saveScoresToSheet(payload) {
  if(!payload||!Array.isArray(payload.teams)||payload.teams.length!==4)
    throw new Error('A valid submission requires four zone entries.');
  var sec=String(payload.section||'').trim().toUpperCase();
  var roster=getMatchData(sec,payload.round);
  if(!roster)throw new Error('Invalid section or round.');
  var name=sec+' Scores',ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName(name);
  if(!sheet)throw new Error('Missing '+name+'. Run setupSprintPirateSheets first.');
  var headers=sheet.getRange(1,1,1,SPRINT_PIRATE_HEADERS.length).getValues()[0];
  if(headers.some(function(h,i){return h!==SPRINT_PIRATE_HEADERS[i];}))
    throw new Error('Old scoring columns detected in '+name+'. Run setupSprintPirateSheets first.');
  var count={gold:0,jewel:0,black:0,white:0},flags={};
  var flagPts={None:0,First:40,Second:28,Third:17,Fourth:7};
  var davyPts={None:0,ChestShip:11,ChestDock:32,HeartShip:39,HeartDock:61};
  var seen={},rows=[],stamp=new Date();
  function amount(value,max,label) {
    var n=Number(value);
    if(!Number.isInteger(n)||n<0||n>max)throw new Error('Invalid '+label+': '+value);
    return n;
  }
  var byColor={YELLOW:'yellow',BLACK:'black',BLUE:'blue',WHITE:'white'};
  payload.teams.forEach(function(t){
    var col=String(t.color||'').toUpperCase(),key=byColor[col];
    if(!key||seen[col])throw new Error('Missing or duplicate zone '+col);
    seen[col]=true;
    var id=String(t.teamId||'').trim().toUpperCase();
    if(id!==String(roster[key].id).toUpperCase())throw new Error('Team assignment mismatch for '+col);
    var gTop=amount(t.goldTop,10,'gold on mound'),gZone=amount(t.goldZone,10,'gold surrounding');
    var jTop=amount(t.jewelTop,5,'jewels on mound'),jZone=amount(t.jewelZone,5,'jewels surrounding');
    var black=amount(t.blackCannonballs,16,'black cannonballs'),white=amount(t.whiteParcels,4,'white parcels');
    var mTop=amount(t.mermaidsToppled,2,'mermaids toppled'),mMoved=amount(t.mermaidsMoved,2,'mermaids moved');
    if(mTop+mMoved>2)throw new Error('Max 2 mermaids per team.');
    if(!Object.prototype.hasOwnProperty.call(flagPts,t.flagRank))throw new Error('Invalid flag rank.');
    if(!Object.prototype.hasOwnProperty.call(davyPts,t.davyOutcome))throw new Error('Invalid Davy Jones outcome.');
    if(t.flagRank!=='None') {
      if(flags[t.flagRank])throw new Error('Duplicate flag rank: '+t.flagRank);
      flags[t.flagRank]=true;
    }
    count.gold+=gTop+gZone;count.jewel+=jTop+jZone;count.black+=black;count.white+=white;
    var activated=t.activate===true||String(t.activate).toUpperCase()==='TRUE';
    var left=t.leaveDock===true||String(t.leaveDock).toUpperCase()==='TRUE';
    if(left&&!activated)throw new Error('Cannot leave dock without activating.');
    var dq=t.isDQ===true||String(t.isDQ).toUpperCase()==='TRUE';
    if(!activated&&!dq)throw new Error('Team must be activated or disqualified: '+id);
    var score=dq?0:(activated?1:0)+(left?10:0)+gTop*10+gZone*4+jTop*15+jZone*9+
      flagPts[t.flagRank]+black*9-white*10+mTop*12+mMoved*5+davyPts[t.davyOutcome];
    rows.push([String(payload.round),col,id,String(roster[key].name),activated,left,
      gTop,gZone,jTop,jZone,t.flagRank,black,white,mTop,mMoved,t.davyOutcome,score,dq,
      String(t.dqReason||''),stamp]);
  });
  if(count.gold>10||count.jewel>5||count.black>16||count.white>4)
    throw new Error('Shared resource maximum exceeded (gold 10, jewels 5, black 16, white 4).');
  var lock=LockService.getDocumentLock()||LockService.getScriptLock();
  lock.waitLock(30000);
  try{sheet.getRange(sheet.getLastRow()+1,1,rows.length,SPRINT_PIRATE_HEADERS.length).setValues(rows);}
  finally{lock.releaseLock();}
  return true;
}

function clearSprintData() {
  var ui=SpreadsheetApp.getUi();
  var answer=ui.alert('WARNING','Clear current pirate Sprint 3 scores in A01–A12 Scores tabs? Archived tabs will not be touched.',ui.ButtonSet.YES_NO);
  if(answer!==ui.Button.YES)return;
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  for(var i=1;i<=12;i++) {
    var sheet=ss.getSheetByName('A'+('0'+i).slice(-2)+' Scores');
    if(sheet&&sheet.getLastRow()>1)
      sheet.getRange(2,1,sheet.getLastRow()-1,SPRINT_PIRATE_HEADERS.length).clearContent();
  }
  ui.alert('Current Sprint 3 scores cleared.');
}
