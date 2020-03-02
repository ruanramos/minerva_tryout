var NUMBER_OF_GAMES = 7;
var ROLE_COLUMN_NUMBER = 3;
var NUMBER_OF_DECIMAL_CASES = 2;
var NUM_INFO_COLUMNS = 3;
var TRYOUT_SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
var RESULT_SHEET = TRYOUT_SPREADSHEET.getSheetByName("Results");
var ALL_SHEETS_NAMES = getAllSheetsNames(); // names as string array
var PLAYER_HEADER_CELL = "A1";
var PLAYER_ROLE_CELL = "A2";
var ROLES_STEMS = ['TOP', 'MID', 'AD', 'JUN', 'SUP'];

function onOpen()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuItems=[{name: 'Go To Tab', functionName: 'goToTab'} ];
  ss.addMenu('Find Tab By Name', menuItems);

  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Generate Results')
      .addItem('Get Results', 'generateResults')
      .addToUi();
}

function getRawPlayerRole(playerSheet) {
  return playerSheet.getRange(PLAYER_ROLE_CELL).getValue()
  .split('\n')[2]
  .split(':')[1]
  .trim();
}

function treatPlayerRole(rawRole) {
  if (rawRole.indexOf('xxx') > -1) {
    return '';
  }
  var returnValue = null;
  ROLES_STEMS.forEach(function(x) {
    if (rawRole.toUpperCase().indexOf(x) > -1) {
      returnValue = x;
    }
  });
  return returnValue === 'JUN' ? returnValue + 'GLE' : returnValue;
}

function createHeader() {
  var headerElements = ['Name', 'Nickname', 'Role'];
  for (var i = 1; i <= NUMBER_OF_GAMES; i++) {
    headerElements.push("Game " + i);
  }
  headerElements.push('Final Note');
  RESULT_SHEET.appendRow(headerElements);
  var headerRange = RESULT_SHEET.getRange(1, 1, 1, RESULT_SHEET.getLastColumn());
  headerRange.setFontSize(14);
  headerRange.setHorizontalAlignment("center");
  headerRange.setFontWeight("bold");
  //RESULT_SHEET.hideRows(1);
}

function addPlayers() {
  ALL_SHEETS_NAMES.forEach(function(x) {
    if (x === "Results" || x.indexOf("EXEMPLO") > -1) {
    } else {
      try {
        var playerSheet = TRYOUT_SPREADSHEET.getSheetByName(x);
        var rawPlayerRole = getRawPlayerRole(playerSheet);
        var playerRole = treatPlayerRole(rawPlayerRole);
        if (playerRole != '') {
          //setAverage();
          var playerNotes = [];
          for (var i = 8; i < NUMBER_OF_GAMES * 10; i += 9) {
            var note = playerSheet.getRange(i, 2).getValue();
            if (typeof(note) === 'number') {
              playerNotes.push(Math.round(note*100)/100);
            } else {
              playerNotes.push('');
            }
          }
          var nameAndNickname = x.split("-");
          var name = nameAndNickname[0].trim();
          var nickName = nameAndNickname[1].trim();
          var newPlayerRow = RESULT_SHEET.getRange(RESULT_SHEET.getLastRow() + 1, 1, 1, RESULT_SHEET.getLastColumn());
          if (newPlayerRow.getFontSize() !== 10 || newPlayerRow.getFontWeight() !== 'normal') {
            newPlayerRow.setFontSize(10);
            newPlayerRow.setFontWeight('normal');
          }
          newPlayerRow.setWrap(true);
          newPlayerRow.setHorizontalAlignment("center");
          newPlayerRow.setVerticalAlignment("middle");
          var finalPlayerRow = [name, nickName, playerRole].concat(playerNotes);
          finalPlayerRow.push(getAverageNote(playerNotes));
          Logger.log("Player: %s", finalPlayerRow);
          RESULT_SHEET.appendRow(finalPlayerRow);
          RESULT_SHEET.hideRows(1);
          RESULT_SHEET.sort(NUMBER_OF_GAMES + NUM_INFO_COLUMNS + 1, false).sort(ROLE_COLUMN_NUMBER);
          RESULT_SHEET.unhideRow(RESULT_SHEET.getRange('A1'));
        }
        
      } catch (e) {
        Logger.log(e);
      }
    }
  });
  RESULT_SHEET.autoResizeColumns(1, RESULT_SHEET.getLastColumn());
  RESULT_SHEET.setColumnWidth(3, 100);
}

function getAverageNote(notes) {
  var sum = 0;
  var numNotes = 0;
  for (var i = 0; i < notes.length; i++) {
    if (typeof(notes[i]) == "number") {
      sum += notes[i];
      numNotes++;
    }
  }
  return (parseFloat((sum / numNotes).toFixed(NUMBER_OF_DECIMAL_CASES)));
}

function setAverage() {
  // this is such a gambiarra, I'm ashamed
  var cellPrefix = 'G';
  for (var i = 2; i <= RESULT_SHEET.getLastRow(); i++) {
    var cell = cellPrefix + i;
    var formula = '=AVERAGE(D' + i + ':F' + i + ')';
    RESULT_SHEET.getRange(cell).setFormula(formula);
  }
  
}

function generateResults() {
  RESULT_SHEET.activate();
  RESULT_SHEET.clearContents();
  RESULT_SHEET.clearFormats();
  createHeader();
  addPlayers();
  styleResultSheet();
}

function styleResultSheet() {
  RESULT_SHEET.activate();
  var lastColumnNumber = RESULT_SHEET.getLastColumn();
  var lastRowNumber = RESULT_SHEET.getLastRow();
  var colors = {
    "AD": "#FF2E64",
    "TOP": "#3AE5FA",
    "JUNGLE": "#AEF14B",
    "MID": "#E970FF",
    "SUP": "#37F8D8",
  };
  // header coloring
  RESULT_SHEET.getRange(1, 1, 1, lastColumnNumber).setBackground("#B0B0B0");
  
  // sheet coloring
  for (var i = 2; i < lastRowNumber + 1; i++) {
    var role = RESULT_SHEET.getRange(i, ROLE_COLUMN_NUMBER).getValue();
    var color = null;
    if (role === "TOP") {
      color = colors["TOP"];
    } else if (role === "AD") {
      color = colors["AD"];
    } else if (role === "SUP") {
      color = colors["SUP"];
    } else if (role === "JUNGLE") {
      color = colors["JUNGLE"];
    } else if (role === "MID") {
      color = colors["MID"];
    }
    
    RESULT_SHEET.getRange(i, 1, 1, lastColumnNumber).setBackground(color);
  }
}

function getAllSheetsNames() {
  var names = TRYOUT_SPREADSHEET.getSheets().map(function(x) {
    return x.getName();
  });
  return names;
}

function goToResultsTab() {
  TRYOUT_SPREADSHEET.getSheetByName("Results").activate();
}

function goToTabByName(name, inputName) {
  try {
    TRYOUT_SPREADSHEET.getSheetByName(name).activate();
  } catch(e) {
    Browser.msgBox('Sheet named: "' + inputName + '" does not exists!');
  }
}

function goToTab() {
  var inputName = Browser.inputBox('Enter Tab Name:','',Browser.Buttons.OK_CANCEL);
  inputName = inputName.toUpperCase();
  var names = getAllSheetsNames();
  
  names.forEach(function(x) {
    x1 = x.toUpperCase();
    if (x1.indexOf(inputName) > -1) {
      goToTabByName(x);
    }
  });
};

function getCurrentCellValue()
{
  var resultsTabName = "Results";
  TRYOUT_SPREADSHEET.getSheetByName("Filipy - ChaikaOne").activate();
  
  var cell = TRYOUT_SPREADSHEET.getRange(17, 2).activateAsCurrentCell();
  var a1 = cell.getA1Notation();
  Logger.log(a1);
  Logger.log(cell);
  Logger.log(typeof(cell.getValue()));
  
  var playerSheets = TRYOUT_SPREADSHEET.getSheets();
  playerSheets.forEach(function() {
    
  });
  //SpreadsheetApp.getUi().alert("The active cell "+a1+" value is "+val);
}

function bla() {
  RESULT_SHEET.sort(7);
}
