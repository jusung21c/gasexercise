function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('메뉴이름')
      .addItem('First item', 'menuItem1') //.addItem('메뉴의 이름', '실행될 펑션')
      .addSeparator() //구분자
      .addSubMenu(ui.createMenu('Sub-menu') //서브메뉴 이름
          .addItem('Second item', 'menuItem2'))
      .addToUi();
}

function hello() {
  Logger.log('hello');
}

function menuItem1() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('You clicked the first menu item!');
}

function menuItem2() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('You clicked the second menu item!');
}

function copyThisSheet() {
  var nameArr = inputFormOKCancel('이름리스트','이름리스트를 컴마로 구분하여 작성해주세요').split(',');  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try{
    if(nameArr) {
      for (i in nameArr){
        ss.getActiveSheet() 
        .copyTo(ss)
        .setName(nameArr[i])
      }    
    } else {
      Browser.msgBox('입력하신 값이 잘못 되었습니다.');
    }
  } catch(e){
    Browser.msgBox(e);
  }
}

//콤마로 구분된느 이름셋과 시트 이름을 입력받아 스크립트 복사
function test() {
  var nameArr = inputFormOKCancel('이름셋','이름리스트를 컴마로 구분하여 작성해주세요').split(',');
  var originSheet = inputFormOKCancel('복사할 시트 이름','하단의 탭중 복사할 시트 이름을 입력해주세요');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try{
    if(nameArr && originSheet) {
      for (i in nameArr){
        ss.getSheetByName(originSheet) 
        .copyTo(ss)
        .setName(nameArr[i])
      }    
    } else {
      Browser.msgBox('입력하신 값이 잘못 되었습니다.');
    }
  } catch(e){
    Browser.msgBox(e);
  }
}

//해당 column (ex: B) 의 데이터 갯수를 반환한다
function getNumOfValsByCol(col){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var toAnotation =  col + "1:"+ col;  
  var vals = ss.getRange("A1:A").getValues();
  var numOfVals = vals.filter(String).length;  
  return numOfVals;
}

//해당 칼럼의 데이터들을 배열 형태로 반환한다
function getValuesByCol(col){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();    
  var rangeValues = ss.getRange(1,1,getNumOfValsByCol(col)).getValues();  
  var arr = [];
  for (i in rangeValues) {
    arr.push(rangeValues[i][0]); 
  }
  return arr;  
}

//첫번째줄 헤더에서 headername이 존재하는 index 반환
function getHeader(headername) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headRow = 1;
  var headers = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0]; 
  var row = [];
  for (i in headers) {
    if(headers[i] == headername) {
       Logger.log(i,"번째");
    }
  }
}

function sayHelloAlert() {
 var greeting = 'hello world';
  ui = SpreadsheetApp.getUi();
  ui.alert(greeting);
  Browser.msgBox('hi');
}

function printSelectionDetails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var selectedRng = ss.getActiveRange();
  Logger.log('Selected Range Details:');
  Logger.log('-- Sheet: ' + selectedRng.getSheet().getSheetName());
  Logger.log('-- Address: ' + selectedRng.getA1Notation());
  Logger.log('-- Row Count: '  + ((selectedRng.getLastRow() + 1) - selectedRng.getRow()));
  Logger.log('-- Column Count: ' + ((selectedRng.getLastColumn() + 1) - selectedRng.getColumn()));
}

// The Spreadsheet method getSheets() returns
//  an array.
// The code "ss.getSheets()[0]"
//  returns the first sheet and is equivalent to
// "ActiveWorkbook.Worksheets(1)" in VBA.
// Note that the VBA version is 1-based!
function offsetDemo() {
  var ss =
      SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getSheets()[0],
      cell = sh.getRange('B2');
  cell.setValue('Middle');
  cell.offset(-1,-1).setValue('Top Left');
  cell.offset(0, -1).setValue('Left');
  cell.offset(1, -1).setValue('Bottom Left');
  cell.offset(-1, 0).setValue('Top');
  cell.offset(1, 0).setValue('Bottom');
  cell.offset(-1, 1).setValue('Top Right');
  cell.offset(0, 1).setValue('Right');
  cell.offset(1, 1).setValue('Bottom Right');
}
function setRangeFontBold (rangeAddress) {
  var sheet = 
    SpreadsheetApp.getActiveSheet();
  sheet.getRange(rangeAddress)
    .setFontWeight('bold');
}

function inputFormOKCancel (title, body) {
  var ui = SpreadsheetApp.getUi();
  // var response = ui.prompt('제목', '내용', ui.ButtonSet.OK_CANCEL);
  var response = ui.prompt(title, body, ui.ButtonSet.OK_CANCEL);
  return response.getResponseText();
}
