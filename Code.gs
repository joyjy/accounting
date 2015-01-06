function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('记账')
    .addItem('初始化流水表', 'createCashSheet')
    .addItem('初始化账户表', 'createAccountSheet')
    .addItem('新预算', 'newBudget')
    //.addSeparator()
    //.addSubMenu(ui.createMenu('招商')
    //               .addItem('信用卡债务','balance')
    //               .addItem('还款方案对比', 'repayments'))
    .addToUi();
}

// 当任意单元格更新时回调
function onEdit(){
  var sheet = SpreadsheetApp.getActiveSheet();
  if(sheet.getName() == '流水'){
    newCash(sheet, sheet.getActiveCell());
  } else if(sheet.getName() == '账户'){
    newAccount(sheet, sheet.getActiveCell());
  } else if(sheet.getName() == '预算'){
    updateBudget(sheet, sheet.getActiveCell());
  }
}

// 复制
function copy(fromSheet, rangeName, targetSheet, startRow, startColumn){
  var range = fromSheet.getRange(rangeName);
  range.copyTo(targetSheet.getRange(startRow, startColumn));
}

// 获取表
function findSheet(name){
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for(var i=0; i < sheets.length; i++){
    if(sheets[i].getName() == name){
      return sheets[i];
    }
  }
}

// 为行填充值
function setRowValues(sheet, row, column, values){
  for(var i=0; i<values.length; i++){
    sheet.getRange(row, column+i).setValue(values[i]);
  }
}

// 为列填充值
function setColumnValues(sheet, row, column, values){
  for(var i=0; i<values.length; i++){
    sheet.getRange(row+i, column).setValue(values[i]);
  }
}

// 获取指定行列有值的范围
function getLastRange(sheet, colNames, startRow, endRow){
  if(endRow == undefined){
    endRow = getLastRow(sheet, colNames, startRow);
  }
  return colNames[0]+startRow+':'+colNames[colNames.length-1]+endRow;
}

// 获取指定列最后有值的行
function getLastRow(sheet, colNames, startRow){
  var endRow = startRow;
  for(var i = 0; i< colNames.length; i++){
    var values = sheet.getRange(colNames[i]+':'+colNames[i]).getValues();
    for(var j=startRow; j< values.length; j++)
    {
      if(values[j] == '') break;
    }
    var temp = j;
    if(temp > endRow){
      endRow = temp;
    }
  }
  return endRow;
}

// 填充值
function fill(value, size, row){
  var array = new Array();
  for(var i=0; i<size; i++){
    if(row){
      array.push(value.replace(/{}/g, row));
      row++;
    }else{
      array.push(value);
    }
  }
  return array;
}

function formatDate(date){
  return Utilities.formatDate(date, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd");
}