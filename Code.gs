function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('记账')
      .addItem('初始化账户', 'createAccountSheet')
      .addItem('初始化流水', 'createCashSheet')
      .addItem('新建预算', 'newBudget')
      .addToUi();
}

var accountColumns = ['F'];
var inColumns = ['J','K'];
var outColumns = ['L','M','N'];

function createAccountSheet(){
  
  if(findSheet('账户')){
    return;
  }
  
  var accountSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('账户');
  
  // 
  var accountCategory = ['分类','名称','信用','现金','余额','活期','定期','福利','投资'];
  for(var row=1;row<=accountCategory.length;row++){
    accountSheet.getRange(row, 1).setValue(accountCategory[row-1]);
  }
  accountSheet.getRange(2, 2).setValue('金额');
  accountSheet.getRange(2, 3).setValue('分布');
  for(var row=3;row<=accountCategory.length;row++){
    accountSheet.getRange(row, 2).setValue('=sumif(E:E,A'+row+',G:G)');
    var f = '=iferror(if(B'+row+'<0,B'+row+'/sumif(B:B,"<0"),B'+row+'/sumif(B:B,">=0")),0)';
    accountSheet.getRange(row, 3).setValue(f);
  }
  accountSheet.getRange('B:B').setNumberFormat('¥0.00');
  accountSheet.getRange('C:C').setNumberFormat('0.00%');
  
  //
  accountSheet.getRange(1, 5).setValue('账户');
  setRowValues(accountSheet, 2, 5, ['类型','名称','余额','分布']); 
  var cell = accountSheet.getRange('A3:A'+row);
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(cell).build();
  accountSheet.getRange('E3').setDataValidation(rule);
  if(findSheet('流水')){
    var f="=sumif('流水'!G:G,F3,'流水'!D:D)+SUMIFS('流水'!D:D,'流水'!F:F,F3,'流水'!B:B,\"收入\")+SUMIFS('流水'!D:D,'流水'!F:F,F3,'流水'!B:B,\"支出\")";
    Logger.log(f);
    accountSheet.getRange('G3').setValue(f);
  }
  accountSheet.getRange('G:G').setNumberFormat('¥0.00');
  accountSheet.getRange('H3').setValue('=iferror(G3/sumif(E:E,E3,G:G),0)')
  accountSheet.getRange('H:H').setNumberFormat('0.00%');
  
  //
  setColumnValues(accountSheet, 1, 10, ['收入','固定收入','工资']);
  setColumnValues(accountSheet, 2, 11, ['其他收入','奖金','投资','其他']);
  setColumnValues(accountSheet, 1, 12, ['支出','固定支出','房贷']);
  setColumnValues(accountSheet, 2, 13, ['常规支出','饮食','水电日用','通信','交通']);
  setColumnValues(accountSheet, 2, 14, ['其他支出','衣帽鞋包','社交娱乐','网络数码','阅读','健身','其他']);
  
  for (var i = 1; i < 15; i++) {
    accountSheet.setColumnWidth(i, 80);
  }
}

function createCashSheet(){
  
  if(findSheet('流水')){
    return;
  }
  
  var cashSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('流水');
  
  cashSheet.appendRow(['日期','出入','类型','金额','说明','账户']);
  
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['支出','转账','收入']).build();
  cashSheet.getRange('B:B').setDataValidation(rule);
  cashSheet.getRange('B1').clearDataValidations();
}

function onEdit(){
  var sheet = SpreadsheetApp.getActiveSheet();
  if(sheet.getName() == '流水'){
    newCash(sheet, sheet.getActiveCell());
  }
}

function newCash(cashSheet, cell){
  Logger.log(cell.getRow()+","+cell.getColumn())
  
  var row = cell.getRow();
  var accountSheet = findSheet('账户');
  
  if(cell.getColumn() == 2){
    // 如果金额未填，自动添加日期
    if(cashSheet.getRange(row, 4).getValue() == ""){
      var time = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd");
      cashSheet.getRange(row, 1).setValue(time);
    }
    
    cashSheet.getRange('C'+row+':C'+row).clear();
    cashSheet.getRange('C'+row+':C'+row).clearDataValidations();
    // 根据 cell.getValue() 为类型添加不同DataValidation
    var rangeName;
    if(cell.getValue() == '支出'){
      rangeName = getLastRange(accountSheet, outColumns, 3);
    }else if(cell.getValue() == '收入'){
      rangeName = getLastRange(accountSheet, inColumns, 3);
    }
    if(rangeName != undefined){      
      var range = accountSheet.getRange(rangeName);
      var rule = SpreadsheetApp.newDataValidation().requireValueInRange(range);
      cashSheet.getRange('C'+row+':C'+row).setDataValidation(rule);
    }
  }
  
  var range = accountSheet.getRange(getLastRange(accountSheet, accountColumns, 3));
  rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
  cashSheet.getRange('F'+row+':G'+row).setDataValidation(rule);
}

function newBudget(){
  var budgetSheet = findSheet('预算');
  
  if(budgetSheet == undefined){
    var budgetSheet = ss.insertSheet('预算');
  }
  
  var lastRow = budgetSheet.getLastRow();
  
  appendBudget(budgetSheet, lastRow == 0 ? 1:lastRow+2);
}

function appendBudget(sheet, row){
  var nextMonth = new Date().getMonth()+1;
  if(nextMonth == 12) nextMonth = 0;
  nextMonth+=1;
  
  var cell = sheet.getRange(row, 1);
  cell.setValue(nextMonth+'月');
}

function findSheet(name){
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for(var i=0; i < sheets.length; i++){
    if(sheets[i].getName() == name){
      return sheets[i];
    }
  }
}

function setRowValues(sheet, row, column, values){
  for(var i=0; i<values.length; i++){
    sheet.getRange(row, column+i).setValue(values[i]);
  }
}

function setColumnValues(sheet, row, column, values){
  for(var i=0; i<values.length; i++){
    sheet.getRange(row+i, column).setValue(values[i]);
  }
}

function getLastRange(sheet, colNames, startRow){
  var endRow = startRow;
  for(var i = 0; i< colNames.length; i++){
    var temp = sheet.getRange(colNames[i]+':'+colNames[i]).getValues().length;
    Logger.log(colNames[i]+":"+temp);
    if(temp > endRow){
      endRow = temp;
    }
  }
  
  Logger.log(colNames[0]+startRow+':'+colNames[colNames.length-1]+endRow);
  return colNames[0]+startRow+':'+colNames[colNames.length-1]+endRow;
}