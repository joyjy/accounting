// 创建流水表
function createCashSheet(){
  if(findSheet('流水')){
    return;
  }
  
  var cashSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('流水');
  
  cashSheet.appendRow(['日期','出入','类型','金额','说明','账户']);
  
  cashSheet.getRange('D:D').setNumberFormat('¥0.00');
  
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['支出','转账','收入']).build();
  cashSheet.getRange('B:B').setDataValidation(rule);
  cashSheet.getRange('B1').clearDataValidations();
  
  cashSheet.setColumnWidth(7, 20);
}

// 新建流水项
function newCash(cashSheet, cell){
  Logger.log('newCash: '+cell.getRow()+","+cell.getColumn())
  
  var row = cell.getRow();
  var accountSheet = findSheet('账户');
  
  if(cell.getColumn() == 2){
    // 如果金额已填，不更新行
    if(cashSheet.getRange(row, 4).getValue() != ''){
      return;
    }
    
    // 自动添加日期
    var time = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd");
    cashSheet.getRange(row, 1).setValue(time);
    
    // 根据 cell.getValue() 为类型（C列）添加不同DataValidation
    cashSheet.getRange('C'+row+':C'+row).clear();
    cashSheet.getRange('C'+row+':C'+row).clearDataValidations();
    var rangeColumnName;
    if(cell.getValue() == '支出'){
      rangeColumnName = getLastRange(accountSheet, outColumns, 3);
    }else if(cell.getValue() == '收入'){
      rangeColumnName = getLastRange(accountSheet, inColumns, 3);
    }
    if(rangeColumnName != undefined){      
      var range = accountSheet.getRange(rangeColumnName);
      var rule = SpreadsheetApp.newDataValidation().requireValueInRange(range);
      cashSheet.getRange('C'+row+':C'+row).setDataValidation(rule);
    }
    
    // 添加目标账户（F列）DataValidation
    var range = accountSheet.getRange(getLastRange(accountSheet, accountColumns, 3));
    var rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
    cashSheet.getRange('F'+row).setDataValidation(rule);
    
    // 根据 cell.getValue() 为转账来源（G、H列）添加不同DataValidation
    cashSheet.getRange('G'+row+':H'+row).clear();
    cashSheet.getRange('G'+row+':H'+row).clearDataValidations();
    if(cell.getValue() == '转账'){
      cashSheet.getRange('G'+row).setValue('<-');
      cashSheet.getRange('H'+row).setDataValidation(rule);
    }
    
    return;
  }
  
  if(cell.getColumn() == 1){
    Logger.log(cell.getRow());
  }
}