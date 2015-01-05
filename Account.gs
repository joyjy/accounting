var accountCategoryColumns = ['A']
var accountColumns = ['F'];
var inColumns = ['J','K'];
var outColumns = ['L','M','N'];

// 创建帐户表
function createAccountSheet(){
  if(findSheet('账户')){
    return;
  }
  
  var accountSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('账户');
  
  // 建立账户分类
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
  newAccountValidation(accountSheet,3);
  accountSheet.getRange('B:B').setNumberFormat('¥0.00');
  accountSheet.getRange('C:C').setNumberFormat('0.00%');
  
  // 帐户表
  accountSheet.getRange(1, 5).setValue('账户');
  setRowValues(accountSheet, 2, 5, ['类型','名称','余额','分布']); 
  
  accountSheet.getRange('G:G').setNumberFormat('¥0.00');
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

// 新建账户
function newAccount(accountSheet, cell){
  Logger.log('newAccount: '+cell.getRow()+","+cell.getColumn())
  var row = cell.getRow();
  
  if(cell.getColumn() == 6){
    // 如果类型已选，不更新行
    if(accountSheet.getRange(row, 5).getValue() != ''){
      return;
    }
    newAccountValidation(accountSheet, row);
  }
}

// 自动填充账户前后列
function newAccountValidation(accountSheet, row){
  var cell = accountSheet.getRange(getLastRange(accountSheet, accountCategoryColumns, 3));
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(cell).build();
  accountSheet.getRange('E'+row).setDataValidation(rule);
  if(findSheet('流水')){
    var f="=sumif('流水'!F:F,F"+row+",'流水'!D:D)-sumif('流水'!H:H,F"+row+",'流水'!D:D)";
    accountSheet.getRange('G'+row).setValue(f);
  }
  accountSheet.getRange('H'+row).setValue('=iferror(G'+row+'/sumif(E:E,E'+row+',G:G),0)')
}