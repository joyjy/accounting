// 如果没有预算表则新建
function newBudget(){
  var budgetSheet = findSheet('预算');
  
  if(budgetSheet == undefined){
    var budgetSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('预算');
  }
  
  var lastRow = budgetSheet.getLastRow();
  appendBudget(budgetSheet, lastRow == 0 ? 1:lastRow+2);
  if(lastRow == 0){
    budgetSheet.setColumnWidth(1, 20);
    budgetSheet.setColumnWidth(2, 80);
    budgetSheet.setColumnWidth(3, 80);
    budgetSheet.setColumnWidth(4, 80);
    budgetSheet.setColumnWidth(5, 80);
    budgetSheet.setColumnWidth(6, 80);
    budgetSheet.getRange('D:F').setNumberFormat('¥0.00');
  }
}
  
var sumif = "=sumif('流水'!B:B,{},'流水'!D:D)";

// 追加预算
function appendBudget(sheet, row){
  sheet.activate();
  
  var time = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "MM月");
  
  setRowValues(sheet, row, 1, ['#',time,'','预算','实际','差额']);
  
  var index = row+1;
  setColumnValues(sheet, index, 1, ['','-']);
  sheet.getRange(index, 4).setValue(0);
  sheet.getRange(index, 5).setValue(sumif.replace('{}','"收入"'));
  sheet.getRange(index, 6).setValue('=-(D'+index+'-E'+index+')');
  
  var accountSheet = findSheet('账户');
  var fixRow = getLastRow(accountSheet, ['L'], 2);
  if(fixRow > 2){
    setColumnValues(sheet, index, 2, ['收入']);
    sheet.getRange('B'+index+':F'+index).setBackground('#B6D7A8');
    index++;
    
    setColumnValues(sheet, index, 2, ['固定支出']);
    var end = outBudget(accountSheet, ['L'], sheet, index);
    
    sheet.getRange('B'+index+':F'+(end-1)).setBackground('#F4CCCC');
    
    var f = 'D'+index+':'+'D'+(end-1);
    index = end;
  }
  
  setColumnValues(sheet, index, 2, ['可支配收入']);
  if(f){
    sheet.getRange(index, 4).setValue('=D2-sum('+f+')');
    sheet.getRange(index, 5).setValue('=D2-sum('+f+')');
    sheet.getRange(index, 6).setValue('=-(D'+index+'-E'+index+')');
  }
  
  sheet.getRange('B'+index+':F'+index).setBackground('#B6D7A8');
  index++;
  
  var startRow = index;
  
  setColumnValues(sheet, index, 2, ['常规支出']);
  index = outBudget(accountSheet, ['M'], sheet, index);
  
  setColumnValues(sheet, index, 2, ['其他支出']);
  index = outBudget(accountSheet, ['N'], sheet, index);
  
  var endRow = index-1;
  
  setRowValues(sheet, index, 2, ['总计','','=sum(D'+startRow+':D'+endRow+')','=sum(E'+startRow+':E'+endRow+')','=D'+index+'-E'+index]);
  sheet.getRange('B'+startRow+':F'+index).setBackground('#F4CCCC');
  
  index++;
  setRowValues(sheet, index, 2, ['结余','','=D'+(startRow-1)+'-D'+(index-1),'=E'+(startRow-1)+'-E'+(index-1),'=-(D'+index+'-E'+index+')']);
  sheet.getRange('B'+index+':F'+index).setBackground('#B6D7A8');
  
  //var monthSpan = '1/1/2015 - 1/31/2015'
  //if(cell != undefined){
  //  var lastMonthSpan = sheet.getRange(cell.getRow(), cell.getColumn()+1).getValue();
  //  monthSpan = lastMonthSpan;
  //}
  
  //var nextMonth = new Date().getMonth()+1;
  //if(nextMonth == 12) nextMonth = 0;
  //nextMonth+=1;
  
  //var cell = sheet.getRange(row, 1);
  //cell.setValue(nextMonth+'月');
}

function outBudget(accountSheet, defColumns, budgetSheet, index){
  
  var lastRow = getLastRow(accountSheet, defColumns, 2);
  copy(accountSheet, getLastRange(accountSheet, defColumns, 3, lastRow), budgetSheet, index, 3);
  
  var count = lastRow-2;
  setColumnValues(budgetSheet, index, 4, fill(0, count))
  setColumnValues(budgetSheet, index, 5, fill(sumif.replace('{}','C{}'), count, index));
  setColumnValues(budgetSheet, index, 6, fill("=D{}-E{}", count, index));
  return index + count;
}