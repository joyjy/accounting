// 如果没有预算表则新建
function newBudget(){
  var budgetSheet = findSheet('预算');
  
  if(budgetSheet == undefined){
    var budgetSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('预算');
  }
  
  var lastRow = budgetSheet.getLastRow();
  appendBudget(budgetSheet, lastRow == 0 ? 1:lastRow+2);
  if(lastRow == 0){
    budgetSheet.setColumnWidth(1, 12);
    budgetSheet.setColumnWidth(2, 80);
    budgetSheet.setColumnWidth(3, 80);
    budgetSheet.setColumnWidth(4, 80);
    budgetSheet.setColumnWidth(5, 80);
    budgetSheet.setColumnWidth(6, 80);
    budgetSheet.getRange('D:F').setNumberFormat('¥0.00');
  }
}
  
var inSumif;
var outSumif;

// 追加预算
function appendBudget(sheet, row){
  sheet.activate();
  
  var month = new Date().getMonth();
  var startDate = formatDate(new Date(2015, month+1, 1));
  var endDate = formatDate(new Date(2015, month+2, 0));
  
  setRowValues(sheet, row, 1, ['#',startDate,endDate,'预算','实际','差额']);
  inSumif = "=sumifs('流水'!D:D,'流水'!B:B,\"收入\",'流水'!A:A,\">="+startDate+"\",'流水'!A:A,\"<="+endDate+"\")"
  outSumif = "=-sumifs('流水'!D:D,'流水'!B:B,\"支出\",'流水'!C:C,C{},'流水'!A:A,\">="+startDate+"\",'流水'!A:A,\"<="+endDate+"\")"
  
  var index = row+1;
  setColumnValues(sheet, index, 1, ['','-']);
  sheet.getRange(index, 4).setValue(0);
  sheet.getRange(index, 5).setValue(inSumif);
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
  
  setRowValues(sheet, index, 1, ["'=",'可支配收入']);
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
  
  setRowValues(sheet, index, 1, ['-','总计','','=sum(D'+startRow+':D'+endRow+')','=sum(E'+startRow+':E'+endRow+')','=D'+index+'-E'+index]);
  sheet.getRange('B'+startRow+':F'+index).setBackground('#F4CCCC');
  
  index++;
  setRowValues(sheet, index, 1, ["'=",'结余','','=D'+(startRow-1)+'-D'+(index-1),'=E'+(startRow-1)+'-E'+(index-1),'=-(D'+index+'-E'+index+')']);
  sheet.getRange('B'+index+':F'+index).setBackground('#B6D7A8');
}

function outBudget(accountSheet, defColumns, budgetSheet, index){
  
  var lastRow = getLastRow(accountSheet, defColumns, 2);
  copy(accountSheet, getLastRange(accountSheet, defColumns, 3, lastRow), budgetSheet, index, 3);
  
  var count = lastRow-2;
  setColumnValues(budgetSheet, index, 4, fill(0, count))
  setColumnValues(budgetSheet, index, 5, fill(outSumif, count, index));
  setColumnValues(budgetSheet, index, 6, fill("=D{}-E{}", count, index));
  return index + count;
}

function updateBudget(sheet, cell){
      
  var row = cell.getRow();
  
  if(sheet.getRange(row,1).getValue() == '#' && cell.getColumn() == 3){
    var startDate = formatDate(sheet.getRange(row, 2).getValue());
    var endDate = formatDate(cell.getValue());

    inSumif = "=sumifs('流水'!D:D,'流水'!B:B,\"收入\",'流水'!A:A,\">="+startDate+"\",'流水'!A:A,\"<="+endDate+"\")"
    outSumif = "=-sumifs('流水'!D:D,'流水'!B:B,\"支出\",'流水'!C:C,C{},'流水'!A:A,\">="+startDate+"\",'流水'!A:A,\"<="+endDate+"\")"
    
    for(var i=row+1; true; i++){
      var range = sheet.getRange(i, 5);
      if(range.getValue() === ''){
        break;
      }
      
      if(range.getBackground() == '#b6d7a8'){
        range.setValue(inSumif);
      } else if(range.getBackground() == '#f4cccc') {
        range.setValue(outSumif.replace('{}', i));
      }
    }
  }
}