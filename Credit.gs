// 还款计算器
function repaymentCalculator(){
  var calSheet = findSheet('还款计算器');
  
  if(calSheet == undefined){
    var calSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('还款计算器');
    calSheet.appendRow(['本金','每月计划新增','每月计划还款'])
    calSheet.appendRow(['0','0','0'])
    calSheet.appendRow(['','循环信用时长','每月实际还款','利息总计'])
    calSheet.appendRow(['','=countif(A12:A35, ">0")','=C2','=SUM(B12:B35)','=if(C2<A2/10, "最低还款必须大于10%","")'])
    calSheet.appendRow(['','分期数（月）','单期金额','单期手续费'])
    calSheet.appendRow(['','=iferror(OFFSET(C12,Match(CEILING(A2/(C2-B2),1),C12:C18,-1)-1,0),0)','=iferror(A2/B6,0)','=iferror(A2*VLOOKUP(B6,C12:D18,2,false),0)','=if(B6=0,"未达到最低分期金额","")'])
    calSheet.appendRow(['','','每月实际应付','费用总计'])
    calSheet.appendRow(['','','=B2+C6+D6','=D6*B6'])
    calSheet.appendRow([' ']);
    calSheet.appendRow(['参见：'])
    calSheet.appendRow(['循环利息','','分期手续费'])
    
    var next24months = getNextMonthDayNumber(24);
    calSheet.appendRow(['=A2+B2','=A12*0.0005*'+next24months[0]])
    var lastRow = 12;
    for(var i=1; i<next24months.length; i++){
      calSheet.appendRow(['=if(A'+lastRow+'-C$2 > 0, A'+lastRow+'-C$2+B$2, 0)','=A'+(lastRow+1)+'*0.0005*'+next24months[i]])
      lastRow++;
    }
    
    setColumnValues(calSheet, 12, 3, ['24','18','12','10','6','3','2']);
    setColumnValues(calSheet, 12, 4, ['0.68%','0.68%','0.66%','0.70%','0.75%','0.90%','1.00%']);
  }
  
  calSheet.getRange('A:B').setNumberFormat('¥0.00');
  calSheet.getRange('A1:D8').setNumberFormat('¥0.00');
  calSheet.getRange('B3:B6').setNumberFormat('0');
  
  for(var i=1; i<12; i+=2){
    if(i == 9) {
      continue;
    }
    calSheet.getRange('A'+i+':D'+i).setBackground('#d9d9d9');
  }
  
  calSheet.getRange('C5:D6').setFontColor('#999999');
  calSheet.getRange('E:E').setFontColor('#ff0000');
  
  calSheet.activate();
}

// 刷新当前时间接下来的月天数（用于更新利息）
function refreshRepaymentCalculator(){
  
  var calSheet = findSheet('还款计算器');
  if(calSheet == undefined){
    return;
  }
  
  var next24months = getNextMonthDayNumber(24);
  setColumnValues(calSheet, 12, 2, mergeValue(fill('=A{}*0.0005*',24,12),next24months));
}

// 获取接下来N个月的天数
function getNextMonthDayNumber(n){
  
  var result = new Array();
  
  var now = new Date();
  var year = now.getYear();
  var month = now.getMonth();
  for(var i=1; i<= n; i++){
    var date = new Date(year, month+i, 0);
    result.push(date.getDate())
  }
  
  return result;
}