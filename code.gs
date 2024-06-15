function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('QR Check Point Daily Report')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RptDailyChkPnt"); 
  var data = sheet.getDataRange().getValues();
  
  var rowsWithValues = getRowsWithValues(data);
  var columnsWithValues = getColumnsWithValues(data);

  var html = '';
  var currentDate = new Date();
  var day = currentDate.getDate();
  var month = currentDate.getMonth() + 1;
  var year = currentDate.getFullYear();
  html += '<table style="border: 1px solid #d3d3d3; border-collapse: collapse; width: 100%; max-width: 800px; margin: auto; font-size: 12px;">';
  html += '<tr><th colspan="' + columnsWithValues.length + '" style="background-color: navy; color: white; padding: 5px;"> QR Check Point Daily Report: '+day.toString().padStart(2,'0')+' / ' + month.toString().padStart(2, '0') + ' / ' + year + ' </th></tr>';
  
  html += '<tr style="background-color: navy; color: white;">';
  for (var j = 0; j < columnsWithValues.length; j++) {
    var columnHeader = data[0][columnsWithValues[j]];
    html += '<th style="border: 1px solid #d3d3d3; padding: 5px;">' + columnHeader + '</th>';
  }
  html += '</tr>';

  for (var i = 1; i < rowsWithValues.length; i++) {
    html += '<tr>';
    for (var j = 0; j < columnsWithValues.length; j++) {
      var cellData = data[rowsWithValues[i]][columnsWithValues[j]];
      if (j === 0) {
        html += '<td style="border: 1px solid #d3d3d3; padding: 5px;">' + formatDate(cellData) + '</td>';
      }  else {
        html += '<td style="border: 1px solid #d3d3d3; padding: 5px;">' + cellData + '</td>';
      }
    }
    html += '</tr>';
  }

  html += '<tr><td colspan="' + (columnsWithValues.length) + '" style="text-align: center; padding: 5px;"><span style="color: red;">© Copyright 2024:</span> <a href="https://www.narawadee-sne.com" target="_blank">Security Narawadee Express</a>';
  html += '</td></tr></table>';

  return html;
}

function formatDate(date) {
  if (date instanceof Date) {
    var year = date.getFullYear();
    var month = ('0' + (date.getMonth() + 1)).slice(-2);
    var day = ('0' + date.getDate()).slice(-2);
    return year + '/' + month + '/' + day;
  } else {
    return date;
  }
}

function getRowsWithValues(data) {
  var rowsWithValues = [];
  for (var i = 0; i < data.length; i++) {
    var rowHasValues = false;
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] !== "") {
        rowHasValues = true;
        break;
      }
    }
    if (rowHasValues) {
      rowsWithValues.push(i);
    }
  }
  return rowsWithValues;
}

function getColumnsWithValues(data) {
  var columnsWithValues = [];
  if (data.length > 0) {
    for (var j = 0; j < data[0].length; j++) {
      var columnHasValues = false;
      for (var i = 0; i < data.length; i++) {
        if (data[i][j] !== "") {
          columnHasValues = true;
          break;
        }
      }
      if (columnHasValues) {
        columnsWithValues.push(j);
      }
    }
  }
  return columnsWithValues;
}
