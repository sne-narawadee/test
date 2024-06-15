
 /** Auto Refresh  from sheet pattern **/
 function doGet() {
  /**  สร้าง HTML สำหรับเว็บเพจ **/
  var html = '<html><head><style>';
  //html += '.redText { color: red; }'; // เพิ่ม CSS เพื่อกำหนดสีแดง
  // html += '.redText { color: red; font-size: 12px; }'; // เพิ่ม font-size: 12px; เพื่อกำหนดขนาดตัวอักษร 12px
  /** เพิ่ม font-size: 12px; เพื่อกำหนดขนาดตัวอักษร 12px **/
  html += 'table, th, td { border: 1px solid #d3d3d3; border-collapse: collapse; padding: 5px; font-size: 12px; }';

  /** เพิ่ม font-size: 12px; เพื่อกำหนดขนาดตัวอักษร 12px **/
  html += 'table { margin: auto; width: 100%; max-width: 800px; font-size: 12px; }'; 
  html += '@media screen and (max-width: 600px) { table { width: 100%; } }';

  /**  เพิ่ม font-size: 12px; เพื่อกำหนดขนาดตัวอักษร 12px **/
  html += 'tr:first-child { background-color: navy; color: white; font-size: 12px; }'; 
  html += '</style>';

  /**  JavaScript to Refresh Data **/
  html += '<script>';
  html += 'function updateData() {';
  html += 'google.script.run.withSuccessHandler(function(data) {';
  html += 'document.getElementById("data-table").innerHTML = data;';
  html += '}).getData();';
  html += '}';
  html += 'setInterval(updateData, 5000);'; // refresh ทุก 5 วินาที
  html += '</script>';
  html += '</head><body>';
  
  /** Create Table HTML by Data is not null **/ 
  html += '<table id="data-table">';
  html += '</table></body></html>';
  
  return HtmlService.createHtmlOutput(html);
}
////////////////////////////////////////////////////////

/**[Required sheet name] create function getData() for pull data from sheet pattern **/
function getData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RptDailyChkPnt"); 
  var data = sheet.getDataRange().getValues();
  
  /**  ตรวจสอบแถวที่มีข้อมูลว่างและคอลัมน์ที่มีข้อมูลว่าง **/ 
  var rowsWithValues = getRowsWithValues(data);
  var columnsWithValues = getColumnsWithValues(data);

  /** สร้าง HTML สำหรับตารางข้อมูล **/
  var html = '';
  
  /** [Required] เพิ่มแถวหัวตาราง "Header Name" **/
  var currentDate = new Date();
  var day = currentDate.getDate(); //number of Day
  var month = currentDate.getMonth() + 1; // เพิ่ม 1 เนื่องจาก getMonth() นับเดือนเริ่มจาก 0
  var year = currentDate.getFullYear();
  html += '<tr><th colspan="' + columnsWithValues.length + '" style="background-color: navy; color: white;"> QR Check Point Daily Report: '+day.toString().padStart(2,'0')+' / ' + month.toString().padStart(2, '0') + ' / ' + year + ' </th></tr>';

  /**  เพิ่มแถวหัวตารางของชื่อคอลัมน์ **/
  html += '<tr style="background-color: navy; color: white;">';
  for (var j = 0; j < columnsWithValues.length; j++) {
    var columnHeader = data[0][columnsWithValues[j]];
    html += '<th>' + columnHeader + '</th>';
  }
  html += '</tr>';

  /** เริ่มที่แถวที่ 1 เนื่องจากแถวที่ 0 เป็นหัวตาราง **/
  for (var i = 1; i < rowsWithValues.length; i++) {
    html += '<tr>';
    for (var j = 0; j < columnsWithValues.length; j++) {
      var cellData = data[rowsWithValues[i]][columnsWithValues[j]];
      /**  [Required] เช็คว่าเป็นคอลัมน์ A หรือไม่  Check Column index for Date format "yyyy/MM/dd" **/
      if (j === 0) { 
        html += '<td>' + formatDate(cellData) + '</td>'; // [Required] Function formatDate() 
      }  else {
        html += '<td>' + cellData + '</td>';
      }
    }
    html += '</tr>';
  }
      /**  เพิ่มข้อความลงในตาราง **/
      html += '<tr><td colspan="' + (columnsWithValues.length) + '"><center><span style="color: red;">©Copyright 2024:</span> <a href="https://www.narawadee-sne.com" target="_blank">Security Narawadee Express</a>';
      html += '</center></td></tr>';

  return html;
}

/////////////////////////////////////////////////////

/**  Date Format: "yyyy/MM/dd" **/
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

/** ฟังก์ชันเพื่อหาแถวที่มีข้อมูลทั้งหมด **/
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

/**  ฟังก์ชันเพื่อหาคอลัมน์ที่มีข้อมูลทั้งหมด **/
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
