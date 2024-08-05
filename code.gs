function doGet(e) {  
  return HtmlService.createTemplateFromFile("index").evaluate()
  .setTitle("WebApp: Search By Password")
  .addMetaTag('viewport', 'width=device-width, initial-scale=1')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* PROCESS FORM */
function processForm(formObject){ 
  var concat = formObject.searchtext.toString().toLowerCase()+formObject.searchtext2;
  var result = "";
  if(concat){//Execute if form passes search text
      result = search(concat);
  }
  return result;
}

//SEARCH FOR MATCHED CONTENTS ;
function search(searchtext){
  var range = SpreadsheetApp.getActive().getSheetByName('Data').getDataRange();
  var data = range.getValues();
  var ar = [];
  
  data.forEach(function(f) {
    if (~[f[0].toString().toLowerCase()+f[1]].indexOf(searchtext)) {
     ar.push([ f[2], f[3]]);
    }
  });
                                           
  return ar;
};
