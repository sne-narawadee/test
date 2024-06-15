function doGet() {
  var data = {
    message: "Hello, world!",
    date: new Date()
  };
  return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
}
