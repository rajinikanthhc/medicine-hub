function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Medicine Hub");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}