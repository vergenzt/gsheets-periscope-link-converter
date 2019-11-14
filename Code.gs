function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
  .addItem('Convert links in selection', 'convertSelection')
  .addItem('Convert links in current sheet', 'convertSheet')
  .addItem('Convert links in spreadsheet', 'convertSpreadsheet')
  .addToUi();
}

function convertSelection() {
  convert(SpreadsheetApp.getSelection().getActiveRangeList().getRanges());
}

function convertSheet() {
  convert([SpreadsheetApp.getActiveSheet().getDataRange()]);
}

function convertSpreadsheet() {
  convert(SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function (sheet) { return sheet.getDataRange() }));
}

function convert(ranges) {
  for(var i=0; i<ranges.length; i++) {
    var range = ranges[i];
    var values = range.getValues();
    var numReplacements = 0;
    
    for (var j=0; j<range.getHeight(); j++) {
      for (var k=0; k<range.getWidth(); k++) {
        var match = values[j][k].match(/^\[\s*([^\]]+)\]\(([^\)]+)\)/);
        if (match) {
          var label = match[1];
          var url = match[2];
          range.getCell(j+1, k+1).setFormula("=hyperlink(\"" + url + "\", \"" + label + "\")");
          numReplacements++;
        }
      }
    }
    
    SpreadsheetApp.getUi().alert(numReplacements > 0 ? "Replaced " + numReplacements + " Periscope hyperlink(s) with GSheet equivalent." : "No Periscope hyperlinks found.");
  }
}
