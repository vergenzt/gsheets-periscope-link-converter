function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
  .addItem('Convert links in selection', 'convertSelection')
  .addItem('Convert links in current sheet', 'convertSheet')
  .addItem('Convert links in spreadsheet', 'convertSpreadsheet')
  .addToUi();
}

function onInstall(e) {
  onOpen(e);
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
  var replacements = [];
  for(var i=0; i<ranges.length; i++) {
    var range = ranges[i];
    var values = range.getValues();
    
    for (var j=0; j<range.getHeight(); j++) {
      for (var k=0; k<range.getWidth(); k++) {
        var match = values[j][k].match(/^\[\s*([^\]]+)\]\(([^\)]+)\)/);
        if (match) {
          var targetCell = range.getCell(j+1, k+1);
          var label = match[1];
          var url = match[2];
          var newValue = "=hyperlink(\"" + url + "\", \"" + label + "\")";
          var replacement = [targetCell, newValue];
          replacements.push(replacement);
        }
      }
    }
  }
  
  // batch updates so the undo system treats it as a single change
  for(var i=0; i<replacements.length; i++) {
    var targetCell = replacements[i][0];
    var newValue = replacements[i][1];
    targetCell.setFormula(newValue);
  }
  SpreadsheetApp.flush();
  
  var numReplacements = replacements.length;
  SpreadsheetApp.getUi().alert(numReplacements > 0 ? "Replaced " + numReplacements + " Periscope hyperlink(s) with GSheets equivalent." : "No Periscope hyperlinks found.");
}
