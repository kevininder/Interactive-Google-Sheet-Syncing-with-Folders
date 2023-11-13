/* Get contents on the Google Sheet you specify for mapping. */

function getfiles() {
    // clear all
    var ss = SpreadsheetApp.getActive();
    ss.getRange('A7:F7').activate();
    var currentCell = ss.getCurrentCell();
    ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
    currentCell.activateAsCurrentCell();
    ss.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  
    var dApp = DriveApp;
    var folderIter = dApp.getFoldersByName("Music Library");
    var folder = folderIter.next();
    var filesIter = folder.getFiles();
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var i = 7;
  
    while(filesIter.hasNext()){
      var file = filesIter.next();
      var filename = file.getName();
      var filelink = file.getUrl();
      filename = filename.replace('.wav', '');
      filename = filename.replace('.mp3', '');
      ss.getRange(i, 1).setValue(filename);
      ss.getRange('G7').setValue(filelink);
      ss.getRange(i, 1).activate();
      ss.getRange(i, 1).splitTextToColumns('_');
      name = ss.getRange(i, 1).getValue();
      
      // insert links
      ss.getRange(i, 1).activate();
      ss.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
      .setText(name)
      .setLinkUrl(ss.getRange('G7').getValue())
      .build());
      i++;
    };
  
      ss.getRange('G7').setValue("");
  
      // sort
      ss.getRange('A7:F7').activate();
      var currentCell = ss.getCurrentCell();
      ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      ss.getActiveRange().sort({column: 1, ascending: true});
  
      // reset center
      ss.getRange('C7:E7').activate();
      var currentCell = ss.getCurrentCell();
      ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      ss.getActiveRangeList().setHorizontalAlignment('center');
  
      // reset color & fontline
      ss.getRange('B7:F7').activate();
      var currentCell = ss.getCurrentCell();
      ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      ss.getActiveRangeList().setFontColor(null)
      .setFontLine('none');
  
      // reset A7:B7
      ss.getRange('A7:B7').activate();
      var currentCell = ss.getCurrentCell();
      ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      ss.getActiveRangeList().setHorizontalAlignment('left')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  
      // reset F7
      ss.getRange('F7').activate();
      var currentCell = ss.getCurrentCell();
      ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      ss.getActiveRangeList().setHorizontalAlignment('left')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
  